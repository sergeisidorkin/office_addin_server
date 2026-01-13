# C:\DocOpsRunner\DocOps-Agent.ps1
param(
  [string]$Base  = $env:DOCOPS_BASE,
  [string]$Bases = $env:DOCOPS_BASES,
  [string]$Agent = $env:DOCOPS_AGENT,
  [int]   $Port  = 17603,
  [int]   $NoJobWaitSec = 45,
  [int]   $IdleMinSec   = 3,
  [int]   $IdleMaxSec   = 20
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Off
if([string]::IsNullOrWhiteSpace($Agent)){ $Agent = "addin-auto" }

$psExe   = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$dir     = 'C:\DocOpsRunner'
$runPS   = Join-Path $dir 'Interactive-RunWSH.ps1'
$finPS   = Join-Path $dir 'Word-Finalize-WSH.ps1'
$logsDir = Join-Path $dir 'logs'
if(-not (Test-Path $logsDir)){ New-Item -ItemType Directory -Path $logsDir | Out-Null }
$logFile = Join-Path $logsDir ("agent-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
function Log([string]$m){ "{0:yyyy-MM-dd HH:mm:ssK}|[AGENT] {1}" -f (Get-Date),$m | Tee-Object -FilePath $logFile -Append | Out-Null }

try { . 'C:\DocOpsRunner\Lib-ServerLog.ps1' } catch {}
try{ Rotate-DocOpsLogs -Days 3 }catch{}

function Parse-Bases([string]$bases,[string]$fallbackBase){
  $list = @()
  if(-not [string]::IsNullOrWhiteSpace($bases)){
    $list = $bases -split '[;,\s]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
  }
  if(-not $list -and -not [string]::IsNullOrWhiteSpace($fallbackBase)){ $list = @($fallbackBase) }
  if(-not $list){ $list = @("https://snappishly-primed-blackfish.cloudpub.ru") }
  return $list
}

function Stop-OldFinalizers(){
  try{
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -EA 0 |
      Where-Object { $_.CommandLine -match 'Word-Finalize-WSH\.ps1' } |
      ForEach-Object { try{ Stop-Process -Id $_.ProcessId -Force -EA 0 }catch{} }
  }catch{}
}

function Prefetch-Trace([string]$base,[string]$agent){
  $hdr = @{ 'ngrok-skip-browser-warning'='1'; 'Content-Type'='application/json' }
  try{
    $url = "$($base.TrimEnd('/'))/api/agents/$agent/pull"
    $r = Invoke-RestMethod -Method Post -Uri $url -Headers $hdr -Body '{}' -TimeoutSec 15 -EA Stop
    if($r -and $r.ok -and $r.job){
      $tid = ""; $jid = ""
      try{ $tid = [string]$r.job.traceId }catch{}
      try{ $jid = [string]$r.job.id }catch{}
      if([string]::IsNullOrWhiteSpace($tid)){ $tid = $jid }  #  ключевая строка: fallback на jobId
      Log ("Prefetch[{0}]: jobId={1} trace={2}" -f $base,$jid,$tid)
      try{ Send-ServerLog -Level debug -Phase 'agent' -Event 'prefetch.ok' -Message ("jobId={0} trace={1}" -f $jid,$tid) -TraceId $tid -JobId $jid }catch{}
      return @{ trace=$tid; jobId=$jid }
    }
  }catch{
    Log ("Prefetch[{0}]: ERR -> {1}" -f $base, $_.Exception.Message)
    try{ Send-ServerLog -Level warn -Phase 'agent' -Event 'prefetch.err' -Message ("{0}: {1}" -f $base, $_.Exception.Message) }catch{}
  }
  return @{ trace=""; jobId="" }
}

function Start-Finalizer([string]$traceId){
  $args = "-NoProfile -ExecutionPolicy Bypass -File `"$finPS`" -TimeoutSec 300 -WaitForMarkerSec 0 -ListenLocalHttpPort $Port -BgWaitMaxMs 3500"
  if(-not [string]::IsNullOrWhiteSpace($traceId)){ $args += (" -TraceId {0}" -f $traceId) }
  Log "Starting Finalizer: $args"
  try{
    $p = Start-Process -FilePath $psExe -ArgumentList $args -PassThru -WindowStyle Hidden
    try{ Send-ServerLog -Level info -Phase 'agent' -Event 'finalizer.start' -Message ("pid={0}" -f $p.Id) -TraceId $traceId }catch{}
    return $p
  } catch {
    Log ("ERR: Start finalizer -> " + $_.Exception.Message)
    try{ Send-ServerLog -Level error -Phase 'agent' -Event 'finalizer.start.err' -Message $_.Exception.Message -TraceId $traceId }catch{}
    return $null
  }
}

function Run-RunnerSync([string]$traceId,[string]$jobId,[string]$base){
  if([string]::IsNullOrWhiteSpace($traceId)){ $env:DOCOPS_TRACE=""; $env:TRACE="" } else { $env:DOCOPS_TRACE=$traceId; $env:TRACE=$traceId }
  $env:DOCOPS_BASE = $base
  try{ Send-ServerLog -Level info -Phase 'agent' -Event 'runner.start' -Message ("runner sync (base={0})" -f $base) -TraceId $traceId -JobId $jobId }catch{}
  Log ("Running Runner (TraceId={0}, Base={1})..." -f $traceId, $base)
  try{
    Start-Process -FilePath $psExe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"$runPS`" -FinalizePort {0}" -f $Port) -Wait -WindowStyle Hidden
  }catch{
    Log ("Runner error: " + $_.Exception.Message)
    try{ Send-ServerLog -Level warn -Phase 'agent' -Event 'runner.err' -Message $_.Exception.Message -TraceId $traceId -JobId $jobId }catch{}
  }
  Log "Runner finished"
}

# --- single instance ---
$createdNew=$false
$mutexName = "Global\DocOpsAgent-$($env:USERNAME)"
$mutex=$null
try { $mutex = New-Object System.Threading.Mutex($false,$mutexName,[ref]$createdNew) } catch { Log ("ERR: Mutex create -> "+$_.Exception.Message) }
if($null -eq $mutex){ Log "ERR: Mutex is null -> exit"; try{ Send-ServerLog -Level error -Phase 'agent' -Event 'mutex.null' -Message 'mutex is null' }catch{}; exit 1 }
if(-not $mutex.WaitOne(0)){ Log "Another instance already running -> exit"; try{ Send-ServerLog -Level info -Phase 'agent' -Event 'already.running' -Message 'another instance detected' }catch{}; exit 0 }

# --- boot ---
$basesList = Parse-Bases -bases $Bases -fallbackBase $Base
Log ("Boot. Bases=[{0}] Agent={1} Port={2}" -f ([string]::Join(' ', $basesList)),$Agent,$Port)
try{ Send-ServerLog -Level info -Phase 'agent' -Event 'boot' -Message ("Bases=[{0}] Agent={1} Port={2}" -f ([string]::Join(' ', $basesList)),$Agent,$Port) }catch{}

# --- main loop ---
$idle = $IdleMinSec
try{
  while($true){
    Log "=== TICK start ==="

    if(-not (Test-Path $runPS)){ Log "ERR: Runner not found: $runPS"; try{ Send-ServerLog -Level error -Phase 'agent' -Event 'runner.missing' -Message $runPS }catch{}; Start-Sleep -Seconds 5; continue }
    if(-not (Test-Path $finPS)){ Log "ERR: Finalizer not found: $finPS"; try{ Send-ServerLog -Level error -Phase 'agent' -Event 'finalizer.missing' -Message $finPS }catch{}; Start-Sleep -Seconds 5; continue }

    Stop-OldFinalizers

    # ищем джобу по очереди баз
    $pref = $null; $selectedBase = $null
    foreach($b in $basesList){
      $pref = Prefetch-Trace -base $b -agent $Agent
      if($pref.jobId){ $selectedBase = $b; break }
    }

    if(-not $selectedBase){
      Log "No jobs on all bases -> sleep"
      Start-Sleep -Seconds $IdleMinSec
      continue
    }

    # нормализуем trace и прокидываем окружение ДО старта дочерних процессов
    $trace = $pref.trace
    if([string]::IsNullOrWhiteSpace($trace)){ $trace = $pref.jobId }   #  тот же fallback на jobId
    $env:DOCOPS_BASE  = $selectedBase
    if([string]::IsNullOrWhiteSpace($trace)){ $env:DOCOPS_TRACE = ""; $env:TRACE = "" } else { $env:DOCOPS_TRACE = $trace; $env:TRACE = $trace }

    # стартуем финалайзер и раннер на найденной базе
    $finProc = Start-Finalizer -traceId $trace
    if($null -eq $finProc){ Log "ERR: Finalizer didn't start"; Start-Sleep -Seconds 3; continue }

    Run-RunnerSync -traceId $trace -jobId $pref.jobId -base $selectedBase

    # ждём /done от панели
    $deadline = (Get-Date).AddSeconds([Math]::Max(5,$NoJobWaitSec))
    $jobProcessed = $false
    try{
      while((Get-Date) -lt $deadline){
        if($finProc -and $finProc.HasExited){ $jobProcessed = $true; break }
        Start-Sleep -Milliseconds 300
      }
    }catch{ Log ("ERR: Wait finProc -> " + $_.Exception.Message) }

    if(-not $jobProcessed){
      Log "No /done received in ${NoJobWaitSec}s -> killing finalizer"
      try{ if($finProc -and -not $finProc.HasExited){ Stop-Process -Id $finProc.Id -Force -EA 0; Log "Finalizer killed (no-job timeout)" } }catch{}
      try{ Send-ServerLog -Level info -Phase 'agent' -Event 'idle.nojob' -Message ("no /done in {0}s" -f $NoJobWaitSec) -TraceId $trace -JobId $pref.jobId }catch{}
      $idle = [Math]::Min([int]([Math]::Ceiling($idle*1.6)), $IdleMaxSec)
      Log ("Sleep (idle/no-jobs): {0}s" -f $idle)
      Start-Sleep -Seconds $idle
      continue
    }

    Log "Finalizer exited -> job processed"
    try{ Send-ServerLog -Level info -Phase 'agent' -Event 'job.processed' -Message 'finalizer exited' -TraceId $trace -JobId $pref.jobId }catch{}
    $idle = $IdleMinSec
    Log ("Sleep (post-job): {0}s" -f $idle)
    Start-Sleep -Seconds $idle
  }
}
finally{
  try{ if($mutex){ $mutex.ReleaseMutex() } }catch{}
  if($mutex){ $mutex.Dispose() }
}
