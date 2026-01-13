# C:\DocOpsRunner\Interactive-RunWSH.ps1
param(
  [int]$FinalizePort = 17603  # параметр принимаем для совместимости, не используем
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Off

# === конфиг ===
$Base  = $env:DOCOPS_BASE  ; if([string]::IsNullOrWhiteSpace($Base))  { $Base  = "https://snappishly-primed-blackfish.cloudpub.ru" }
$Agent = $env:DOCOPS_AGENT ; if([string]::IsNullOrWhiteSpace($Agent)) { $Agent = "addin-auto" }

# === лог в файл ===
$logDir = 'C:\DocOpsRunner\logs'
if(-not (Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir | Out-Null }
$log    = Join-Path $logDir ("interactive2-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
function Log([string]$m){ "{0:yyyy-MM-dd HH:mm:ssK}|[INTER2] {1}" -f (Get-Date),$m | Tee-Object -FilePath $log -Append | Out-Null }

# библиотека серверных логов (тихо, если что-то не так)
try { . 'C:\DocOpsRunner\Lib-ServerLog.ps1' } catch {}
try{ Rotate-DocOpsLogs -Days 3 }catch{}

try{
    $sid = (Get-Process -Id $PID).SessionId
    $envTrace = ""; try{ $envTrace = [string]$env:DOCOPS_TRACE }catch{}
    Log ("Boot. User={0} SessionId={1} Trace(env)={2}" -f $env:USERNAME,$sid,$envTrace)
    try{ Send-ServerLog -Level info -Phase 'runner' -Event 'boot' -Message ("User={0} Sid={1}" -f $env:USERNAME,$sid) -TraceId $envTrace }catch{}

    $wordExe = (Get-Command winword.exe -EA 0 | Select-Object -Expand Source -EA 0)
    if(-not $wordExe -or -not (Test-Path $wordExe)){
    $cand = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
    if(Test-Path $cand){ $wordExe = $cand } else { $wordExe = 'winword.exe' }
    }
    Log ("Using Word: {0}" -f $wordExe)

    # Тянем задание
    $hdr = @{ 'ngrok-skip-browser-warning'='1'; 'Content-Type'='application/json' }
    $pullUrl = "$Base/api/agents/$Agent/pull"
    Log ("Pull: {0}" -f $pullUrl)
    $pull = $null
    try{ $pull = Invoke-RestMethod -Method Post -Uri $pullUrl -Headers $hdr -Body '{}' -EA Stop }catch{ Log ("ERR: pull -> " + $_.Exception.Message) }

    if(-not $pull -or -not $pull.ok -or -not $pull.job){
    Log "No job -> exit"
    try{ Send-ServerLog -Level debug -Phase 'runner' -Event 'pull.none' -Message 'no job' -TraceId $envTrace }catch{}
    return
    }

    if(-not $pull -or -not $pull.ok -or -not $pull.job){
      Log "No job -> exit"
      try{ Send-ServerLog -Level debug -Phase 'runner' -Event 'pull.none' -Message 'no job' -TraceId $envTrace }catch{}
      return
    }

    # --- прокинем BASE в панель через локальный файл ---
    try{
      $cfgPath = 'C:\inetpub\wwwroot\addin\config.js'
      $cfgBody = "window.DOCOPS_BASE = " + ('"{0}"' -f $Base) + ";"
      Set-Content -Path $cfgPath -Value $cfgBody -Encoding UTF8
      Log ("config.js written: {0} -> {1}" -f $cfgPath, $Base)
      try{ Send-ServerLog -Level info -Phase 'runner' -Event 'finalizer.start' -Message 'config.js updated' -TraceId $tid }catch{}
    }catch{
      Log ("WARN: config.js write failed: " + $_.Exception.Message)
    }

    # traceId из pull  в env и в логи
    $tid = ""; try{ $tid = [string]$pull.job.traceId }catch{}
    if(-not [string]::IsNullOrWhiteSpace($tid)){
    $env:DOCOPS_TRACE = $tid
    Log ("Trace from pull: {0}" -f $tid)
    try{ Send-ServerLog -Level info -Phase 'runner' -Event 'trace.set' -Message 'trace from pull' -TraceId $tid -JobId ([string]$pull.job.id) }catch{}
    } else { $tid = $envTrace }

    $msLink = $pull.job.msLink
    $webUrl = $pull.job.webUrl
    if([string]::IsNullOrWhiteSpace($msLink) -and [string]::IsNullOrWhiteSpace($webUrl)){
    Log "WARN: msLink/webUrl missing -> exit"
    try{ Send-ServerLog -Level warn -Phase 'runner' -Event 'open.missing_url' -Message 'msLink/webUrl missing' -TraceId $tid -JobId ([string]$pull.job.id) }catch{}
    return
    }

    $openArg = if($msLink){ $msLink } else { "ms-word:ofe|u|$webUrl" }
    Log ("OpenArg: {0}" -f $openArg)
    try{ Send-ServerLog -Level info -Phase 'runner' -Event 'open' -Message 'open word' -TraceId $tid -JobId ([string]$pull.job.id) -DocUrl $webUrl -Data @{ arg=$openArg } }catch{}

    # Открываем Word
    $p = Start-Process -FilePath $openArg -PassThru -WindowStyle Normal
    Log ("WINWORD started. StubPID={0}" -f $p.Id)

    # Ждём HWND
    [uint32]$hwnd = 0
    $pidReal = 0; $deadline = (Get-Date).AddSeconds(35)
    do{
    $w = Get-Process WINWORD -EA 0 | Where-Object { $_.SessionId -eq $sid } | Sort-Object StartTime -Desc | Select -First 1
    if($w){ $pidReal = $w.Id; $h64 = [Int64]$w.MainWindowHandle; if($h64 -gt 0){ $hwnd = [uint32]$h64 } else { Start-Sleep -Milliseconds 250 } }
    else { Start-Sleep -Milliseconds 250 }
    } while($hwnd -eq 0 -and (Get-Date) -lt $deadline)

    if($pidReal){ Log ("Guard(MWH): PID={0} HWND=0x{1:X}" -f $pidReal,$hwnd) }
    if($hwnd -eq 0){ Log "WARN: HWND=0 after wait (helper anyway)"; try{ Send-ServerLog -Level warn -Phase 'runner' -Event 'hwnd.zero' -Message 'no main window yet' -TraceId $tid -JobId ([string]$pull.job.id) }catch{} }

    Log "Calling Word-SearchInvoke-WSH"
    try{ Send-ServerLog -Level debug -Phase 'runner' -Event 'helper.call' -Message 'Word-SearchInvoke-WSH Show Task Pane' -TraceId $tid -JobId ([string]$pull.job.id) }catch{}
    & 'C:\DocOpsRunner\Word-SearchInvoke-WSH.ps1' -Hwnd ([uint32]$hwnd) -Query 'Show Task Pane'
    Log "Helper finished"

    # Раннер НЕ вызывает финализацию; её вызывает панель/WSH-хелпер
    try{ Send-ServerLog -Level info -Phase 'runner' -Event 'no_auto_finalize' -Message 'runner leaves finalize to add-in' -TraceId $tid -JobId ([string]$pull.job.id) }catch{}
}
catch{
  Log ("ERR: " + $_.Exception.Message)
  try{ Send-ServerLog -Level error -Phase 'runner' -Event 'exception' -Message $_.Exception.Message -TraceId $tid -JobId ([string]$pull.job.id) }catch{}
}
finally{
  [Environment]::Exit(0)
}
