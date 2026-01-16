# C:\DocOpsRunner\Lib-ServerLog.ps1
$ErrorActionPreference = "Stop"
Set-StrictMode -Off

function Get-LogsBase {
  $b = ""; try { $b = [string]$env:DOCOPS_BASE } catch {}
  if([string]::IsNullOrWhiteSpace($b)){ $b = "https://admiringly-conscious-remora.cloudpub.ru" }
  return $b.TrimEnd('/')
}

function Get-LogsToken {
  $t = ""; try { $t = [string]$env:DOCOPS_LOGS_TOKEN } catch {}
  if(-not [string]::IsNullOrWhiteSpace($t)){ return $t.Trim() }
  $path = 'C:\DocOpsRunner\serverlog.token'
  if(Test-Path $path){
    try { return (Get-Content $path -Raw).Trim() } catch {}
  }
  return ""
}

function Send-ServerLog {
  param(
    [Parameter(Mandatory=$true)][ValidateSet('debug','info','warn','error')][string]$Level,
    [Parameter(Mandatory=$true)][string]$Phase,
    [Parameter(Mandatory=$true)][string]$Event,
    [string]$Message = "",
    [string]$TraceId = "",
    [string]$JobId = "",
    [string]$DocUrl = "",
    [hashtable]$Data = @{}
  )
  $base   = Get-LogsBase
  $token  = Get-LogsToken
  if([string]::IsNullOrWhiteSpace($token)){
    return [pscustomobject]@{ ok=$false; ingested=0; err="no token" }
  }

  # traceId: из параметра -> из env -> новый guid
  if([string]::IsNullOrWhiteSpace($TraceId)){
    try { $TraceId = [string]$env:DOCOPS_TRACE } catch {}
  }
  if([string]::IsNullOrWhiteSpace($TraceId)){
    try { $TraceId = [guid]::NewGuid().ToString() } catch {}
  }

  $url = "$base/logs/api/logs/ingest"
  $hdr = @{
    'Content-Type'              = 'application/json; charset=utf-8'
    'ngrok-skip-browser-warning'= '1'
    'X-IMC-Logs-Token'          = $token
  }

  $payload = @{ level=$Level; phase=$Phase; event=$Event }
  if($Message){ $payload["message"] = $Message }
  if($TraceId){ $payload["traceId"] = $TraceId }
  if($JobId){   $payload["jobId"]   = $JobId }
  if($DocUrl){  $payload["docUrl"]  = $DocUrl }
  if($Data){    $payload["data"]    = $Data }

  $json = $payload | ConvertTo-Json -Depth 8 -Compress

  $dbg = [bool]($env:DOCOPS_LOGS_DEBUG -eq '1')
  $dbgFile = 'C:\DocOpsRunner\logs\ingest-debug.log'

  try{
    $resp = Invoke-RestMethod -Method Post -Uri $url -Headers $hdr -Body $json -TimeoutSec 10 -ErrorAction Stop
    if($dbg){ "[{0:s}] OK {1}" -f (Get-Date), ($resp | ConvertTo-Json -Compress) | Out-File $dbgFile -Append -Encoding UTF8 }
    return $resp
  } catch {
    $msg = $_.Exception.Message
    $code = ""; $text = ""
    try{
      $r = $_.Exception.Response
      if($r){
        $code = [string]$r.StatusCode.value__
        $sr = New-Object IO.StreamReader($r.GetResponseStream())
        $text = $sr.ReadToEnd()
      }
    }catch{}
    if($dbg){ "[{0:s}] ERR code={1} msg={2} body={3}" -f (Get-Date),$code,$msg,$text | Out-File $dbgFile -Append -Encoding UTF8 }
    return [pscustomobject]@{ ok=$false; ingested=0; err=("{0} {1}" -f $code,$msg); body=$text }
  }
}

# --- Rotate DocOps logs helper (keep N days) ---
function Rotate-DocOpsLogs {
  param([int]$Days = 3, [string]$Dir = 'C:\DocOpsRunner\logs')
  try {
    if (-not (Test-Path $Dir)) { return }
    $cutoff = (Get-Date).AddDays(-[Math]::Max(1,$Days))
    Get-ChildItem $Dir -File -Filter *.log -ErrorAction SilentlyContinue |      
      Where-Object { $_.LastWriteTime -lt $cutoff } |
      Remove-Item -Force -ErrorAction SilentlyContinue
  } catch {}
}