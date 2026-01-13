# C:\DocOpsRunner\Word-SearchInvoke-WSH.ps1
param(
  [object]$Hwnd = 0,
  [string]$Query = "Show Task Pane"
)

# ========== logging ==========
$logDir = 'C:\DocOpsRunner\logs'
if(-not (Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$log = Join-Path $logDir ("search-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
function W([string]$m){ "[{0:yyyy-MM-dd HH:mm:ssK}]|[SEARCH] {1}" -f (Get-Date), $m | Out-File -FilePath $log -Append -Encoding UTF8 }
W ("Boot. Query='{0}'" -f $Query)

# ========== HWND normalize ==========
[System.IntPtr]$HWND_PTR = [System.IntPtr]::Zero
[uint32]$HWND32 = 0
if($PSBoundParameters.ContainsKey('Hwnd') -and $null -ne $Hwnd){
  if($Hwnd -is [System.IntPtr]){ $HWND_PTR = [System.IntPtr]$Hwnd; $HWND32 = [uint32]([int64]$Hwnd) }
  else { $HWND32 = [uint32]$Hwnd; $HWND_PTR = [System.IntPtr]::new([int64]$HWND32) }
}else{
  try{
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    if($app -and $app.Hwnd -gt 0){
      $HWND32  = [uint32]$app.Hwnd
      $HWND_PTR = [System.IntPtr]::new([int64]$HWND32)
    }
  }catch{}
}
W ("HWND_PTR=0x{0:X} HWND32={1}" -f $HWND_PTR.ToInt64(), $HWND32)
if($HWND_PTR -eq [IntPtr]::Zero){ W "ERR: HWND=0"; exit 1 }

# ========== UIAutomation ==========
try{ Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes }catch{
  try{
    [void][System.Reflection.Assembly]::LoadWithPartialName('UIAutomationClient')
    [void][System.Reflection.Assembly]::LoadWithPartialName('UIAutomationTypes')
  }catch{ W "ERR: UIAutomation assemblies not available"; exit 2 }
}
$root = [System.Windows.Automation.AutomationElement]::FromHandle($HWND_PTR)
if($null -eq $root){ W "ERR: UIA root from HWND failed"; exit 3 }
$wordPid = $root.Current.ProcessId

# ========== WinForms SendKeys (предпочтительно) ==========
$formsOK = $false
try{ Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop; $formsOK=$true }catch{ W "WARN: System.Windows.Forms not available (will fallback to WScript.SendKeys)" }

# ========== Native helpers (SendInput, SC_KEYMENU, Focus) ==========
$nativeCode = @"
using System;
using System.Runtime.InteropServices;
namespace IMC {
  public static class Kbd {
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT { public uint type; public InputUnion U; }
    [StructLayout(LayoutKind.Explicit)]
    public struct InputUnion { [FieldOffset(0)] public KEYBDINPUT ki; }
    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
    [DllImport("user32.dll")] public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    public const uint INPUT_KEYBOARD    = 1;
    public const uint KEYEVENTF_KEYUP   = 0x0002;
    public const uint WM_SYSCOMMAND     = 0x0112;
    public const uint SC_KEYMENU        = 0xF100;
    public const ushort VK_MENU         = 0x12;   // Alt

    public static INPUT KD(ushort vk){ INPUT i=new INPUT(); i.type=INPUT_KEYBOARD; i.U.ki.wVk=vk; return i; }
    public static INPUT KU(ushort vk){ INPUT i=new INPUT(); i.type=INPUT_KEYBOARD; i.U.ki.wVk=vk; i.U.ki.dwFlags=KEYEVENTF_KEYUP; return i; }
    public static int SZ(){ return Marshal.SizeOf(typeof(INPUT)); }

    public static void ForceFocus(IntPtr hwnd){
      uint cur = GetCurrentThreadId();
      IntPtr fg = GetForegroundWindow();
      uint fgThread = 0; if (fg != IntPtr.Zero) GetWindowThreadProcessId(fg, out fgThread);
      AttachThreadInput(cur, fgThread, true);
      ShowWindow(hwnd, 9); // SW_RESTORE
      BringWindowToTop(hwnd);
      SetForegroundWindow(hwnd);
      AttachThreadInput(cur, fgThread, false);
    }
  }
}
"@
if (-not ('IMC.Kbd' -as [type])) {
  try { Add-Type -TypeDefinition $nativeCode -Language CSharp -ErrorAction Stop }
  catch { W ("ERR: Add-Type native -> " + $_.Exception.Message); exit 5 }
}

# ========== UIA props & helpers ==========
$P_CT   = [System.Windows.Automation.AutomationElement]::ControlTypeProperty
$P_PID  = [System.Windows.Automation.AutomationElement]::ProcessIdProperty
$P_NAME = [System.Windows.Automation.AutomationElement]::NameProperty
$P_CLASS= [System.Windows.Automation.AutomationElement]::ClassNameProperty
$P_NWH  = [System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty

function Dump-Top([System.Windows.Automation.AutomationElement]$el){
  try{
    $nm = $el.Current.Name; $cl = $el.Current.ClassName; $r = $el.Current.BoundingRectangle
    W ("Top: Name='{0}' Class='{1}' Rect={2}x{3}+{4}+{5}" -f $nm,$cl,[int]$r.Width,[int]$r.Height,[int]$r.X,[int]$r.Y)
  }catch{}
}

function Wait-WordReady([int]$maxMs=12000){
  $deadline=[Environment]::TickCount + $maxMs; $seen=0
  do{
    try{
      $condPid = New-Object System.Windows.Automation.PropertyCondition ($P_PID,$wordPid)
      $els=[System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Subtree,$condPid)
      $hasRibbon=$false
      for($i=0;$i -lt $els.Count;$i++){
        $e=$els.Item($i)
        $cl=""; try{ $cl=[string]$e.GetCurrentPropertyValue($P_CLASS) }catch{}
        if($cl -in @('NetUIHWNDElement','NetUIRibbon','MsoCommandBar','MsoWorkPane')){ $hasRibbon=$true; break }
      }
      if($hasRibbon){ $seen++ } else { $seen=0 }
      if($seen -ge 2){ W "Ribbon detected (ready)"; return $true }
    }catch{ W ("Wait-WordReady transient: " + $_.Exception.Message) }
    Start-Sleep -Milliseconds 300
  } while([Environment]::TickCount -lt $deadline)
  W "WARN: Ribbon not detected in time"; return $false
}

function Get-Ribbon(){
  $condPid = New-Object System.Windows.Automation.PropertyCondition ($P_PID,$wordPid)
  $els=[System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Subtree,$condPid)
  for($i=0;$i -lt $els.Count;$i++){
    $e=$els.Item($i)
    $cl=""; try{ $cl=[string]$e.GetCurrentPropertyValue($P_CLASS) }catch{}
    if($cl -in @('NetUIHWNDElement','NetUIRibbon','MsoCommandBar')){
      return $e
    }
  }
  return $null
}
function Get-RibbonHwnd(){
  $e = Get-Ribbon
  if($e -ne $null){
    try{ $nwh=[int]$e.GetCurrentPropertyValue($P_NWH) }catch{ $nwh=0 }
    if($nwh -gt 0){ try{ $e.SetFocus() }catch{}; return [System.IntPtr]::new([int64]$nwh) }
  }
  return [IntPtr]::Zero
}

function Get-TaskPaneSnapshot([System.Windows.Automation.AutomationElement]$scope){
  $win = $scope.Current.BoundingRectangle; $right=$win.Right; $left=$win.Left
  $condType = New-Object System.Windows.Automation.OrCondition (
    (New-Object System.Windows.Automation.PropertyCondition ($P_CT,[System.Windows.Automation.ControlType]::Pane)),
    (New-Object System.Windows.Automation.PropertyCondition ($P_CT,[System.Windows.Automation.ControlType]::Custom))
  )
  $condPid  = New-Object System.Windows.Automation.PropertyCondition ($P_PID,$wordPid)
  $condAll  = New-Object System.Windows.Automation.AndCondition @($condType,$condPid)
  $els=$null
  try{ $els=[System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Subtree,$condAll) }
  catch{ W ("Get-TaskPaneSnapshot transient: " + $_.Exception.Message); return @() }
  $res=@()
  for($i=0;$i -lt $els.Count;$i++){
    $e=$els.Item($i); $r=$e.Current.BoundingRectangle
    if($r.Width -le 150 -or $r.Height -le 150){ continue }
    $touchRight = [Math]::Abs($right - $r.Right) -le 16
    $touchLeft  = [Math]::Abs($left  - $r.Left ) -le 16
    $floatOK    = -not $touchRight -and -not $touchLeft -and $r.Width -ge 260 -and $r.Height -ge 260
    if($touchRight -or $touchLeft -or $floatOK){
      $res += [pscustomobject]@{
        Name=$e.Current.Name; Class=$e.Current.ClassName;
        X=[int]$r.X; Y=[int]$r.Y; W=[int]$r.Width; H=[int]$r.Height;
        Side = $( if($touchRight){"Right"} elseif($touchLeft){"Left"} else {"Float"} )
      }
    }
  }
  return ,$res
}
function Snap-Changed($before,$after){
  if($after.Count -gt $before.Count){ return $true }
  foreach($a in $after){
    $m = $before | Where-Object { $_.Side -eq $a.Side -and $_.W -eq $a.W -and $_.H -eq $a.H -and $_.X -eq $a.X -and $_.Y -eq $a.Y }
    if(-not $m){ return $true }
  }
  return $false
}
function Dump-Snap([string]$label,$arr){
  W ("{0}: {1} candidate(s)" -f $label,$arr.Count)
  foreach($i in $arr){ W ("  {5} [{0}x{1}+{2}+{3}] Name='{4}' Class='{6}'" -f $i.W,$i.H,$i.X,$i.Y,$i.Name,$i.Side,$i.Class) }
}

# ========== Focus helpers ==========
function Focus-Word(){
  try{ [IMC.Kbd]::ForceFocus($HWND_PTR) }catch{}
  try{ (New-Object -ComObject WScript.Shell).AppActivate($wordPid) | Out-Null }catch{}
  try{ $root.SetFocus() }catch{}
  Start-Sleep -Milliseconds 120
}
function Focus-Ribbon([ref]$ribbonHwnd){
  $h = Get-RibbonHwnd
  if($h -ne [IntPtr]::Zero){ $ribbonHwnd.Value = $h; return $true }
  return $false
}

# ========== KeyTips core ==========
function Tap-Alt(){
  $sz=[IMC.Kbd]::SZ()
  [IMC.Kbd]::SendInput(1, ([IMC.Kbd]::KD([IMC.Kbd]::VK_MENU) -as [IMC.Kbd+INPUT[]]), $sz) | Out-Null
  Start-Sleep -Milliseconds 70
  [IMC.Kbd]::SendInput(1, ([IMC.Kbd]::KU([IMC.Kbd]::VK_MENU) -as [IMC.Kbd+INPUT[]]), $sz) | Out-Null
  Start-Sleep -Milliseconds 160
}
function Send-YA-E(){
  if($formsOK){
    [System.Windows.Forms.SendKeys]::SendWait('+я')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait('+э')
  } else {
    $ws = New-Object -ComObject WScript.Shell
    $ws.SendKeys('+я'); Start-Sleep -Milliseconds 200
    $ws.SendKeys('+э')
  }
  W "SendKeys(+я,+э) sent"
}

function PassA([ref]$ribbonH){
  # прицельно посылаем SC_KEYMENU в ленту, затем лёгкий Alt, затем +я +э
  if(-not (Focus-Ribbon ([ref]$ribbonH))){ W "WARN: Ribbon HWND not found"; return }
  $rh = $ribbonH.Value
  W ("Ribbon HWND=0x{0:X}" -f $rh.ToInt64())
  [void][IMC.Kbd]::PostMessage($rh, [IMC.Kbd]::WM_SYSCOMMAND, [IntPtr][IMC.Kbd]::SC_KEYMENU, [IntPtr]::Zero)
  Start-Sleep -Milliseconds 220
  Tap-Alt
  Start-Sleep -Milliseconds 120
  # На всякий случай ещё раз фокус на ленту 45 чтобы буквы точно не ушли в документ
  try{ $null = (Get-Ribbon).SetFocus() }catch{}
  Start-Sleep -Milliseconds 60
  Send-YA-E
}

# ========== Undo-guard (если вдруг символы попали в документ без панели) ==========
function Undo-IfTyped($beforeSnap,$afterSnap){
  if(Snap-Changed $beforeSnap $afterSnap){ return $false } # панель появилась 45 ничего не трогаем
  # мягкий откат одного действия
  try{
    if($formsOK){ [System.Windows.Forms.SendKeys]::SendWait('^z') } else { (New-Object -ComObject WScript.Shell).SendKeys('^z') }
    W "Guard: sent Ctrl+Z to rollback accidental text"
    return $true
  }catch{ return $false }
}

# ========== MAIN ==========
try{
  Dump-Top $root
  [void](Wait-WordReady 12000)

  Focus-Word
  $before = Get-TaskPaneSnapshot $root
  Dump-Snap "Before" $before

  $ribbonH=[IntPtr]::Zero
  # Первая попытка
  W "Attempt #1 (SC_KEYMENU -> AltTap -> +я +э)"
  PassA ([ref]$ribbonH)
  $deadline = [Environment]::TickCount + 6500
  $ok = $false
  do{
    $after = Get-TaskPaneSnapshot $root
    if(Snap-Changed $before $after){ Dump-Snap "After(1)" $after; $ok=$true; break }
    Start-Sleep -Milliseconds 200
  } while([Environment]::TickCount -lt $deadline)

  if(-not $ok){
    # Защита от постороннего ввода и ретрай
    [void](Undo-IfTyped $before (Get-TaskPaneSnapshot $root))
    Focus-Word
    Start-Sleep -Milliseconds 120
    W "Attempt #2 (refocus ribbon + repeat)"
    PassA ([ref]$ribbonH)
    $deadline2 = [Environment]::TickCount + 6500
    do{
      $after2 = Get-TaskPaneSnapshot $root
      if(Snap-Changed $before $after2){ Dump-Snap "After(2)" $after2; $ok=$true; break }
      Start-Sleep -Milliseconds 200
    } while([Environment]::TickCount -lt $deadline2)
  }

  if($ok){ W "Task Pane detected (KeyTips)"; exit 0 }
  else    { W "WARN: Task Pane not detected in time (KeyTips)"; exit 4 }
}
catch{
  W ("ERR: " + $_.Exception.Message)
  exit 9
}
PS C:\Users\officeuser> Get-Content C:\DocOpsRunner\Word-Finalize-WSH.ps1
param(
  [int]$TimeoutSec = 300,              # общий потолок жизни финализатора
  [int]$WaitForMarkerSec = 0,          # 0 = ориентируемся только на /done; >0 = авар. fallback по <BLOCK:
  [int]$ListenLocalHttpPort = 17603,   # порт локального сигнала
  [int]$BgWaitMaxMs = 3500,            # максимум ожидания фоновых сохранений для WOPI (ms)
  [string]$TraceId = ""                # единый traceId из агента
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Off

# --- Логирование в файл ---
$logDir = 'C:\DocOpsRunner\logs'
if(-not (Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$log = Join-Path $logDir ("finalize-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
function W([string]$m){ "[{0:yyyy-MM-dd HH:mm:ssK}]|[FINALIZE] {1}" -f (Get-Date),$m | Out-File -FilePath $log -Append -Encoding UTF8 }

# --- Серверные логи ---
try { . 'C:\DocOpsRunner\Lib-ServerLog.ps1' } catch {}
try{ Rotate-DocOpsLogs -Days 3 }catch{}

# --- Утилиты ---
function Try-GetWordApp {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { return $null }
}
function Wait-ForWord([int]$maxMs){
  $deadline = [Environment]::TickCount + [Math]::Max(0,$maxMs)
  do{
    $app = Try-GetWordApp
    if($app){ return $app }
    Start-Sleep -Milliseconds 200
  } while([Environment]::TickCount -lt $deadline)
  return $null
}
function Is-WopiDoc([object]$doc){
  try{
    $fn = ""
    try{ $fn = [string]$doc.FullName }catch{}
    if($fn -and $fn.ToLower().StartsWith("http")){ return $true }
  }catch{}
  return $false
}
function Wait-BackgroundIdle([object]$app, [int]$maxMs){
  if($maxMs -le 0){ W "BgWait: skip (maxMs<=0)"; return $true }
  $deadline = [Environment]::TickCount + $maxMs
  $zeros = 0
  do{
    $busy = 0
    try{ $busy = [int]$app.BackgroundSavingStatus }catch{ $busy = 0 }
    if($busy -le 0){
      $zeros++
      if($zeros -ge 2){ W "BgWait: idle (two consecutive zeros)"; return $true }
    } else {
      $zeros = 0
      W ("BgWait: busy={0}" -f $busy)
    }
    Start-Sleep -Milliseconds 200
  } while([Environment]::TickCount -lt $deadline)
  W "BgWait: timeout"
  return $false
}
function Test-MarkerAnyBlock {
  try{
    $app = Try-GetWordApp
    if($app -and $app.Documents.Count -gt 0){
      $doc=$app.ActiveDocument
      $r=$doc.Content; $f=$r.Find; $f.ClearFormatting(); $f.Text="<BLOCK:"
      if($f.Execute()){ return $true }
    }
  }catch{}
  return $false
}

function Safe-SaveCloseQuit([object]$app,[string]$trace,[string]$jobId){
  try{
    if(-not $app){ W "WARN: app=null in Safe-SaveCloseQuit"; return }
    $doc=$null
    try{ $doc=$app.ActiveDocument }catch{}
    if(-not $doc -and $app.Documents.Count -gt 0){ try{ $doc=$app.Documents.Item(1) }catch{} }

    $docsCount = 0; try{ $docsCount = $app.Documents.Count }catch{}
    W ("COM: Attached. Visible={0} Docs={1}" -f $app.Visible, $docsCount)
    try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'com.attached' -Message ('docs={0}' -f $docsCount) -TraceId $trace -JobId $jobId }catch{}

    $isWopi = $false
    try{ if($doc){ $isWopi = Is-WopiDoc $doc } }catch{}
    W ("Doc: WOPI={0}" -f $isWopi)

    try{
      if($doc -and $doc.Sync){
        $st = $doc.Sync.Status.value__
        W ("Sync.Status(before)={0}" -f $st)
        $doc.Sync.PutUpdate()
        W "Sync.PutUpdate() called"
      }
    }catch{ W ("WARN: Sync.PutUpdate -> " + $_.Exception.Message) }

    if($doc){
      $saved=$false
      try{ $doc.Save(); $saved=$true; W "Saved via .Save()" }catch{ W ("WARN: .Save() -> " + $_.Exception.Message) }
      try{ $app.WordBasic.FileSave(); W "WordBasic.FileSave (extra)" }catch{}
      try{ if($saved){ Send-ServerLog -Level info -Phase 'finalize' -Event 'saved' -Message 'doc.Save()' -TraceId $trace -JobId $jobId } }catch{}
    } else {
      W "WARN: No ActiveDocument for .Save()"
    }

    $bgMs = 0
    if($isWopi){ $bgMs = [Math]::Max(0,$BgWaitMaxMs) }
    [void](Wait-BackgroundIdle -app $app -maxMs $bgMs)

    $closed=$false
    try{
      if($doc){ $doc.Close(-1); $closed=$true; W "Closed ActiveDocument (SaveChanges)"; try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'closed' -Message 'doc.Close(-1)' -TraceId $trace -JobId $jobId }catch{} }
    }catch{ W ("WARN: .Close(-1) -> " + $_.Exception.Message) }

    if(-not $closed){
      try{ $app.CommandBars.ExecuteMso('FileClose'); $closed=$true; W "ExecuteMso(FileClose)"; try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'closed' -Message 'ExecuteMso(FileClose)' -TraceId $trace -JobId $jobId }catch{} }
      catch{ W ("WARN: ExecuteMso(FileClose) -> " + $_.Exception.Message) }
    }

    if(-not $closed){
      try{
        $shell=New-Object -ComObject WScript.Shell
        try{ [void]$shell.AppActivate('Microsoft Word') }catch{}
        Start-Sleep -Milliseconds 200
        $shell.SendKeys("^s"); W "SendKeys: Ctrl+S"
        Start-Sleep -Milliseconds 700
        $shell.SendKeys("%{F4}"); W "SendKeys: Alt+F4"
        try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'closed' -Message 'SendKeys Alt+F4' -TraceId $trace -JobId $jobId }catch{}
      }catch{ W ("WARN: SendKeys fallback -> " + $_.Exception.Message) }
    }

    try{ $app.Quit(); W "Quit Word"; try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'quit' -Message 'app.Quit()' -TraceId $trace -JobId $jobId }catch{} }
    catch{ W ("WARN: .Quit() -> " + $_.Exception.Message) }
  }catch{
    W ("ERR in Safe-SaveCloseQuit: " + $_.Exception.Message)
    try{ Send-ServerLog -Level error -Phase 'finalize' -Event 'exception' -Message $_.Exception.Message -TraceId $trace -JobId $jobId }catch{}
  }
}

# --- Главная логика ---
try{
  $sess = (Get-Process -Id $PID).SessionId
  W ("Boot. User={0} Session={1}" -f $env:USERNAME, $sess)

  # HTTP listener
  $listener = New-Object System.Net.HttpListener
  $prefix = "http://127.0.0.1:{0}/" -f $ListenLocalHttpPort
  $listener.Prefixes.Add($prefix)
  try{
    $listener.Start()
    W "HTTP listen $prefix"
    try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'listen' -Message $prefix -TraceId $TraceId }catch{}
  }
  catch{
    W ("ERR: HTTP listen failed -> " + $_.Exception.Message)
    try{ Send-ServerLog -Level error -Phase 'finalize' -Event 'listen.fail' -Message $_.Exception.Message -TraceId $TraceId }catch{}
    throw
  }

  $done   = $false
  $jobId  = ""
  $start  = Get-Date
  $deadline = $start.AddSeconds([Math]::Max(5,$TimeoutSec))

  # Асинхронно ждём HTTP-запросы
  $ar = $listener.BeginGetContext($null,$null)

  while(-not $done -and (Get-Date) -lt $deadline){
    if($ar.AsyncWaitHandle.WaitOne(200)){
      try{
        $ctx = $listener.EndGetContext($ar)
        $path = "/"; try{ $path = $ctx.Request.Url.AbsolutePath.ToLowerInvariant() }catch{}
        if($path -eq "/done"){
          try{ $jobId = [string]$ctx.Request.QueryString["jobId"] }catch{}
          W ("HTTP /done received. jobId={0}" -f $jobId)
          try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'trigger' -Message 'HTTP /done' -TraceId $TraceId -JobId $jobId }catch{}
          $bytes=[Text.Encoding]::UTF8.GetBytes("OK"); $ctx.Response.StatusCode=200
          $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length); $ctx.Response.Close()
          $done = $true; break
        } elseif ($path -eq "/ping"){
          $bytes=[Text.Encoding]::UTF8.GetBytes("pong"); $ctx.Response.StatusCode=200
          $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length); $ctx.Response.Close()
        } elseif ($path -eq "/health"){
          $bytes=[Text.Encoding]::UTF8.GetBytes("OK"); $ctx.Response.StatusCode=200
          $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length); $ctx.Response.Close()
        } else {
          $ctx.Response.StatusCode=404; $ctx.Response.Close()
        }
      }catch{ W ("WARN: HTTP accept -> " + $_.Exception.Message) }
      finally{ if(-not $done){ $ar = $listener.BeginGetContext($null,$null) } }
    }

    if(-not $done -and $WaitForMarkerSec -gt 0){
      if((Get-Date) -ge $start.AddSeconds($WaitForMarkerSec)){
        if(Test-MarkerAnyBlock){ W "Marker '<BLOCK:' detected (fallback)"; $done=$true; break }
      }
    }
  }

  if(-not $done){ W "Timeout waiting for /done/marker (will attempt safe close anyway)" }

  $remainMs = [int][Math]::Max(1200, ($deadline - (Get-Date)).TotalMilliseconds)
  $app = Wait-ForWord $remainMs
  if(-not $app){ W "WARN: Word not running (will not COM-close)" }
  else { Safe-SaveCloseQuit -app $app -trace $TraceId -jobId $jobId }

  try{ Send-ServerLog -Level info -Phase 'finalize' -Event 'done' -Message ("jobId={0}" -f $jobId) -TraceId $TraceId -JobId $jobId }catch{}
} catch {
  W ("ERR: " + $_.Exception.Message)
  try{ Send-ServerLog -Level error -Phase 'finalize' -Event 'exception' -Message $_.Exception.Message -TraceId $TraceId -JobId $jobId }catch{}
} finally {
  try{ if($listener -and $listener.IsListening){ $listener.Stop() } }catch{}
}