# Global SSO watcher (started hidden at logon). Clicks "Telenor SSO" in the Smart View EPM Cloud "Login" dialog
# (#32770) of ANY Excel: the smartview MCP's, excel-mcp's and the user's own. Only invokes that one button; never types.
# .NET UI Automation is used because the Python 'uiautomation' package blocked on this dialog in interactive Excel.
param([int]$IntervalMs = 1000, [string]$LogDir = (Join-Path $env:LOCALAPPDATA 'smartass\logs'))
$ErrorActionPreference = 'Continue'
$created = $false
$mutex = [System.Threading.Mutex]::new($true, 'Local\svkit-sso-watcher-ps', [ref]$created)
if (-not $created) { exit 0 }

$logDir = $LogDir
New-Item -ItemType Directory -Force $logDir | Out-Null
$log = Join-Path $logDir 'sso_watcher_ps.log'
function Log($m) {
    try { [System.IO.File]::AppendAllText($log, ("{0} {1}`r`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m)) } catch { }
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$scope = [System.Windows.Automation.TreeScope]
$loginCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'Login')
$btnCond = New-Object System.Windows.Automation.AndCondition(
    (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)),
    (New-Object System.Windows.Automation.PropertyCondition($AE::IsEnabledProperty, $true)))

Log "global SSO watcher started (pid $PID)"
$beat = Get-Date
$hbFile = Join-Path $logDir 'sso_watcher.heartbeat'
$hbLast = [datetime]::MinValue
$seen = @{}                                   # Login dialog hwnd -> first time seen (diagnoses slow clicks)
$loopAt = Get-Date
while ($true) {
    $gap = ((Get-Date) - $loopAt).TotalSeconds
    if ($gap -gt 5) { Log ("loop stalled {0:N1}s (a UI Automation call blocked)" -f $gap) }
    $loopAt = Get-Date
    if (((Get-Date) - $hbLast).TotalSeconds -ge 10) {      # liveness for the MCP (restarts a stuck watcher)
        try { [System.IO.File]::WriteAllText($hbFile, "$PID $((Get-Date).ToString('o'))") } catch { }
        $hbLast = Get-Date
    }
    try {
        $pids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
        if ($pids.Count) {
            foreach ($w in $AE::RootElement.FindAll($scope::Children, $loginCond)) {
                if ($pids -notcontains $w.Current.ProcessId -or $w.Current.ClassName -ne '#32770') { continue }
                $h = $w.Current.NativeWindowHandle
                if (-not $seen.ContainsKey($h)) { $seen[$h] = Get-Date; Log "Login dialog seen in Excel pid $($w.Current.ProcessId) (hwnd $h)" }
                $btn = $w.FindAll($scope::Descendants, $btnCond) | Where-Object { $_.Current.Name -match 'Telenor SSO' } | Select-Object -First 1
                if ($btn) {
                    $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                    Log ("clicked Telenor SSO in Login dialog of Excel pid {0} ({1:N1}s after it appeared)" -f $w.Current.ProcessId, ((Get-Date) - $seen[$h]).TotalSeconds)
                    Start-Sleep -Seconds 3
                }
            }
        }
    } catch {
        Log "sweep error: $($_.Exception.Message)"
    }
    if (((Get-Date) - $beat).TotalMinutes -ge 10) { Log "alive; excel pids: $($pids -join ',')"; $beat = Get-Date }
    Start-Sleep -Milliseconds $IntervalMs
}
