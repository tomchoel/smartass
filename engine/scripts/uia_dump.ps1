# Dumps top-level windows of Excel / WebView2 processes and their UI Automation tree (buttons, links, text).
param([int]$MaxDepth = 12, [string]$Filter = '')
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$pids = @(Get-Process EXCEL, msedgewebview2 -ErrorAction SilentlyContinue | ForEach-Object Id)
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = [System.Windows.Automation.Condition]::TrueCondition
$tops = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

function Walk($el, $depth) {
    if ($depth -gt $MaxDepth) { return }
    $c = $walker.GetFirstChild($el)
    while ($c) {
        try {
            $ct = $c.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
            $nm = $c.Current.Name
            if ($nm -and ($ct -in 'Button','Hyperlink','Text','Edit','Window','Pane','Document','ListItem') -and ($Filter -eq '' -or $nm -match $Filter)) {
                ('  ' * $depth) + "[$ct] $nm"
            }
        } catch {}
        Walk $c ($depth + 1)
        $c = $walker.GetNextSibling($c)
    }
}

foreach ($t in $tops) {
    if ($pids -contains $t.Current.ProcessId) {
        "=== PID $($t.Current.ProcessId) | $($t.Current.Name) | class=$($t.Current.ClassName)"
        $n = $t.Current.Name
        if ($n -notmatch '^Book|- Excel$' -or $Filter) { Walk $t 1 }
    }
}
