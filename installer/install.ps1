<#  smartass installer - per user, no admin, safe to re-run.
    From a release folder: copies the release to %LOCALAPPDATA%\smartass\app and installs there.
    From a git checkout (dev mode): installs in place and enables the pre-commit hook.  #>
param([switch]$NoDoctor)
$ErrorActionPreference = 'Stop'
$src = Split-Path $PSScriptRoot -Parent
$dev = Test-Path (Join-Path $src '.git')
if ($dev) { $app = $src } else {
    if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is not set' }
    # /MIR deletes extras in the target, so the target is only ever the dedicated app subfolder. User state
    # (config.toml, atlas.db, logs) lives one level up in %LOCALAPPDATA%\smartass and is never touched.
    $app = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'smartass\app'))
    if ([IO.Path]::GetFullPath($src).TrimEnd('\') -ieq $app) {
        Write-Host 'Running from the installed copy; installing in place.'
    } else {
        if ((Test-Path (Join-Path $app 'config.toml')) -or (Test-Path (Join-Path $app 'atlas.db'))) {
            throw "Refusing to mirror into $app : it contains user state (config.toml/atlas.db)."
        }
        Get-CimInstance Win32_Process -Filter "Name='pythonw.exe' or Name='python.exe'" |
            Where-Object { $_.CommandLine -like "*$app*" -and $_.CommandLine -like '*cli serve*' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
        robocopy $src $app /MIR /XD .venv __pycache__ .git /NFL /NDL /NJH /NJS /NP | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "robocopy failed ($LASTEXITCODE)" }
    }
}
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    $env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
}
$engine = Join-Path $app 'engine'
Push-Location $engine
try {
    if ($dev) { uv sync --frozen } else { uv sync --frozen --no-dev }
    if ($LASTEXITCODE) { throw "uv sync failed ($LASTEXITCODE)" }
    $py = Join-Path $engine '.venv\Scripts\python.exe'
    & $py -m smartass.ops.cli setup
    if ($LASTEXITCODE) { throw "smartass setup failed ($LASTEXITCODE)" }
    if ($dev) { git -C $app config core.hooksPath .githooks }
    if (-not $NoDoctor) { & $py -m smartass.ops.cli doctor }
} finally { Pop-Location }
Write-Host "smartass installed ($(if ($dev) {'dev checkout'} else {$app})). Restart Copilot CLI / Claude Code to load it."
