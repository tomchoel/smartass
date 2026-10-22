<# Publish the committed HEAD as a release: <Folder>\<version>\ + latest.json (+ atlas snapshot), tag v<version>. #>
param([Parameter(Mandatory)][string]$Folder, [switch]$WithAtlas)
$ErrorActionPreference = 'Stop'
$Folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Folder)   # .NET calls need it absolute
$repo = Split-Path $PSScriptRoot -Parent
if (git -C $repo status --porcelain) { throw 'working tree not clean - commit first' }
$ver = (Select-String -Path "$repo\engine\pyproject.toml" -Pattern '^version = "(.+)"').Matches[0].Groups[1].Value
$dest = Join-Path $Folder $ver
if (Test-Path $dest) { throw "release $ver already exists in $Folder - bump the version" }
$zip = Join-Path $env:TEMP "smartass-$ver.zip"
git -C $repo archive --format zip -o $zip HEAD
Expand-Archive $zip $dest; Remove-Item $zip
if ($WithAtlas) {
    New-Item -ItemType Directory -Force (Join-Path $Folder 'atlas') | Out-Null
    $py = "$repo\engine\.venv\Scripts\python.exe"
    & $py -c "import sqlite3,sys; s=sqlite3.connect(sys.argv[1]); d=sqlite3.connect(sys.argv[2]); s.backup(d); d.close(); s.close()" `
        "$env:LOCALAPPDATA\smartass\atlas.db" (Join-Path $Folder 'atlas\atlas.db')
}
# UTF-8 without BOM: PowerShell 5.1's -Encoding utf8 adds a BOM, which update.py's json.loads rejects.
$json = @{ version = $ver; published = (Get-Date -Format 'yyyy-MM-dd') } | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $Folder 'latest.json'), $json, (New-Object Text.UTF8Encoding $false))
git -C $repo tag "v$ver"; git -C $repo push origin "v$ver"
"published $ver to $Folder"
