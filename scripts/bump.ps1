<# Roll a code change into running smartass servers: bump engine/VERSION (and the same version in
   engine/pyproject.toml). The relay run_server.py reloads its server on the next request once VERSION changes;
   a .py edit alone does not. Alternative without a bump: the tool call fin_status(restart=True).
   -Part patch (default) | minor | major. -Set 1.2.3 sets an exact version. #>
param([ValidateSet('patch', 'minor', 'major')][string]$Part = 'patch', [string]$Set)
$ErrorActionPreference = 'Stop'
$engine = Join-Path (Split-Path $PSScriptRoot -Parent) 'engine'
$verFile = Join-Path $engine 'VERSION'
$pyproject = Join-Path $engine 'pyproject.toml'
$old = (Get-Content $verFile -Raw).Trim()
if ($Set) { $new = $Set }
else {
    $p = $old.Split('.') | ForEach-Object { [int]$_ }
    switch ($Part) {
        'major' { $new = "$($p[0] + 1).0.0" }
        'minor' { $new = "$($p[0]).$($p[1] + 1).0" }
        default { $new = "$($p[0]).$($p[1]).$($p[2] + 1)" }
    }
}
if ($new -notmatch '^\d+\.\d+\.\d+$') { throw "not a version: $new" }
$utf8 = New-Object Text.UTF8Encoding $false
[IO.File]::WriteAllText($verFile, "$new`n", $utf8)
$text = [IO.File]::ReadAllText($pyproject)
$text = [regex]::Replace($text, '(?m)^version = "[^"]+"', "version = `"$new`"", 1)
[IO.File]::WriteAllText($pyproject, $text, $utf8)
"smartass $old -> $new (engine/VERSION + pyproject.toml); running servers reload on their next request"
