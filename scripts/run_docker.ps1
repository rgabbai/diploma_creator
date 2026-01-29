$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Resolve-Path (Join-Path $ScriptDir "..")

Set-Location $RepoRoot

docker compose up --build

Write-Host "Open: http://localhost:8000"
Write-Host "Output folder: webappGamilAPI/output"
