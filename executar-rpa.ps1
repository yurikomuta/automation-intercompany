$ErrorActionPreference = 'Stop'

$projectPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonPath = Join-Path $projectPath '.venv\Scripts\python.exe'
$scriptPath = Join-Path $projectPath 'rpa.py'

if (-not (Test-Path $pythonPath)) {
    throw "Python da venv nao encontrado em '$pythonPath'. Crie o ambiente com: python -m venv .venv"
}

if (-not (Test-Path $scriptPath)) {
    throw "Arquivo rpa.py nao encontrado em '$projectPath'."
}

& $pythonPath $scriptPath
exit $LASTEXITCODE