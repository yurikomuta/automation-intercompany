$ErrorActionPreference = "Stop"

$projectDir = $PSScriptRoot
$python = Join-Path $projectDir ".venv\Scripts\python.exe"
$pyinstaller = Join-Path $projectDir ".venv\Scripts\pyinstaller.exe"
$buildDir = Join-Path $projectDir "build-homologacao"
$workDir = Join-Path $buildDir "work"
$distDir = Join-Path $projectDir "dist-homologacao"

if (-not (Test-Path $python) -or -not (Test-Path $pyinstaller)) {
	throw "Ambiente de build ausente. Crie .venv e instale requirements-build.txt."
}

$fletVersion = & $python -c "import flet_desktop.version; print(flet_desktop.version.version)"
if ($LASTEXITCODE -ne 0) {
	throw "Nao foi possivel identificar a versao do cliente Flet instalado."
}

$runtimeDir = Join-Path $env:USERPROFILE ".flet\client\flet-desktop-full-$fletVersion\flet"
if (-not (Test-Path (Join-Path $runtimeDir "flet.exe"))) {
	throw "Cliente desktop Flet nao encontrado em $runtimeDir. Execute o app uma vez nesta maquina de build."
}

$payloadDir = Join-Path $buildDir "payload"
$runtimeZip = Join-Path $payloadDir "flet-windows.zip"
New-Item -ItemType Directory -Force -Path $payloadDir, $workDir, $distDir | Out-Null
Compress-Archive -Path $runtimeDir -DestinationPath $runtimeZip -CompressionLevel Fastest -Force

$arguments = @(
	"--clean",
	(Join-Path $projectDir "app.py"),
	"--noconsole",
	"--windowed",
	"--onefile",
	"--name", "automacao_intercompany",
	"--distpath", $distDir,
	"--workpath", $workDir,
	"--specpath", $buildDir,
	"--add-data", "$(Join-Path $projectDir 'attachment_archive.py');.",
	"--add-data", "$(Join-Path $projectDir 'company_folder.py');.",
	"--add-data", "$(Join-Path $projectDir 'outlook_sync.py');.",
	"--add-data", "$(Join-Path $projectDir 'eliminacao-patrimonial.py');.",
	"--add-data", "$(Join-Path $projectDir 'eliminacao-resultado.py');.",
	"--add-data", "$(Join-Path $projectDir 'composicao-geral.py');.",
	"--add-data", "$(Join-Path $projectDir 'quadro-resumo-patrimonial.py');.",
	"--add-data", "$(Join-Path $projectDir 'interno.py');.",
	"--add-data", "$(Join-Path $projectDir 'config');config",
	"--add-data", "$runtimeZip;flet_desktop/app",
	"--hidden-import", "pythoncom",
	"--hidden-import", "pywintypes",
	"--hidden-import", "win32com",
	"--hidden-import", "win32com.client",
	"--hidden-import", "flet_desktop",
	"--collect-all", "flet",
	"--collect-all", "flet_desktop"
)

Push-Location $projectDir
try {
	& $pyinstaller @arguments
	if ($LASTEXITCODE -ne 0) {
		throw "PyInstaller terminou com codigo $LASTEXITCODE."
	}
}
finally {
	Pop-Location
}

$exe = Join-Path $distDir "automacao_intercompany.exe"
Compress-Archive -Path $exe -DestinationPath (Join-Path $distDir "automacao_intercompany.zip") -Force
Write-Host "Executavel: $exe"
Write-Host "Pacote ZIP: $(Join-Path $distDir 'automacao_intercompany.zip')"