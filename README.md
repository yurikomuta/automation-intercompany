# automation-intercompany
automação para administração de arquivos intercompany

#### Criação do ambiente virtual

` python -m venv .venv `

#### Ativação da venv win

` .\.venv\Scripts\Activate.ps1 `

#### Ativação unix\macos

 ` source .venv/bin/activate`

#### Execucao pelo PowerShell

Na pasta do projeto, execute:

` .\executar-rpa.ps1 `

Se o PowerShell bloquear a execucao de scripts locais, use apenas nesta sessao:

` Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass `