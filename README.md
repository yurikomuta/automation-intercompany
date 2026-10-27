# Automação Intercompany

Aplicativo desktop para atualizar a caixa de entrada do Outlook e salvar anexos dos relatórios de Eliminação Patrimonial e Eliminação Resultado.

## Funcionalidades

- **Atualizar caixa de entrada:** executa uma sincronização do Outlook e aguarda a conclusão, com limite de cinco minutos.
- **Eliminação Patrimonial:** procura os e-mails correspondentes na Caixa de Entrada e salva os anexos reconhecidos.
- **Eliminação Resultado:** procura os e-mails correspondentes na Caixa de Entrada e salva os anexos reconhecidos.
- **Log na tela:** apresenta horário, resultado da ação, quantidade de anexos salvos e erros. O botão de lixeira limpa o log da sessão.

Durante uma operação, os outros botões ficam desabilitados para evitar execuções simultâneas. As rotinas de eliminação também sincronizam o Outlook antes de buscar e-mails.

## Organização dos arquivos

Os anexos são organizados por ano, mês, tipo de eliminação e empresa:

```text
2 - Intercompany/
└── ANO_<ano>/
	└── <mês>_<ano>/
		└── APLICATIVO ELIMINAÇÃO/
			├── Eliminação Patrimonial/
			│   └── <empresa>/
			│       ├── Checklist/
			│       └── Versões Anteriores/
			└── Eliminação Resultado/
				└── <empresa>/
					├── Checklist/
					└── Versões Anteriores/
```

O app também organiza pastas de empresa já existentes antes da busca de e-mails. Para arquivos com o mesmo nome-base e extensão, a maior versão numérica (`_V<número>`) permanece na pasta da empresa; versões anteriores são movidas para `Versões Anteriores`. Um arquivo sem sufixo é tratado como versão zero quando existe uma versão numerada correspondente. Arquivos ZIP são extraídos na pasta da empresa e removidos após a extração bem-sucedida. Arquivos já existentes não são salvos novamente.

## Requisitos

- Windows 10 ou 11 de 64 bits.
- Python 3.10 ou posterior.
- Microsoft Outlook desktop instalado, com perfil e Caixa de Entrada configurados. A automação usa Outlook via COM; o novo Outlook não oferece essa interface COM.
- Acesso à pasta de destino configurada nos scripts (`Z:\brtssisp003\bbb0133$\2 - Intercompany`).

## Instalação

Abra o PowerShell na pasta do projeto e crie o ambiente virtual:

```powershell
python -m venv .venv
```

Instale as dependências do aplicativo:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-app.txt
```

O arquivo `requirements-app.txt` fixa a versão do Flet e instala o cliente desktop e o `pywin32`.

## Executar

Abra `Abrir-Automacao.bat` com duplo clique ou execute no PowerShell:

```powershell
.\.venv\Scripts\python.exe app.py
```

## Observações

- O log exibido no app existe apenas durante a sessão e mantém as 300 mensagens mais recentes; ele não é gravado em arquivo.
- Erros da automação são exibidos no log e no indicador de status da janela.
- O caminho de destino pode ser alterado pela constante `PASTA_DESTINO` nos scripts `eliminacao-patrimonial.py` e `eliminacao-resultado.py`.