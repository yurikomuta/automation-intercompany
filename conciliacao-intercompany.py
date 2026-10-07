"""Conciliação Intercompany — bibliotecas recomendadas.

Portal E2E e download:
	playwright — biblioteca open-source e gratuita (licença Apache 2.0), com
	automação web confiável, esperas por elementos e downloads. Não exige licença
	paga.
"""


"""
Autenticação e 2FA:
	Como o usuário utiliza o Outlook no PC, use pywin32 (win32com.client) para
	acessar o Outlook instalado e consultar a caixa de entrada, se autorizado
	pela empresa. A integração usa o perfil já configurado no Outlook; não
	armazene senhas no código.

Datas, arquivos e logs:
	datetime/calendar (biblioteca padrão) para calcular o primeiro dia útil;
	pathlib/shutil para localizar e mover o arquivo à pasta de rede; logging
	para registrar execuções e falhas. Use pandas/openpyxl somente se precisar
	ler ou editar o Excel.

Execuções várias vezes ao dia:
	Agendador de Tarefas do Windows é a opção mais simples para iniciar o script
	em horários definidos. APScheduler serve se o script precisar ficar rodando.
	Evite execuções simultâneas e confirme que o download terminou antes de mover
	o arquivo.

Instalação básica (confirme a autenticação de e-mail com a equipe de TI):
	pip install playwright pywin32
	playwright install chromium
"""

from playwright.sync_api import sync_playwright


def acessar_portal():
	with sync_playwright() as playwright:
		browser = playwright.chromium.launch(headless=False)
		page = browser.new_page()
		page.goto("https://e2e.bps-accenture.com/", wait_until="domcontentloaded")
		page.get_by("LoginForm").fill("yuri.komuta.ext@telefonica.com")
		page.get_by_role("form").fill("senha123")
		page.get_by_role("button", name="Login").click()
		input("Pressione Enter para fechar o navegador...")
		browser.close()


if __name__ == "__main__":
	acessar_portal()
