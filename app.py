import asyncio
from datetime import datetime
import importlib.util
from pathlib import Path
import sys
import flet as ft
import pythoncom
import win32com.client

PROJECT_DIR = Path(__file__).resolve().parent
if str(PROJECT_DIR) not in sys.path:
	sys.path.insert(0, str(PROJECT_DIR))

from outlook_sync import sincronizar_todas_as_pastas


def carregar_script(nome_modulo: str, nome_arquivo: str):
	caminho = PROJECT_DIR / nome_arquivo
	especificacao = importlib.util.spec_from_file_location(nome_modulo, caminho)
	if especificacao is None or especificacao.loader is None:
		raise ImportError(f"Nao foi possivel carregar {nome_arquivo}.")

	modulo = importlib.util.module_from_spec(especificacao)
	especificacao.loader.exec_module(modulo)
	return modulo


modulo_patrimonial = carregar_script("eliminacao_patrimonial", "eliminacao-patrimonial.py")
modulo_resultado = carregar_script("eliminacao_resultado", "eliminacao-resultado.py")


def executar_com_com(funcao):
	pythoncom.CoInitialize()
	try:
		return funcao()
	finally:
		pythoncom.CoUninitialize()


def atualizar_caixa_de_entrada() -> None:
	namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	sincronizar_todas_as_pastas(namespace)


def main(page: ft.Page) -> None:
	page.title = "Automacao Intercompany"
	page.theme_mode = ft.ThemeMode.LIGHT
	page.bgcolor = "#F3F6F3"
	page.padding = 0
	page.scroll = ft.ScrollMode.AUTO
	page.window.width = 980
	page.window.height = 760
	page.window.min_width = 760
	page.window.min_height = 620

	registros = []
	botoes = []
	ocupado = False

	texto_status = ft.Text("Pronto para iniciar", size=13, color="#285D4D")
	icone_status = ft.Icon(ft.Icons.CHECK_CIRCLE, size=17, color="#287A58")
	indicador_status = ft.Container(
		content=ft.Row([icone_status, texto_status], spacing=8),
		bgcolor="#E5F2E9",
		padding=ft.Padding.symmetric(horizontal=14, vertical=9),
		border_radius=20,
	)
	progresso = ft.ProgressRing(visible=False, width=20, height=20, stroke_width=2)
	caixa_log = ft.TextField(
		value="Aguardando a primeira ação.",
		multiline=True,
		read_only=True,
		min_lines=12,
		max_lines=12,
		text_style=ft.TextStyle(font_family="Cascadia Code", size=12, color="#34413A"),
		border=ft.NoInputBorder(),
		bgcolor=ft.Colors.TRANSPARENT,
		expand=True,
	)
	botao_limpar = ft.IconButton(
		icon=ft.Icons.DELETE_OUTLINE,
		tooltip="Limpar log",
		icon_color="#64736A",
	)

	def atualizar_log(mensagem: str) -> None:
		horario = datetime.now().strftime("%H:%M:%S")
		registros.append(f"{horario}  {mensagem}")
		caixa_log.value = "\n".join(registros[-300:])

	def definir_status(mensagem: str, cor: str, fundo: str, icone) -> None:
		texto_status.value = mensagem
		texto_status.color = cor
		icone_status.icon = icone
		icone_status.color = cor
		indicador_status.bgcolor = fundo

	async def executar_acao(nome: str, funcao) -> None:
		nonlocal ocupado
		if ocupado:
			return

		ocupado = True
		for botao in botoes:
			botao.disabled = True
		botao_limpar.disabled = True
		progresso.visible = True
		definir_status(f"Executando: {nome}", "#8A5A12", "#FBF0DA", ft.Icons.SYNC)
		atualizar_log(f"Iniciando {nome.lower()}...")
		page.update()

		try:
			resultado = await asyncio.to_thread(executar_com_com, funcao)
			if resultado is None:
				atualizar_log("Caixa de entrada atualizada.")
			else:
				atualizar_log(f"Processamento concluido. Anexos salvos: {resultado}")
			definir_status("Concluído", "#285D4D", "#E5F2E9", ft.Icons.CHECK_CIRCLE)
		except Exception as erro:
			atualizar_log(f"Falha na operacao: {erro}")
			definir_status("A operação encontrou um erro", "#A33B32", "#F8E8E5", ft.Icons.ERROR_OUTLINE)
		finally:
			ocupado = False
			for botao in botoes:
				botao.disabled = False
			botao_limpar.disabled = False
			progresso.visible = False
			page.update()

	async def clicar_atualizar(_evento) -> None:
		await executar_acao("atualização da caixa de entrada", atualizar_caixa_de_entrada)

	async def clicar_patrimonial(_evento) -> None:
		await executar_acao(
			"Eliminação Patrimonial",
			modulo_patrimonial.salvar_anexos_dos_relatorios,
		)

	async def clicar_resultado(_evento) -> None:
		await executar_acao(
			"Eliminação Resultado",
			modulo_resultado.salvar_anexos_dos_relatorios,
		)

	def limpar_log(_evento) -> None:
		registros.clear()
		caixa_log.value = ""
		page.update()

	botao_limpar.on_click = limpar_log

	botao_atualizar = ft.Button(
		content=ft.Row(
			[
				ft.Icon(ft.Icons.SYNC, size=20),
				ft.Text("Atualizar caixa de entrada", size=15, weight=ft.FontWeight.W_600),
			],
			alignment=ft.MainAxisAlignment.CENTER,
			spacing=10,
		),
		on_click=clicar_atualizar,
		bgcolor="#176B5B",
		color=ft.Colors.WHITE,
		height=58,
		expand=True,
	)
	botao_patrimonial = ft.Button(
		content=ft.Row(
			[
				ft.Icon(ft.Icons.ACCOUNT_BALANCE, size=21),
				ft.Text("Eliminação Patrimonial", size=15, weight=ft.FontWeight.W_600),
			],
			alignment=ft.MainAxisAlignment.CENTER,
			spacing=11,
		),
		on_click=clicar_patrimonial,
		bgcolor=ft.Colors.WHITE,
		color="#294239",
		height=72,
		expand=True,
	)
	botao_resultado = ft.Button(
		content=ft.Row(
			[
				ft.Icon(ft.Icons.ASSESSMENT, size=21),
				ft.Text("Eliminação Resultado", size=15, weight=ft.FontWeight.W_600),
			],
			alignment=ft.MainAxisAlignment.CENTER,
			spacing=11,
		),
		on_click=clicar_resultado,
		bgcolor=ft.Colors.WHITE,
		color="#294239",
		height=72,
		expand=True,
	)
	botoes.extend([botao_atualizar, botao_patrimonial, botao_resultado])

	cabecalho = ft.Row(
		[
			ft.Column(
				[
					ft.Text(
						"AUTOMAÇÃO DE PROCESSOS",
						size=11,
						weight=ft.FontWeight.W_700,
						color="#63736A",
					),
					ft.Text("Intercompany", size=28, weight=ft.FontWeight.W_700, color="#23352D"),
				],
				spacing=4,
			),
			ft.Row([progresso, indicador_status], spacing=10),
		],
		alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
		vertical_alignment=ft.CrossAxisAlignment.CENTER,
	)

	log_cabecalho = ft.Row(
		[
			ft.Column(
				[
					ft.Text("Log da execução", size=17, weight=ft.FontWeight.W_600, color="#23352D"),
					ft.Text("Resultados e mensagens das rotinas", size=12, color="#68766E"),
				],
				spacing=3,
			),
			botao_limpar,
		],
		alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
		vertical_alignment=ft.CrossAxisAlignment.CENTER,
	)

	conteudo = ft.Column(
		[
			cabecalho,
			ft.Divider(height=1, color="#DCE4DD"),
			ft.Text("Ações", size=13, weight=ft.FontWeight.W_700, color="#63736A"),
			botao_atualizar,
			ft.Row([botao_patrimonial, botao_resultado], spacing=14),
			ft.Divider(height=1, color="#DCE4DD"),
			log_cabecalho,
			ft.Container(
				content=caixa_log,
				bgcolor=ft.Colors.WHITE,
				padding=12,
				border_radius=6,
				height=340,
			),
		],
		spacing=18,
		alignment=ft.MainAxisAlignment.START,
		horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
	)
	page.add(ft.Container(content=conteudo, padding=28, expand=True))


if __name__ == "__main__":
	ft.run(main)