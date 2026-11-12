from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZipFile
import win32com.client
from outlook_sync import sincronizar_todas_as_pastas


PASTA_DESTINO = Path(r"C:\Users\ykomuta\Documents")
PASTA_VERSOES_ANTERIORES = "Versões Anteriores"
ASSUNTO = re.compile(
	r"^QUADRO RESUMO\s*[-–]\s*"
	r"(?P<mes>0[1-9]|1[0-2])/(?P<ano>\d{4})"
	r"(?:\s*_V(?P<versao>\d+))?$",
	re.IGNORECASE,
)
ARQUIVO_QUADRO = re.compile(
	r"^Quadro Resumo\s*-\s*(?P<empresa>.+?)\s*-\s*"
	r"(?P<mes>0?[1-9]|1[0-2])_(?P<ano>\d{4})\.xlsx$",
	re.IGNORECASE,
)
CODIGOS_EMPRESA = {
	"TBRA": "0185",
	"TIS": "0306",
	"TERRA": "0367",
	"TGLOG": "0999",
	"POP": "1871",
	"RECICLA V": "1872",
	"CLOUD": "2008",
	"FIBRASIL": "2011",
	"VSS": "2207",
	"VSCD": "2234",
}


class ListaArquivosSalvos(list):
	"""Lista de arquivos salvos com formatação legível para o log."""

	def __str__(self) -> str:
		if not self:
			return "0\nNenhum arquivo novo foi salvo."
		linhas = "\n".join(f"- {arquivo}" for arquivo in self)
		return f"{len(self)}\n{linhas}"


def analisar_assunto(assunto: str) -> dict[str, str | None] | None:
	"""Retorna competência e versão do assunto do e-mail, se reconhecido."""
	correspondencia = ASSUNTO.fullmatch(assunto.strip())
	return correspondencia.groupdict() if correspondencia else None


def analisar_nome_arquivo(nome: str) -> dict[str, str] | None:
	"""Extrai empresa e competência do nome do XLSX."""
	correspondencia = ARQUIVO_QUADRO.fullmatch(Path(nome).name)
	if not correspondencia:
		return None
	dados = correspondencia.groupdict()
	dados["mes"] = dados["mes"].zfill(2)
	return dados


def obter_codigo_empresa(empresa: str) -> str | None:
	"""Resolve o código da empresa ou retorna None se ela não estiver cadastrada."""
	chave = " ".join(empresa.split()).casefold()
	for nome, codigo in CODIGOS_EMPRESA.items():
		if nome.casefold() == chave:
			return codigo
	return None


def aplicar_versao(nome: str, versao: str | None) -> str:
	"""Aplica a versão do assunto ao nome do arquivo sem duplicar sufixos."""
	arquivo = Path(nome)
	base = re.sub(r"[_ -]V\d+$", "", arquivo.stem, flags=re.IGNORECASE)
	sufixo = f"_V{versao}" if versao else ""
	return f"{base}{sufixo}{arquivo.suffix}"


def _remover_subpasta_versoes_anteriores(pasta_empresa: Path) -> None:
	"""Restaura arquivos de 'Versões Anteriores' para a pasta da empresa e remove a subpasta."""
	pasta_anteriores = pasta_empresa / PASTA_VERSOES_ANTERIORES
	if not pasta_anteriores.is_dir():
		return
	for item in pasta_anteriores.iterdir():
		destino = pasta_empresa / item.name
		if item.is_file() and not destino.exists():
			shutil.move(str(item), str(destino))
	shutil.rmtree(pasta_anteriores, ignore_errors=True)


def _obter_pasta_empresa(
	dados: dict[str, str], versao: str | None
) -> tuple[Path, str] | None:
	codigo = obter_codigo_empresa(dados["empresa"])
	if codigo is None:
		return None
	empresa = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", dados["empresa"]).strip(" .")
	pasta = (
		PASTA_DESTINO
		/ f"ANO_{dados['ano']}"
		/ f"{dados['mes']}_{dados['ano']}"
		/ "QUADRO RESUMO"
		/ f"{codigo}-{empresa}"
	)
	pasta.mkdir(parents=True, exist_ok=True)
	_remover_subpasta_versoes_anteriores(pasta)
	return pasta, aplicar_versao(Path(dados["nome"]).name, versao)


def _competencia_compativel(arquivo: dict[str, str], assunto: dict[str, str | None]) -> bool:
	return arquivo["mes"] == assunto["mes"] and arquivo["ano"] == assunto["ano"]


def _salvar_anexo_xlsx(
	anexo,
	assunto: dict[str, str | None],
	destinos_salvos: set[Path] | None = None,
) -> list[str]:
	dados = analisar_nome_arquivo(str(anexo.FileName))
	if not dados or not _competencia_compativel(dados, assunto):
		return []

	resultado = _obter_pasta_empresa(
		{**dados, "nome": str(anexo.FileName)}, assunto["versao"]
	)
	if resultado is None:
		return []
	pasta, nome = resultado
	destino = pasta / nome
	if destinos_salvos is not None and destino in destinos_salvos:
		return []
	anexo.SaveAsFile(str(destino))
	if destinos_salvos is not None:
		destinos_salvos.add(destino)
	return [str(destino.relative_to(PASTA_DESTINO))]


def _salvar_xlsx_do_zip(
	caminho_zip: Path,
	assunto: dict[str, str | None],
	destinos_salvos: set[Path] | None = None,
) -> list[str]:
	arquivos_salvos = ListaArquivosSalvos()
	with ZipFile(caminho_zip) as arquivo_zip:
		for membro in arquivo_zip.infolist():
			if membro.is_dir():
				continue
			dados = analisar_nome_arquivo(membro.filename)
			if not dados or not _competencia_compativel(dados, assunto):
				continue

			resultado = _obter_pasta_empresa(
				{**dados, "nome": membro.filename}, assunto["versao"]
			)
			if resultado is None:
				continue
			pasta, nome = resultado
			destino = pasta / nome
			if destinos_salvos is not None and destino in destinos_salvos:
				continue
			with arquivo_zip.open(membro) as origem, destino.open("wb") as arquivo_destino:
				shutil.copyfileobj(origem, arquivo_destino)
			if destinos_salvos is not None:
				destinos_salvos.add(destino)
			arquivos_salvos.append(str(destino.relative_to(PASTA_DESTINO)))
	return arquivos_salvos


def salvar_anexos_quadro_resumo() -> list[str]:
	"""Salva os XLSX e retorna os caminhos relativos dos arquivos gravados."""
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	sincronizar_todas_as_pastas(outlook)
	inbox = outlook.GetDefaultFolder(6)  # olFolderInbox
	itens = inbox.Items
	itens.Sort("[ReceivedTime]", True)

	arquivos_salvos = ListaArquivosSalvos()
	destinos_salvos: set[Path] = set()
	for item in itens:
		if getattr(item, "Class", None) != 43:  # olMail
			continue
		assunto = analisar_assunto(str(getattr(item, "Subject", "")))
		if assunto is None:
			continue

		for anexo in item.Attachments:
			extensao = Path(str(anexo.FileName)).suffix.lower()
			if extensao == ".xlsx":
				arquivos_salvos.extend(
					_salvar_anexo_xlsx(anexo, assunto, destinos_salvos)
				)
			elif extensao == ".zip":
				with tempfile.TemporaryDirectory() as diretorio_temporario:
					caminho_zip = Path(diretorio_temporario) / Path(str(anexo.FileName)).name
					anexo.SaveAsFile(str(caminho_zip))
					arquivos_salvos.extend(
						_salvar_xlsx_do_zip(caminho_zip, assunto, destinos_salvos)
					)

	return arquivos_salvos


if __name__ == "__main__":
	arquivos_salvos = salvar_anexos_quadro_resumo()
	print(f"Arquivos salvos: {len(arquivos_salvos)}")
	for arquivo in arquivos_salvos:
		print(f"- {arquivo}")
