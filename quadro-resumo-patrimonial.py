from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZipFile
import win32com.client
from company_folder import arquivo_ja_processado, organizar_versoes
from outlook_sync import sincronizar_todas_as_pastas


PASTA_DESTINO = Path(r"\\brtssisp003\bbb0133$\2 - Intercompany")
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


def obter_codigo_empresa(empresa: str) -> str:
	"""Resolve o código da empresa usando o cadastro do processo."""
	chave = " ".join(empresa.split()).casefold()
	for nome, codigo in CODIGOS_EMPRESA.items():
		if nome.casefold() == chave:
			return codigo
	raise ValueError(f"Empresa não cadastrada no Quadro Resumo: {empresa}")


def aplicar_versao(nome: str, versao: str | None) -> str:
	"""Aplica a versão do assunto ao nome do arquivo sem duplicar sufixos."""
	arquivo = Path(nome)
	base = re.sub(r"[_ -]V\d+$", "", arquivo.stem, flags=re.IGNORECASE)
	sufixo = f"_V{versao}" if versao else ""
	return f"{base}{sufixo}{arquivo.suffix}"


def _obter_pasta_empresa(dados: dict[str, str], versao: str | None) -> tuple[Path, str]:
	codigo = obter_codigo_empresa(dados["empresa"])
	empresa = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", dados["empresa"]).strip(" .")
	pasta = (
		PASTA_DESTINO
		/ f"ANO_{dados['ano']}"
		/ f"{dados['mes']}_{dados['ano']}"
		/ "QUADRO RESUMO"
		/ f"{codigo}-{empresa}"
	)
	pasta.mkdir(parents=True, exist_ok=True)
	return pasta, aplicar_versao(Path(dados["nome"]).name, versao)


def _competencia_compativel(arquivo: dict[str, str], assunto: dict[str, str | None]) -> bool:
	return arquivo["mes"] == assunto["mes"] and arquivo["ano"] == assunto["ano"]


def _salvar_anexo_xlsx(anexo, assunto: dict[str, str | None], pastas_alteradas: set[Path]) -> int:
	dados = analisar_nome_arquivo(str(anexo.FileName))
	if not dados or not _competencia_compativel(dados, assunto):
		return 0

	pasta, nome = _obter_pasta_empresa({**dados, "nome": str(anexo.FileName)}, assunto["versao"])
	pastas_alteradas.add(pasta)
	if arquivo_ja_processado(pasta, nome):
		return 0
	anexo.SaveAsFile(str(pasta / nome))
	return 1


def _salvar_xlsx_do_zip(caminho_zip: Path, assunto: dict[str, str | None], pastas_alteradas: set[Path]) -> int:
	quantidade = 0
	with ZipFile(caminho_zip) as arquivo_zip:
		for membro in arquivo_zip.infolist():
			if membro.is_dir():
				continue
			dados = analisar_nome_arquivo(membro.filename)
			if not dados or not _competencia_compativel(dados, assunto):
				continue

			pasta, nome = _obter_pasta_empresa({**dados, "nome": membro.filename}, assunto["versao"])
			pastas_alteradas.add(pasta)
			if arquivo_ja_processado(pasta, nome):
				continue
			with arquivo_zip.open(membro) as origem, (pasta / nome).open("wb") as destino:
				shutil.copyfileobj(origem, destino)
			quantidade += 1
	return quantidade


def salvar_anexos_quadro_resumo() -> int:
	"""Salva os XLSX do Quadro Resumo por competência e empresa."""
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	sincronizar_todas_as_pastas(outlook)
	inbox = outlook.GetDefaultFolder(6)  # olFolderInbox
	itens = inbox.Items
	itens.Sort("[ReceivedTime]", True)

	quantidade = 0
	pastas_alteradas: set[Path] = set()
	for item in itens:
		if getattr(item, "Class", None) != 43:  # olMail
			continue
		assunto = analisar_assunto(str(getattr(item, "Subject", "")))
		if assunto is None:
			continue

		for anexo in item.Attachments:
			extensao = Path(str(anexo.FileName)).suffix.lower()
			if extensao == ".xlsx":
				quantidade += _salvar_anexo_xlsx(anexo, assunto, pastas_alteradas)
			elif extensao == ".zip":
				with tempfile.TemporaryDirectory() as diretorio_temporario:
					caminho_zip = Path(diretorio_temporario) / Path(str(anexo.FileName)).name
					anexo.SaveAsFile(str(caminho_zip))
					quantidade += _salvar_xlsx_do_zip(caminho_zip, assunto, pastas_alteradas)

	for pasta in pastas_alteradas:
		organizar_versoes(pasta)
	return quantidade


if __name__ == "__main__":
	print(f"Arquivos salvos: {salvar_anexos_quadro_resumo()}")
