"""Organizacao das pastas e versoes dos arquivos das empresas."""

from pathlib import Path
import re
import shutil


PASTA_CHECKLIST = "Checklist"
PASTA_VERSOES_ANTERIORES = "Versões Anteriores"
VERSAO_ARQUIVO = re.compile(r"^(?P<base>.+?)[ _-]+V(?P<versao>\d+)$", re.IGNORECASE)


def arquivo_ja_processado(pasta_empresa: Path, nome: str) -> bool:
	"""Informa se o anexo ja esta na raiz ou no arquivo de versoes."""
	return any(
		(pasta / nome).is_file()
		for pasta in (
			pasta_empresa,
			pasta_empresa / PASTA_VERSOES_ANTERIORES,
		)
	)


def preparar_pasta_empresa(pasta_empresa: Path) -> None:
	"""Cria as subpastas da empresa e organiza as versoes encontradas."""
	pasta_empresa.mkdir(parents=True, exist_ok=True)
	(pasta_empresa / PASTA_CHECKLIST).mkdir(exist_ok=True)
	organizar_versoes(pasta_empresa)


def organizar_versoes(pasta_empresa: Path) -> None:
	"""Mantem a maior versao na raiz e arquiva versoes numericamente menores."""
	pasta_anteriores = pasta_empresa / PASTA_VERSOES_ANTERIORES
	pasta_anteriores.mkdir(exist_ok=True)
	grupos = {}

	for pasta in (pasta_empresa, pasta_anteriores):
		for arquivo in pasta.iterdir():
			if not arquivo.is_file():
				continue
			correspondencia = VERSAO_ARQUIVO.fullmatch(arquivo.stem)
			base = correspondencia.group("base") if correspondencia else arquivo.stem
			versao = int(correspondencia.group("versao")) if correspondencia else 0
			chave = (base.casefold(), arquivo.suffix.casefold())
			grupos.setdefault(chave, []).append((versao, arquivo))

	for arquivos in grupos.values():
		maior_versao = max(versao for versao, _ in arquivos)
		mais_novos = [arquivo for versao, arquivo in arquivos if versao == maior_versao]
		mais_recente = next(
			(arquivo for arquivo in mais_novos if arquivo.parent == pasta_empresa),
			mais_novos[0],
		)
		if mais_recente.parent == pasta_anteriores:
			shutil.move(str(mais_recente), str(pasta_empresa / mais_recente.name))

		for versao, arquivo in arquivos:
			if versao >= maior_versao or arquivo.parent != pasta_empresa:
				continue
			destino = pasta_anteriores / arquivo.name
			contador = 1
			while destino.exists():
				destino = pasta_anteriores / f"{arquivo.stem}_{contador}{arquivo.suffix}"
				contador += 1
			shutil.move(str(arquivo), str(destino))