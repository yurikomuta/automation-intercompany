"""Extracao segura de anexos compactados."""

from pathlib import Path
import shutil
from zipfile import ZipFile


def extrair_e_remover_zip(caminho_zip: Path) -> None:
	"""Extrai o ZIP na pasta atual e o remove somente apos sucesso."""
	pasta_destino = caminho_zip.parent.resolve()
	with ZipFile(caminho_zip) as arquivo_zip:
		membros = arquivo_zip.infolist()
		caminhos = []
		for membro in membros:
			caminho = (pasta_destino / membro.filename).resolve()
			if caminho != pasta_destino and pasta_destino not in caminho.parents:
				raise ValueError(f"Caminho invalido dentro do ZIP: {membro.filename}")
			if caminho == pasta_destino and not membro.is_dir():
				raise ValueError(f"Nome de arquivo invalido dentro do ZIP: {membro.filename}")
			caminhos.append(caminho)

		for membro, caminho in zip(membros, caminhos):
			if membro.is_dir():
				caminho.mkdir(parents=True, exist_ok=True)
				continue
			if caminho.exists():
				continue
			caminho.parent.mkdir(parents=True, exist_ok=True)
			with arquivo_zip.open(membro) as origem, caminho.open("wb") as destino:
				shutil.copyfileobj(origem, destino)

	caminho_zip.unlink()