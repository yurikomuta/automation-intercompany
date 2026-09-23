"""Localiza relatórios no Outlook e salva seus anexos."""

from pathlib import Path
import re

import win32com.client


ASSUNTO = re.compile(
	r"^RELATÓRIO DE ELIMINAÇÃO RESULTADO\s*[–-]\s*"
	r"(?P<empresa>.+?)\s+\((?P<codigo1>[^()]+)\)\s+x\s+"
	r"\((?P<codigo2>[^()]+)\)\s+(?P<empresa2>.+?)\s*[–-]\s*"
	r"(?P<mes>0[1-9]|1[0-2])/(?P<ano>\d{4})$",
	re.IGNORECASE,
)
PASTA_DESTINO = Path(r"C:\Users\ykomuta\Documents")
EXTENSOES_PERMITIDAS = {".xls", ".xlsx", ".csv", ".xlsm"}


def salvar_anexos_dos_relatorios() -> int:
	"""Processa os e-mails da Caixa de Entrada e retorna a quantidade salva."""
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	caixa_de_entrada = outlook.GetDefaultFolder(6)  # olFolderInbox
	itens = caixa_de_entrada.Items
	itens.Sort("[ReceivedTime]", True)

	quantidade = 0
	for item in itens:
		if getattr(item, "Class", None) != 43:  # olMail
			continue

		assunto = str(getattr(item, "Subject", "")).strip()
		correspondencia = ASSUNTO.fullmatch(assunto)
		if not correspondencia:
			continue

		dados = correspondencia.groupdict()
		empresa = dados["empresa"].strip()
		empresa2 = dados["empresa2"].strip()
		pasta = PASTA_DESTINO / (
			f"({dados['codigo1'].strip()}) x ({dados['codigo2'].strip()}) "
			f"{empresa} x {empresa2}"
		)
		pasta.mkdir(parents=True, exist_ok=True)

		for anexo in item.Attachments:
			nome = Path(str(anexo.FileName)).name
			if Path(nome).suffix.lower() not in EXTENSOES_PERMITIDAS:
				continue
			destino = pasta / nome
			contador = 1
			while destino.exists():
				destino = pasta / f"{Path(nome).stem}_{contador}{Path(nome).suffix}"
				contador += 1
			anexo.SaveAsFile(str(destino))
			quantidade += 1

	return quantidade


if __name__ == "__main__":
	print(f"Anexos salvos: {salvar_anexos_dos_relatorios()}")
