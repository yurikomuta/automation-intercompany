from pathlib import Path
import re
import win32com.client
from attachment_archive import extrair_e_remover_zip
from company_folder import (
	arquivo_ja_processado,
	organizar_versoes,
	preparar_pasta_empresa,
)
from outlook_sync import sincronizar_todas_as_pastas

"""Localiza assuntos de composição de saldo patrimonial ou resultado no Outlook."""
ASSUNTO = re.compile(
	r"^COMPOSI[ÇC][ÃA]O DE SALDO (?:PATRIMONIAL|RESULTADO)(?:\s+V\d+)?"
	r"\s*[–-]\s*"
	r"(?P<empresa>.+?)\s+\((?P<codigo>[^()]+)\)\s*[–-]\s*"
	r"(?P<mes>0[1-9]|1[0-2])(?:[/_-])(?P<ano>\d{4})"
	r"(?:\s*[_-]\s*V\d+)?(?:\s+.*)?$",
	re.IGNORECASE,
)
PASTA_DESTINO = Path(r"\\brtssisp003\bbb0133$\2 - Intercompany")
EXTENSOES_PERMITIDAS = {".xls", ".xlsx", ".csv", ".xlsm",".zip"}


def salvar_anexos_dos_relatorios() -> int:
	"""Processa os e-mails da Caixa de Entrada e retorna a quantidade salva."""
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	sincronizar_todas_as_pastas(outlook)
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
		pasta_empresa = f"{dados['codigo'].strip()}-{empresa}"

		for anexo in item.Attachments:
			nome = Path(str(anexo.FileName)).name
			if Path(nome).suffix.lower() not in EXTENSOES_PERMITIDAS:
				continue

			pasta = (
				PASTA_DESTINO
				/ f"ANO_{dados['ano']}"
				/ f"{dados['mes']}_{dados['ano']}"
				/ "COMPOSIÇÃO DE SALDOS"
				/ pasta_empresa
			)
			preparar_pasta_empresa(pasta)

			destino = pasta / nome
			if destino.exists():
				if destino.suffix.lower() == ".zip":
					extrair_e_remover_zip(destino)
					organizar_versoes(pasta)
				continue
			if arquivo_ja_processado(pasta, nome):
				continue

			anexo.SaveAsFile(str(destino))
			if destino.suffix.lower() == ".zip":
				extrair_e_remover_zip(destino)
			organizar_versoes(pasta)
			quantidade += 1

	return quantidade


if __name__ == "__main__":
	print(f"Anexos salvos: {salvar_anexos_dos_relatorios()}")
