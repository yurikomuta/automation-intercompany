from pathlib import Path
import re
import win32com.client
from attachment_archive import extrair_e_remover_zip
from outlook_sync import sincronizar_todas_as_pastas

"""Localiza relatórios de eliminação resultado no Outlook ."""
ASSUNTO = re.compile(
	r"^RELAT[ÓO]RIO DE ELIMINA[ÇC][ÃA]O RESULTADO(?:\s+V\d+)?"
	r"\s*[–-]\s*"
    r"(?P<empresa>.+?)\s+\((?P<codigo1>[^()]+)\)\s+x\s+"
	r"\((?P<codigo2>[^()]+)\)\s+(?P<empresa2>.+?)\s*[–_-]\s*"
	r"(?P<mes>0[1-9]|1[0-2])(?:[/_-])(?P<ano>\d{4})"
	r"(?:\s*[_-]\s*V\d+)?(?:\s+.*)?$",
    re.IGNORECASE,
)
PASTA_DESTINO = Path(r"C:\Users\ykomuta\Documents\2 - Intercompany")
EXTENSOES_PERMITIDAS = {".xls", ".xlsx", ".csv", ".xlsm",".zip"}
COMPETENCIA_ANEXO = re.compile(
	r"(?<!\d)(?P<mes>0?[1-9]|1[0-2])[_/-](?P<ano>\d{4})(?!\d)"
)


def obter_competencia_anexo(nome: str) -> tuple[str, str] | None:
	"""Extrai mês e ano do nome do anexo, por exemplo 8_2026 ou 08-2026."""
	correspondencia = COMPETENCIA_ANEXO.search(nome)
	if not correspondencia:
		return None

	mes = correspondencia.group("mes").zfill(2)
	ano = correspondencia.group("ano")
	return ano, mes


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
		empresa2 = dados["empresa2"].strip()
		pasta_empresa = (
			f"{dados['codigo1'].strip()} x {dados['codigo2'].strip()} - "
			f"{empresa} x {empresa2}"
		)

		for anexo in item.Attachments:
			nome = Path(str(anexo.FileName)).name
			if Path(nome).suffix.lower() not in EXTENSOES_PERMITIDAS:
				continue
				"""realiza a extração do mês e ano do nome do anexo, para criar a pasta de destino correta. Se não for possível extrair, o anexo é ignorado."""
			competencia = obter_competencia_anexo(nome)
			if competencia is None:
				continue
			ano, mes = competencia
			pasta = (
				PASTA_DESTINO
				/ f"ANO_{ano}"
				/ f"{mes}_{ano}"
				/ "APLICATIVO ELIMINAÇÃO"
				/ "Eliminação Resultado"
				/ pasta_empresa
			)
			pasta.mkdir(parents=True, exist_ok=True)

			destino = pasta / nome
			if destino.exists():
				if destino.suffix.lower() == ".zip":
					extrair_e_remover_zip(destino)
				continue

			anexo.SaveAsFile(str(destino))
			if destino.suffix.lower() == ".zip":
				extrair_e_remover_zip(destino)
			quantidade += 1

	return quantidade


if __name__ == "__main__":
	print(f"Anexos salvos: {salvar_anexos_dos_relatorios()}")
