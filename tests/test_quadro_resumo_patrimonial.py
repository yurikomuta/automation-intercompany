import importlib.util
from pathlib import Path
import sys
import tempfile
import types
import unittest
from zipfile import ZipFile

if "pythoncom" not in sys.modules:
	sys.modules["pythoncom"] = types.ModuleType("pythoncom")
if "win32com" not in sys.modules:
	win32com_mod = types.ModuleType("win32com")
	win32com_mod.client = types.ModuleType("win32com.client")
	sys.modules["win32com"] = win32com_mod
	sys.modules["win32com.client"] = win32com_mod.client

PROJETO_DIR = Path(__file__).resolve().parents[1]
if str(PROJETO_DIR) not in sys.path:
	sys.path.insert(0, str(PROJETO_DIR))

ARQUIVO_MODULO = PROJETO_DIR / "quadro-resumo-patrimonial.py"
ESPECIFICACAO = importlib.util.spec_from_file_location("quadro_resumo_patrimonial", ARQUIVO_MODULO)
QUADRO_RESUMO = importlib.util.module_from_spec(ESPECIFICACAO)
ESPECIFICACAO.loader.exec_module(QUADRO_RESUMO)


class QuadroResumoTests(unittest.TestCase):
	def setUp(self):
		self.pasta_destino_original = QUADRO_RESUMO.PASTA_DESTINO

	def tearDown(self):
		QUADRO_RESUMO.PASTA_DESTINO = self.pasta_destino_original

	def test_destino_e_pasta_local_de_teste(self):
		self.assertEqual(
			QUADRO_RESUMO.PASTA_DESTINO,
			Path(r"C:\Users\ykomuta\Documents"),
		)

	def test_analisar_assunto_com_versao(self):
		self.assertEqual(
			QUADRO_RESUMO.analisar_assunto("QUADRO RESUMO - 03/2026_V1"),
			{"mes": "03", "ano": "2026", "versao": "1"},
		)

	def test_analisar_assunto_sem_versao(self):
		self.assertEqual(
			QUADRO_RESUMO.analisar_assunto("QUADRO RESUMO - 08/2026"),
			{"mes": "08", "ano": "2026", "versao": None},
		)

	def test_analisar_nome_arquivo_normaliza_mes(self):
		self.assertEqual(
			QUADRO_RESUMO.analisar_nome_arquivo(
				"Quadro Resumo - TBRA - 3_2026.xlsx"
			),
			{"empresa": "TBRA", "mes": "03", "ano": "2026"},
		)

	def test_resolve_codigo_empresa(self):
		self.assertEqual(QUADRO_RESUMO.obter_codigo_empresa("recicla v"), "1872")

	def test_empresa_nao_cadastrada_nao_tem_codigo(self):
		self.assertIsNone(QUADRO_RESUMO.obter_codigo_empresa("EMPRESA NOVA"))

	def test_aplicar_versao_substitui_sufixo_existente(self):
		self.assertEqual(
			QUADRO_RESUMO.aplicar_versao("Quadro Resumo - TBRA - 03_2026_V1.xlsx", "2"),
			"Quadro Resumo - TBRA - 03_2026_V2.xlsx",
		)

	def test_zip_encaminha_xlsx_para_pastas_de_empresa_sem_subpastas(self):
		with tempfile.TemporaryDirectory() as diretorio:
			QUADRO_RESUMO.PASTA_DESTINO = Path(diretorio)
			caminho_zip = Path(diretorio) / "quadro.zip"
			with ZipFile(caminho_zip, "w") as arquivo_zip:
				arquivo_zip.writestr("Quadro Resumo - TBRA - 03_2026.xlsx", b"tbra")
				arquivo_zip.writestr("subpasta/Quadro Resumo - VSS - 03_2026.xlsx", b"vss")
				arquivo_zip.writestr("Quadro Resumo - EMPRESA NOVA - 03_2026.xlsx", b"ignorar")
				arquivo_zip.writestr("Quadro Resumo - TIS - 02_2026.xlsx", b"outra competencia")

			assunto = {"mes": "03", "ano": "2026", "versao": "1"}
			destinos_salvos = set()
			arquivos_salvos = QUADRO_RESUMO._salvar_xlsx_do_zip(
				caminho_zip, assunto, destinos_salvos
			)

			pasta_tbra = Path(diretorio) / "ANO_2026" / "03_2026" / "QUADRO RESUMO" / "0185-TBRA"
			pasta_vss = Path(diretorio) / "ANO_2026" / "03_2026" / "QUADRO RESUMO" / "2207-VSS"
			self.assertEqual(
				[path.replace("\\", "/") for path in arquivos_salvos],
				[
					"ANO_2026/03_2026/QUADRO RESUMO/0185-TBRA/Quadro Resumo - TBRA - 03_2026_V1.xlsx",
					"ANO_2026/03_2026/QUADRO RESUMO/2207-VSS/Quadro Resumo - VSS - 03_2026_V1.xlsx",
				],
			)
			self.assertEqual((pasta_tbra / "Quadro Resumo - TBRA - 03_2026_V1.xlsx").read_bytes(), b"tbra")
			self.assertEqual((pasta_vss / "Quadro Resumo - VSS - 03_2026_V1.xlsx").read_bytes(), b"vss")
			self.assertEqual([item for item in pasta_tbra.iterdir() if item.is_dir()], [])
			self.assertEqual([item for item in pasta_vss.iterdir() if item.is_dir()], [])
			self.assertEqual(len(destinos_salvos), 2)
			self.assertIn("Quadro Resumo - TBRA - 03_2026_V1.xlsx", str(arquivos_salvos))

	def test_remove_subpasta_versoes_anteriores_legada(self):
		with tempfile.TemporaryDirectory() as diretorio:
			QUADRO_RESUMO.PASTA_DESTINO = Path(diretorio)
			pasta_legada = (
				Path(diretorio)
				/ "ANO_2026"
				/ "03_2026"
				/ "QUADRO RESUMO"
				/ "0185-TBRA"
				/ "Versões Anteriores"
			)
			pasta_legada.mkdir(parents=True)
			(pasta_legada / "Quadro Resumo - TBRA - 03_2026_V1.xlsx").write_bytes(b"v1")

			pasta, nome = QUADRO_RESUMO._obter_pasta_empresa(
				{"empresa": "TBRA", "mes": "03", "ano": "2026", "nome": "Quadro Resumo - TBRA - 03_2026.xlsx"},
				"2",
			)
			self.assertEqual(nome, "Quadro Resumo - TBRA - 03_2026_V2.xlsx")
			self.assertFalse(pasta_legada.exists())
			self.assertTrue((pasta / "Quadro Resumo - TBRA - 03_2026_V1.xlsx").is_file())


if __name__ == "__main__":
	unittest.main()