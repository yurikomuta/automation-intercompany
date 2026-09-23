# pip install pywin32
import os
import shutil
import time
from pathlib import Path
import win32com.client as win32

NOME_ABA = "RELATÓRIO DE ELIMINAÇÃO"


def aguardar_arquivo(path: Path, timeout: int = 30, intervalo: float = 0.5) -> None:
    tempo_inicial = time.time()
    while time.time() - tempo_inicial < timeout:
        if path.exists():
            return
        time.sleep(intervalo)
    raise FileNotFoundError(f"Arquivo não encontrado após o timeout: {path}")


def main() -> None:
    origem = Path.home() / "Downloads" / "teste.xlsx"
    destino_dir = Path.home() / "Documents"
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / "teste.xlsx"

    excel = None
    workbook = None

    try:
        if not origem.exists():
            raise FileNotFoundError(f"Arquivo de origem não encontrado: {origem}")

        if destino.exists():
            destino.unlink()

        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(origem))
        worksheet = workbook.Worksheets(1)
        worksheet.Name = NOME_ABA
        workbook.Save()
        workbook.Close(SaveChanges=True)

        shutil.move(str(origem), str(destino))
        aguardar_arquivo(destino)

        print(f"Aba renomeada para: {NOME_ABA}")
        print(f"Arquivo movido para: {destino}")

    except Exception as exc:
        print(f"Erro na automação: {exc}")
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        raise

    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()

