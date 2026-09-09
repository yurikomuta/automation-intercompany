#pip install pywin32

import win32com.client as win32
import shutil
import time 
import os

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

#1. acessa o arquivo excel na pasta de downloads 
origem = "C:\\Users\\ykomuta\\Downloads\\teste.xlsx"

#2. move o arquivo para a pasta de documentos e abre o arquivo
destino = "C:\\Users\\ykomuta\\Documents\\teste.xlsx"
shutil.move(origem, destino)
workbook = excel.Workbooks.Open(destino)
time.sleep(10)

#2.1 renomeia a planilha para "teste2"
workbook.Worksheets(1).Name = "teste2"
time.sleep(5)

#2.2 salva o arquivo
workbook.Save()

#2.3 renomeia o excel para "movido"


#3. fecha o arquivo excel
excel.Quit()

