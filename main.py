from zipfile import ZipFile

import captura

arquivo = "Dispositivo de controle.xlsx"

with ZipFile(arquivo) as planilha:
    dados = captura.extrair_dados(
        planilha.read("xl/worksheets/sheet3.xml"),
        planilha.read("xl/worksheets/sheet2.xml"),
        planilha.read("xl/sharedStrings.xml"),
    )

print(dados)