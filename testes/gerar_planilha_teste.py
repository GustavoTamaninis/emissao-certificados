"""Gera uma cópia da planilha real com 4 linhas de medição preenchidas manualmente,
simulando o que um técnico da Schulz digitaria no dia a dia. Só o conteúdo das
células é trocado (mesma técnica de edição direta de XML que o captura.py usa pra
ler); nenhuma biblioteca de planilha (openpyxl etc.) é usada.
"""
import re
import zipfile
from pathlib import Path

# Planilha-modelo real da Schulz, usada como base (caminho de quem gerou o teste
# originalmente). Ajuste para o caminho local se for regerar a planilha de teste.
ORIGINAL = r"C:\Users\Furla\Downloads\Dispositivo de controle (2) (1).xlsx"
DESTINO = Path(__file__).resolve().parent / "planilha_teste_preenchida.xlsx"

# parametro_avaliado, valor_nominal, tol_mais, tol_menos, medida_1, medida_2
LINHAS_TESTE = {
    30: ("Diâmetro A", 10.0000, 0.0100, -0.0100, 10.0020, 10.0030),
    31: ("Posição do furo A", 5.0000, 0.1000, 0.0000, 5.0400, 5.0600),
    32: ("Posição do rebaixo B", 8.0000, 0.1000, 0.0000, 7.9400, 7.9600),
    33: ("Diâmetro B", 20.0000, 0.0050, -0.0050, 20.0120, 20.0140),
}


def montar_linha(numero, dados):
    nome, valor_nominal, tol_mais, tol_menos, medida_1, medida_2 = dados
    return (
        f'<row r="{numero}">'
        f'<c r="B{numero}" t="str"><v>{nome}</v></c>'
        f'<c r="C{numero}"><v>{valor_nominal}</v></c>'
        f'<c r="D{numero}"><v>{tol_mais}</v></c>'
        f'<c r="E{numero}"><v>{tol_menos}</v></c>'
        f'<c r="F{numero}"><v>{medida_1}</v></c>'
        f'<c r="G{numero}"><v>{medida_2}</v></c>'
        f'</row>'
    )


with zipfile.ZipFile(ORIGINAL) as origem:
    caminho_sheet1 = "xl/worksheets/sheet1.xml"
    xml_texto = origem.read(caminho_sheet1).decode("utf-8")

    for numero, dados in LINHAS_TESTE.items():
        padrao = re.compile(rf'<row r="{numero}"[^>]*>.*?</row>', re.DOTALL)
        nova_linha = montar_linha(numero, dados)
        xml_texto, trocas = padrao.subn(nova_linha, xml_texto)
        assert trocas == 1, f"linha {numero} não encontrada/substituída ({trocas} trocas)"

    with zipfile.ZipFile(DESTINO, "w", zipfile.ZIP_DEFLATED) as saida:
        for item in origem.infolist():
            conteudo = xml_texto.encode("utf-8") if item.filename == caminho_sheet1 else origem.read(item.filename)
            saida.writestr(item, conteudo)

print("planilha de teste gerada em:", DESTINO)
