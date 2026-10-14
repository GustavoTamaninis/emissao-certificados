import json
import sys
from pathlib import Path
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import calculos
import captura
import certificado

PASTA_TESTES = Path(__file__).resolve().parent


def caminho_da_aba(planilha, nome):
    livro = fromstring(planilha.read("xl/workbook.xml"), forbid_dtd=True)
    rels = fromstring(planilha.read("xl/_rels/workbook.xml.rels"), forbid_dtd=True)
    destinos = {r.get("Id"): r.get("Target") for r in rels.findall("{*}Relationship")}
    for aba in livro.findall(".//{*}sheet"):
        if aba.get("name") == nome:
            id_relacao = next(v for k, v in aba.attrib.items() if k.endswith("}id"))
            destino = destinos[id_relacao]
            return destino.lstrip("/") if destino.startswith("/") else f"xl/{destino}"
    raise ValueError(f"Aba '{nome}' não encontrada na planilha.")


caminho = PASTA_TESTES / "planilha_teste_preenchida.xlsx"

with ZipFile(caminho) as planilha:
    textos = "xl/sharedStrings.xml"
    dados = captura.extrair_dados(
        planilha.read(caminho_da_aba(planilha, "Registro")),
        planilha.read(caminho_da_aba(planilha, "Incerteza")),
        planilha.read(caminho_da_aba(planilha, "Certificado")),
        planilha.read(textos) if textos in planilha.namelist() else None,
    )

resultados, parecer = calculos.calcular_resultados(dados["parametros"])

saida = {
    "dados_extraidos": {k: v for k, v in dados.items() if k != "parametros"},
    "parametros_lidos": dados["parametros"],
    "resultados_calculados": resultados,
    "parecer_geral": parecer,
}

with open(PASTA_TESTES / "resultado_teste.json", "w", encoding="utf-8") as f:
    json.dump(saida, f, ensure_ascii=False, indent=2, default=str)

print(json.dumps(saida, ensure_ascii=False, indent=2, default=str))

pdf_bytes = certificado.gerar_certificado_pdf(dados, resultados, parecer, "TESTE-001/2026")
with open(PASTA_TESTES / "certificado_teste.pdf", "wb") as f:
    f.write(pdf_bytes)
print("\nPDF gerado em testes/certificado_teste.pdf,", len(pdf_bytes), "bytes")
