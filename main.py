import json
import io
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

import captura

arquivo = "Dispositivo de controle.xlsx"


def caminho_da_aba(planilha, nome):
    """Descobre o arquivo XML de uma aba pelo nome (xlsx comum ou Strict)."""
    livro = fromstring(planilha.read("xl/workbook.xml"), forbid_dtd=True)
    rels = fromstring(planilha.read("xl/_rels/workbook.xml.rels"), forbid_dtd=True)
    destinos = {r.get("Id"): r.get("Target") for r in rels.findall("{*}Relationship")}
    for aba in livro.findall(".//{*}sheet"):
        if aba.get("name") == nome:
            # o atributo r:id muda de namespace no Strict, então procura só pelo final do nome
            id_relacao = next(
                v for k, v in aba.attrib.items()
                if k.endswith("}id")
            )
            destino = destinos[id_relacao]
            return destino.lstrip("/") if destino.startswith("/") else f"xl/{destino}"
    raise ValueError(f"Aba '{nome}' não encontrada na planilha.")


with ZipFile(arquivo) as planilha:
    textos = "xl/sharedStrings.xml"
    dados = captura.extrair_dados(
        planilha.read(caminho_da_aba(planilha, "Registro")),
        planilha.read(caminho_da_aba(planilha, "Incerteza")),
        planilha.read(caminho_da_aba(planilha, "Certificado")),
        planilha.read(textos) if textos in planilha.namelist() else None,
    )

print(json.dumps(dados, ensure_ascii=False, indent=2))
