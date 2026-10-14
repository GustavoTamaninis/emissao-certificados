"""Gera duas planilhas de teste (aprovado e reprovado) a partir do modelo real da
Schulz, editando célula por célula no XML interno do .xlsx — sem usar openpyxl e
sem substituir a linha inteira (isso apagaria estilo e fórmulas das colunas que
não preenchemos).

Regra sobre fórmulas (tag <f>): removemos a fórmula quando ela busca dado de uma
aba que não vamos preencher (Importação, Complemento) — senão o Excel recalcula e
some com o nosso valor. Mantemos a fórmula quando ela é calculada a partir de
células que nós mesmos preenchemos (colunas H/I do Registro, e a célula de
dispersão na aba Incerteza, que é MAX(Registro!I30:I64)*1000) — assim o arquivo
continua coerente se alguém abrir no Excel.
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from defusedxml.ElementTree import fromstring

ORIGINAL = r"C:\Users\Furla\Downloads\Dispositivo de controle (2) (1).xlsx"
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


def _celula(xml, endereco):
    """Casa a célula inteira (self-closing ou com filhos), pra poder trocar só o conteúdo."""
    padrao = re.compile(rf'<c r="{endereco}"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
    encontrada = padrao.search(xml)
    if encontrada is None:
        raise ValueError(f"Célula {endereco} não encontrada no XML.")
    return padrao, encontrada.group(0)


def _atributos_sem_tipo(celula_xml):
    """Atributos da abertura da célula (ex.: s="33"), sem o atributo t= antigo."""
    abertura = re.match(r'<c r="[^"]*"([^>]*?)/?>', celula_xml).group(1)
    return re.sub(r'\s*t="[^"]*"', "", abertura)


def _formula(celula_xml, manter_formula):
    """Fórmula original (<f>...</f> ou <f .../>), preservada só quando manter_formula=True."""
    if not manter_formula:
        return ""
    encontrada = re.search(r"<f[^>]*?(?:/>|>.*?</f>)", celula_xml, re.DOTALL)
    return encontrada.group(0) if encontrada else ""


def definir_numero(xml, endereco, valor, manter_formula=False):
    padrao, celula_atual = _celula(xml, endereco)
    atributos = _atributos_sem_tipo(celula_atual)
    formula = _formula(celula_atual, manter_formula)
    nova_celula = f'<c r="{endereco}"{atributos}>{formula}<v>{valor}</v></c>'
    return padrao.sub(nova_celula, xml, count=1)


def definir_texto(xml, endereco, texto, manter_formula=False):
    padrao, celula_atual = _celula(xml, endereco)
    atributos = _atributos_sem_tipo(celula_atual)
    formula = _formula(celula_atual, manter_formula)
    nova_celula = f'<c r="{endereco}" t="inlineStr"{atributos}>{formula}<is><t>{escape(texto)}</t></is></c>'
    return padrao.sub(nova_celula, xml, count=1)


# cabeçalho da aba Registro, igual nos dois arquivos de teste
CABECALHO_REGISTRO = {
    "C5": ("2026-10-14T09:30:00", definir_texto),
    "H5": ("20,3°", definir_texto),
    "C7": ("Usinagem II", definir_texto),
    "C9": ("Marcos Vinícius Pereira", definir_texto),
    "C12": ("1", definir_numero),
    "C18": ("703.2045", definir_texto),
    "C20": ("Dispositivo de controle", definir_texto),
    "C22": ("Exclusivo", definir_texto),
    "C24": ("Não consta", definir_texto),
}

# linhas 30 a 35 da tabela de resultados: parametro, nominal, tol+, tol-, medida1, medida2, media, desvio
LINHAS_COMUNS = {
    30: ("PLANICIDADE_REF", "0", "0.0500", "0.0000", "0.0021", "0.0023", "0.0022", "0.00014142"),
    31: ("25P_DIÂMETRO 1", "25.0000", "0.0100", "-0.0100", "25.0043", "25.0051", "25.0047", "0.00056569"),
    32: ("25P_DIÂMETRO 2", "25.0000", "0.0100", "-0.0100", "24.9962", "24.9958", "24.9960", "0.00028284"),
    34: ("40P_POSIÇÃO 2", "0", "0.1000", "0.0000", "0.0443", "0.0457", "0.0450", "0.00098995"),
}

CENARIOS = {
    "planilha_teste_aprovado.xlsx": {
        "linhas": {
            **LINHAS_COMUNS,
            33: ("40P_POSIÇÃO 1", "0", "0.1000", "0.0000", "0.0612", "0.0628", "0.0620", "0.00113137"),
            35: ("12P_DIÂMETRO 1", "12.0000", "0.0080", "-0.0080", "12.0031", "12.0035", "12.0033", "0.00028284"),
        },
        "dispersao": "1.13137",
    },
    "planilha_teste_reprovado.xlsx": {
        "linhas": {
            **LINHAS_COMUNS,
            33: ("40P_POSIÇÃO 1", "0", "0.1000", "0.0000", "0.1380", "0.1420", "0.1400", "0.00282843"),
            35: ("12P_DIÂMETRO 1", "12.0000", "0.0080", "-0.0080", "12.0105", "12.0115", "12.0110", "0.00070711"),
        },
        "dispersao": "2.82843",
    },
}

# só pra o PDF de demonstração não sair com "----" e "???"; iguais nos dois arquivos
CERTIFICADO_TEXTOS = {
    "F6": "00142/2026",
    "A21": "Certificado de Calibração n° 251014502390 - Carl Zeiss; de 14/10/2025 - Próxima Calibração: 14/10/2027",
    "D26": "Calibração de dispositivo de controle - ITL-0026 (revisão 2).",
}


def montar_registro(xml, linhas):
    for endereco, (valor, escrever) in CABECALHO_REGISTRO.items():
        xml = escrever(xml, endereco, valor)

    for linha, (parametro, nominal, tol_mais, tol_menos, medida_1, medida_2, media, desvio) in linhas.items():
        xml = definir_texto(xml, f"B{linha}", parametro)
        xml = definir_numero(xml, f"C{linha}", nominal)
        xml = definir_numero(xml, f"D{linha}", tol_mais)
        xml = definir_numero(xml, f"E{linha}", tol_menos)
        xml = definir_numero(xml, f"F{linha}", medida_1)
        xml = definir_numero(xml, f"G{linha}", medida_2)
        # H (média) e I (desvio padrão) são calculadas a partir de F e G, que nós
        # preenchemos, então mantemos a fórmula e só atualizamos o valor em cache
        xml = definir_numero(xml, f"H{linha}", media, manter_formula=True)
        xml = definir_numero(xml, f"I{linha}", desvio, manter_formula=True)
    return xml


def montar_certificado(xml):
    for endereco, valor in CERTIFICADO_TEXTOS.items():
        xml = definir_texto(xml, endereco, valor)
    return xml


def montar_incerteza(xml, dispersao):
    # B48 = MAX(Registro!I30:I64)*1000: calculada a partir da coluna I que
    # acabamos de preencher, então mantemos a fórmula
    return definir_numero(xml, "B48", dispersao, manter_formula=True)


with ZipFile(ORIGINAL) as origem:
    caminho_registro = caminho_da_aba(origem, "Registro")
    caminho_certificado = caminho_da_aba(origem, "Certificado")
    caminho_incerteza = caminho_da_aba(origem, "Incerteza")

    # lê tudo pra memória de uma vez: reusar os ZipInfo originais em dois
    # writestr diferentes (um por cenário) corrompe o CRC guardado neles
    conteudo_original = {item.filename: origem.read(item.filename) for item in origem.infolist()}

registro_original = conteudo_original[caminho_registro].decode("utf-8")
certificado_original = conteudo_original[caminho_certificado].decode("utf-8")
incerteza_original = conteudo_original[caminho_incerteza].decode("utf-8")

for nome_arquivo, cenario in CENARIOS.items():
    registro_novo = montar_registro(registro_original, cenario["linhas"])
    certificado_novo = montar_certificado(certificado_original)
    incerteza_novo = montar_incerteza(incerteza_original, cenario["dispersao"])

    conteudo_por_caminho = {
        caminho_registro: registro_novo.encode("utf-8"),
        caminho_certificado: certificado_novo.encode("utf-8"),
        caminho_incerteza: incerteza_novo.encode("utf-8"),
    }

    destino = PASTA_TESTES / nome_arquivo
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as saida:
        for nome_membro, dados_originais in conteudo_original.items():
            saida.writestr(nome_membro, conteudo_por_caminho.get(nome_membro, dados_originais))

    print("planilha de teste gerada em:", destino)
