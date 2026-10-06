"""Leitura de OOXML (incluindo Strict) e XML, apenas em memória."""

import io
import posixpath
import re
from dataclasses import dataclass, field
from zipfile import BadZipFile, ZipFile
from zlib import error as ErroCompressao
from xml.etree.ElementTree import ParseError

from defusedxml import ElementTree
from defusedxml.common import DefusedXmlException

from .sanitizacao import ErroImportacao, chave

MAX_ARQUIVO = 5 * 1024 * 1024
MAX_DESCOMPACTADO = 30 * 1024 * 1024
MAX_CELULAS = 100000


def nome_local(tag):
    return tag.rsplit("}", 1)[-1]


def filhos(no, nome):
    return [c for c in no if nome_local(c.tag) == nome]


def filho(no, nome):
    return next(iter(filhos(no, nome)), None) if no is not None else None


def atributo(no, nome, default=None):
    return next((v for k, v in no.attrib.items() if nome_local(k) == nome), default)


def xml_seguro(conteudo):
    try:
        raiz = ElementTree.fromstring(conteudo, forbid_dtd=True, forbid_entities=True, forbid_external=True)
    except (ParseError, DefusedXmlException, ValueError, LookupError, UnicodeError) as erro:
        raise ErroImportacao("XML inválido ou inseguro: DTDs e entidades externas são proibidos.") from erro
    pilha = [(raiz, 1)]
    quantidade = 0
    while pilha:
        no, nivel = pilha.pop()
        quantidade += 1
        if nivel > 64 or quantidade > 200000:
            raise ErroImportacao("XML ultrapassa os limites de estrutura.")
        pilha.extend((c, nivel + 1) for c in no)
    return raiz


def coordenada(ref):
    match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d{0,6})", ref)
    if not match:
        raise ErroImportacao("Referência de célula inválida no Excel.")
    coluna = 0
    for letra in match[1]:
        coluna = coluna * 26 + ord(letra) - 64
    return int(match[2]), coluna


@dataclass
class Aba:
    nome: str
    celulas: dict = field(default_factory=dict)
    formulas: set = field(default_factory=set)

    def valor(self, ref):
        return self.celulas.get(coordenada(ref))

    def linhas(self):
        resultado = {}
        for (linha, coluna), valor in sorted(self.celulas.items()):
            resultado.setdefault(linha, {})[coluna] = valor
        return resultado


@dataclass
class Planilha:
    abas: list
    epoch1904: bool = False

    def aba(self, nome):
        return next((a for a in self.abas if chave(a.nome) == chave(nome)), Aba(nome))


def ler_xlsx(conteudo):
    try:
        with ZipFile(io.BytesIO(conteudo)) as pacote:
            infos = pacote.infolist()
            nomes = [i.filename for i in infos]
            if len(infos) > 2000 or len(set(nomes)) != len(nomes) or sum(i.file_size for i in infos) > MAX_DESCOMPACTADO:
                raise ErroImportacao("Planilha compactada ultrapassa os limites permitidos.")
            if any("vbaproject" in n.lower() for n in nomes):
                raise ErroImportacao("Planilhas com macros não são aceitas neste MVP.")
            if any(n not in nomes for n in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml")):
                raise ErroImportacao("O arquivo não é uma planilha Excel .xlsx válida.")
            if b"macroenabled" in pacote.read("[Content_Types].xml").lower():
                raise ErroImportacao("Planilhas com macros não são aceitas neste MVP.")
            workbook = xml_seguro(pacote.read("xl/workbook.xml"))
            if nome_local(workbook.tag) != "workbook":
                raise ErroImportacao("Estrutura de pasta de trabalho inválida.")
            rels = {atributo(r, "Id"): r for r in xml_seguro(pacote.read("xl/_rels/workbook.xml.rels"))}
            strings = []
            if "xl/sharedStrings.xml" in nomes:
                strings = ["".join(t.text or "" for t in si.iter() if nome_local(t.tag) == "t")
                           for si in xml_seguro(pacote.read("xl/sharedStrings.xml"))]
            abas = []
            total_celulas = 0
            sheets = filho(workbook, "sheets")
            if sheets is None:
                raise ErroImportacao("A planilha não possui abas.")
            for sheet in sheets:
                rel = rels.get(atributo(sheet, "id"))
                if rel is None or atributo(rel, "TargetMode") == "External":
                    raise ErroImportacao("Referência de aba ausente ou externa.")
                target = atributo(rel, "Target", "")
                caminho = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
                if not caminho.startswith("xl/worksheets/") or caminho not in nomes:
                    raise ErroImportacao("Caminho de aba inválido na planilha.")
                aba = Aba(atributo(sheet, "name", ""))
                raiz = xml_seguro(pacote.read(caminho))
                for c in raiz.iter():
                    if nome_local(c.tag) != "c":
                        continue
                    total_celulas += 1
                    if total_celulas > MAX_CELULAS:
                        raise ErroImportacao("A planilha possui células demais para este MVP.")
                    coord = coordenada(atributo(c, "r", ""))
                    # Cached formula results can be stale. Never accept them as input.
                    if filho(c, "f") is not None or atributo(c, "t") == "e":
                        aba.formulas.add(coord)
                        continue
                    v = filho(c, "v")
                    valor = v.text if v is not None else None
                    tipo = atributo(c, "t")
                    if tipo == "s":
                        try:
                            index = int(valor)
                            if index < 0:
                                raise ValueError()
                            valor = strings[index]
                        except (ValueError, TypeError, IndexError) as erro:
                            raise ErroImportacao("Texto compartilhado inválido na planilha.") from erro
                    elif tipo == "inlineStr":
                        valor = "".join(t.text or "" for t in c.iter() if nome_local(t.tag) == "t")
                    if coord in aba.celulas:
                        raise ErroImportacao("Célula duplicada na planilha.")
                    aba.celulas[coord] = valor
                abas.append(aba)
            pr = filho(workbook, "workbookPr")
            return Planilha(abas, pr is not None and atributo(pr, "date1904") in {"1", "true"})
    except (BadZipFile, KeyError, RuntimeError, NotImplementedError, ErroCompressao, EOFError) as erro:
        raise ErroImportacao("Planilha Excel inválida, corrompida ou criptografada.") from erro


def ler_spreadsheetml(raiz):
    """XML 2003 exportado pelo Excel; índices esparsos são respeitados."""
    abas = []
    quantidade = 0
    for worksheet in filhos(raiz, "Worksheet"):
        aba = Aba(atributo(worksheet, "Name", ""))
        tabela = filho(worksheet, "Table")
        linha = 0
        for row in filhos(tabela, "Row") if tabela is not None else []:
            linha = int(atributo(row, "Index", linha + 1))
            coluna = 0
            for cell in filhos(row, "Cell"):
                coluna = int(atributo(cell, "Index", coluna + 1))
                quantidade += 1
                if quantidade > MAX_CELULAS or linha < 1 or coluna < 1:
                    raise ErroImportacao("Índices ou tamanho inválidos no XML Excel.")
                coord = (linha, coluna)
                if coord in aba.celulas or coord in aba.formulas:
                    raise ErroImportacao("Célula duplicada no XML Excel.")
                data = filho(cell, "Data")
                if atributo(cell, "Formula") or (data is not None and atributo(data, "Type") == "Error"):
                    aba.formulas.add(coord)
                else:
                    aba.celulas[coord] = "".join(data.itertext()) if data is not None else None
                coluna += int(atributo(cell, "MergeAcross", 0))
        abas.append(aba)
    if not abas:
        raise ErroImportacao("XML Excel sem abas.")
    return Planilha(abas)


def ler_arquivo(nome, conteudo):
    if not conteudo:
        raise ErroImportacao("Arquivo vazio.")
    if len(conteudo) > MAX_ARQUIVO:
        raise ErroImportacao("Cada arquivo deve ter no máximo 5 MB.")
    extensao = nome.rsplit(".", 1)[-1].lower()
    if extensao == "xlsx":
        return ler_xlsx(conteudo)
    if extensao == "xml":
        raiz = xml_seguro(conteudo)
        if nome_local(raiz.tag) == "Workbook" and raiz.tag.startswith("{urn:schemas-microsoft-com:office:spreadsheet}"):
            try:
                return ler_spreadsheetml(raiz)
            except (ValueError, TypeError) as erro:
                raise ErroImportacao("Índice inválido no XML Excel.") from erro
        return raiz
    raise ErroImportacao("Formato não aceito. Envie .xlsx ou .xml.")
