"""Limpeza de entradas sem arredondar medições nem executar fórmulas."""

import re
import unicodedata
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, localcontext
from html import unescape


class ErroImportacao(ValueError):
    pass


def texto(valor, *, identificador=False):
    if valor is None:
        return None
    valor = unicodedata.normalize("NFC", unescape(str(valor)))
    valor = re.sub(r"<[^>]*>", "", valor)
    valor = "".join(c for c in valor if not unicodedata.category(c).startswith("C") or c.isspace())
    valor = " ".join(valor.split())
    if not valor or re.fullmatch(r"[-?_*\s]+", valor):
        return None
    if valor.upper() in {"N/A", "NULL", "NONE", "NÃO CONSTA", "NAO CONSTA", "00/00/0000"}:
        return None
    if valor.startswith(("=", "+=", "@=")) or valor.startswith("#"):
        raise ErroImportacao("Foi encontrado um campo com fórmula ou erro de planilha.")
    if len(valor) > 500:
        raise ErroImportacao("Um campo de texto ultrapassa o limite de 500 caracteres.")
    return valor.upper() if identificador else valor


def chave(valor):
    valor = unicodedata.normalize("NFKD", str(valor or ""))
    return re.sub(r"[^a-z0-9+-]", "", "".join(c for c in valor if not unicodedata.combining(c)).lower())


def numero(valor, *, unidade=None, obrigatorio=False):
    valor = texto(valor)
    if valor is None:
        if obrigatorio:
            raise ErroImportacao("Uma medição, valor nominal ou tolerância obrigatória está vazia.")
        return None
    valor = valor.replace("−", "-")
    match = re.fullmatch(r"([+-]?(?:\d[\d.,]*|[.,]\d+)(?:[eE][+-]?\d+)?)\s*([a-zA-Zµμ°º]*)", valor)
    if not match:
        raise ErroImportacao(f"Número inválido: {valor!r}.")
    bruto, sufixo = match.groups()
    if len(bruto) > 45:
        raise ErroImportacao("Número com precisão ou tamanho excessivo.")
    if "," in bruto and "." in bruto:
        decimal = "," if bruto.rfind(",") > bruto.rfind(".") else "."
        agrupador = "." if decimal == "," else ","
        if not re.fullmatch(rf"[+-]?\d{{1,3}}(?:\{agrupador}\d{{3}})+\{decimal}\d+", bruto):
            raise ErroImportacao(f"Separadores numéricos inconsistentes: {valor!r}.")
        bruto = bruto.replace(agrupador, "").replace(decimal, ".")
    else:
        bruto = bruto.replace(",", ".")
    try:
        resultado = Decimal(bruto)
    except InvalidOperation as erro:
        raise ErroImportacao(f"Número inválido: {valor!r}.") from erro
    if not resultado.is_finite() or abs(resultado) > Decimal("1e12") or resultado.as_tuple().exponent < -30:
        raise ErroImportacao("Número fora do limite aceito pelo importador.")
    sufixo = sufixo.lower().replace("μ", "µ").replace("º", "°")
    if unidade == "mm":
        fatores = {"": Decimal(1), "mm": Decimal(1), "µm": Decimal("0.001"), "um": Decimal("0.001"), "cm": Decimal(10)}
    elif unidade == "celsius":
        fatores = {"": Decimal(1), "c": Decimal(1), "°c": Decimal(1), "°": Decimal(1)}
    else:
        fatores = {"": Decimal(1)}
    if sufixo not in fatores:
        raise ErroImportacao(f"Unidade incompatível com {unidade or 'este campo'}: {sufixo!r}.")
    # shift decimal rather than binary floats; no rounding on import
    with localcontext() as contexto:
        contexto.prec = 60
        resultado *= fatores[sufixo]
    if not resultado:
        resultado = Decimal(0)
    return format(resultado, "f")


def data_iso(valor, *, epoch1904=False):
    valor = texto(valor)
    if valor is None:
        return None
    if re.fullmatch(r"\d+(?:\.\d+)?", valor):
        serial = Decimal(valor)
        if serial <= 0 or serial > 100000 or (not epoch1904 and int(serial) == 60):
            raise ErroImportacao("Data serial do Excel inválida.")
        base = datetime(1904, 1, 1) if epoch1904 else datetime(1899, 12, 30)
        if not epoch1904 and serial < 60:
            base = datetime(1899, 12, 31)
        return (base + timedelta(days=int(serial))).date().isoformat()
    try:
        return datetime.fromisoformat(valor).date().isoformat()
    except ValueError:
        for formato in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S"):
            try:
                return datetime.strptime(valor, formato).date().isoformat()
            except ValueError:
                continue
    raise ErroImportacao(f"Data inválida: {valor!r}.")


def hora_iso(valor):
    valor = texto(valor)
    if valor is None:
        return None
    if "T" in valor:
        try:
            return datetime.fromisoformat(valor).strftime("%H:%M:%S")
        except ValueError as erro:
            raise ErroImportacao(f"Horário inválido: {valor!r}.") from erro
    if valor == "0" or re.fullmatch(r"0?\.\d+", valor):
        segundos = int(Decimal(valor) * 86400)
        return f"{segundos // 3600:02}:{segundos % 3600 // 60:02}:{segundos % 60:02}"
    match = re.fullmatch(r"(\d{1,2}):(\d{2}):(\d{2})(?:\.\d+)?", valor)
    if match and all(int(v) < lim for v, lim in zip(match.groups(), (24, 60, 60))):
        return ":".join(f"{int(v):02}" for v in match.groups())
    raise ErroImportacao(f"Horário inválido: {valor!r}.")
