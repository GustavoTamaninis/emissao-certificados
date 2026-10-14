from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime


def arredondar(valor, casas):
    """Arredonda para o número de casas decimais indicado, usando ROUND_HALF_UP.

    Não usamos o round() nativo porque ele arredonda para o par mais próximo
    (arredondamento bancário) e erra casos como round(0.0025, 3), que deveria
    dar 0.003 e dá 0.002.
    """
    decimal_valor = valor if isinstance(valor, Decimal) else Decimal(str(valor))
    quantizador = Decimal(1) / (Decimal(10) ** casas)
    return decimal_valor.quantize(quantizador, rounding=ROUND_HALF_UP)


def formatar_numero(valor, casas=3):
    """Formata um número para exibição: casas decimais fixas e vírgula, padrão brasileiro."""
    if valor is None:
        return ""
    arredondado = arredondar(valor, casas)
    return f"{arredondado:.{casas}f}".replace(".", ",")


def formatar_data(valor):
    """Converte a data da planilha para dd/mm/aaaa.

    A planilha grava a data em ISO com uma fração de segundo muito longa (ex.:
    "2026-08-26T16:47:10.00000000465661175") ou com o texto literal "00/00/0000"
    quando o modelo está em branco. Em ambos os casos só precisamos da parte antes
    do "T". Quando não dá para converter, devolve o texto original sem quebrar.
    """
    if not valor:
        return valor or ""
    try:
        parte_data = valor.split("T")[0]
        return datetime.strptime(parte_data, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return valor
