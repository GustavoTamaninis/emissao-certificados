from decimal import Decimal

from formatacao import arredondar


def calcular_resultados(parametros):
    """Calcula média, erro e aprovação de cada parâmetro a partir das duas medidas.

    A tolerância é assimétrica (tol_menos e tol_mais não são o mesmo valor com
    sinais opostos), por isso o erro é comparado contra os dois limites
    separadamente, em vez de comparar o valor absoluto do erro contra um único
    limite.

    Devolve a lista de resultados (cada parâmetro com "media" e "erro"
    calculados, mais "dentro_da_tolerancia" para uso interno) e o parecer geral:
    "APROVADO" se todos os parâmetros estiverem dentro da tolerância, senão
    "REPROVADO".
    """
    resultados = []
    aprovado_geral = True

    for parametro in parametros:
        medida_1 = Decimal(str(parametro["medida_1"]))
        medida_2 = Decimal(str(parametro["medida_2"]))
        valor_nominal = Decimal(str(parametro["valor_nominal"]))
        tol_mais = Decimal(str(parametro["tol_mais"]))
        tol_menos = Decimal(str(parametro["tol_menos"]))

        media = arredondar((medida_1 + medida_2) / 2, 4)
        erro = arredondar(media - valor_nominal, 4)
        dentro_da_tolerancia = tol_menos <= erro <= tol_mais

        if not dentro_da_tolerancia:
            aprovado_geral = False

        resultados.append({
            "parametro_avaliado": parametro["parametro_avaliado"],
            "valor_nominal": parametro["valor_nominal"],
            "tol_mais": parametro["tol_mais"],
            "tol_menos": parametro["tol_menos"],
            "media": media,
            "erro": erro,
            "dentro_da_tolerancia": dentro_da_tolerancia,
        })

    parecer = "APROVADO" if aprovado_geral else "REPROVADO"
    return resultados, parecer
