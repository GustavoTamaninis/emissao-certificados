from decimal import Decimal

from formatacao import arredondar


def calcular_resultados(parametros):
    """Calcula média, erro e aprovação de cada parâmetro a partir das duas medidas.

    A tolerância é assimétrica (tol_menos e tol_mais não são o mesmo valor com
    sinais opostos), por isso o erro é comparado contra os dois limites
    separadamente, em vez de comparar o valor absoluto do erro contra um único
    limite.

    Devolve três coisas:
      - a lista de resultados (cada parâmetro com "media" e "erro" calculados,
        mais "dentro_da_tolerancia" para uso interno);
      - o parecer geral: "APROVADO", "REPROVADO" ou "NÃO AVALIADO";
      - o motivo, só preenchido quando o parecer é "NÃO AVALIADO" (senão None).

    O parecer é "NÃO AVALIADO" quando não há como calcular: planilha sem
    nenhuma medida, ou alguma característica sem as duas medições. Caso
    contrário, "APROVADO" se todos os parâmetros estiverem dentro da
    tolerância, "REPROVADO" se pelo menos um estiver fora.
    """
    if not parametros:
        return [], "NÃO AVALIADO", "A planilha não contém nenhuma medida preenchida."

    resultados = []
    aprovado_geral = True
    falta_medida = False

    for parametro in parametros:
        if parametro["medida_1"] is None or parametro["medida_2"] is None:
            falta_medida = True
            resultados.append({
                "parametro_avaliado": parametro["parametro_avaliado"],
                "valor_nominal": parametro["valor_nominal"],
                "tol_mais": parametro["tol_mais"],
                "tol_menos": parametro["tol_menos"],
                "media": None,
                "erro": None,
                "dentro_da_tolerancia": None,
            })
            continue

        medida_1 = Decimal(str(parametro["medida_1"]))
        medida_2 = Decimal(str(parametro["medida_2"]))
        valor_nominal = Decimal(str(parametro["valor_nominal"]))
        tol_mais = Decimal(str(parametro["tol_mais"]))
        tol_menos = Decimal(str(parametro["tol_menos"]))

        media = arredondar((medida_1 + medida_2) / 2, 4)
         
         
        #=====================================================================
        # 2. CÁLCULO DO ERRO SISTEMÁTICO (Diferença entre média lida e valor padrão
        # =====================================================================

        erro_sistematico = arredondar(media - valor_nominal, 4) 

        # =====================================================================
        # 3. VALIDAÇÃO POR PONTO E LEA (Comparação ponto a ponto com limites)
        # =====================================================================
        dentro_da_tolerancia = tol_menos <= erro_sistematico <= tol_mais

        if not dentro_da_tolerancia:
            aprovado_geral = False

        resultados.append({
            "parametro_avaliado": parametro["parametro_avaliado"],
            "valor_nominal": parametro["valor_nominal"],
            "tol_mais": parametro["tol_mais"],
            "tol_menos": parametro["tol_menos"],
            "media": media,
            "erro": erro_sistematico,
            "dentro_da_tolerancia": dentro_da_tolerancia,
        })

    if falta_medida:
        return resultados, "NÃO AVALIADO", "Há característica sem as duas medições."

    parecer = "APROVADO" if aprovado_geral else "REPROVADO"
    return resultados, parecer, None
