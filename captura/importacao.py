"""Contrato de entrega para a equipe: entradas saneadas, cálculos ainda nulos."""

from datetime import datetime, timezone
from decimal import Decimal

from .arquivos import Planilha, filho, filhos, ler_arquivo, nome_local
from .sanitizacao import ErroImportacao, chave, data_iso, hora_iso, numero, texto


# Endereços conferidos no modelo Dispositivo de controle.xlsx, inclusive abas ocultas.
CAMPOS_COMPLEMENTO = {
    "numero_certificado": "B75", "codigo_dispositivo": "B40", "codigo_interno": "B39",
    "denominacao_disp": "B41", "fabricante_disp": "B43", "modelo_disp": "B42",
    "numero_serie": "B44", "valor_nominal_disp": "B54", "resolucao_disp": "B55",
    "unidade_disp": "B56", "data_calibracao": "B72", "data_emissao": "B72",
    "proxima_calibracao": "B73", "ordem_servico": "B84", "tecnico_executante": "B76",
    "matricula_tecnico": "B77", "cliente_solicitante": "B24", "cliente_cidade": "B28",
    "cliente_estado": "B31", "laboratorio_nome": "B9", "laboratorio_endereco": "B11",
    "laboratorio_cidade": "B12", "laboratorio_estado": "B14", "laboratorio_telefone": "B15",
    "laboratorio_email": "B17", "procedimento_denominacao": "B2", "procedimento_numero": "B3",
    "procedimento_codigo_alternativo": "B6", "procedimento_revisao": "B4",
}
CAMPOS_REGISTRO = {
    "codigo_dispositivo": "C18", "denominacao_disp": "C20", "fabricante_disp": "C24",
    "data_calibracao": "C5", "cliente_solicitante": "C7", "tecnico_executante": "C9",
    "temperatura_amb": "I5", "ordem_servico": "I7", "padrao_seletor": "C12",
    "observacoes": "A71",
}
CAMPOS_ENSAIO = ("hora_ensaio", "programa_medicao", "numero_desenho", "numero_peca", "codigo_maquina", "fornecedor")
CAMPOS_RESPONSAVEIS = ("signatario_nome", "signatario_cargo")
CAMPOS_CALCULADOS = ("maior_desvio_mm", "dispersao_medidas_um", "fator_k", "incerteza_expand_mm", "status_geral", "caminho_pdf")
IDENTIFICADORES = {"codigo_dispositivo", "codigo_interno", "numero_serie", "codigo_maquina", "numero_peca"}
DATAS = {"data_calibracao", "data_emissao", "proxima_calibracao"}
OBRIGATORIOS = {
    "codigo_dispositivo": "Código/TAG do dispositivo", "denominacao_disp": "Denominação do dispositivo",
    "data_calibracao": "Data da calibração", "tecnico_executante": "Técnico executante",
    "cliente_solicitante": "Solicitante", "temperatura_amb": "Temperatura ambiente",
    "procedimento_denominacao": "Denominação do procedimento", "procedimento_numero": "Número do procedimento",
    "procedimento_revisao": "Revisão do procedimento", "signatario_nome": "Nome do signatário",
    "signatario_cargo": "Cargo do signatário", "padrao_seletor": "Seleção do padrão utilizado",
}
ALIASES = {
    "parametro_avaliado": ("Characteristic", "Dimension", "Característica", "Parâmetro avaliado"),
    "valor_nominal": ("Nominal", "Valor nominal"),
    "tolerancia_positiva": ("Upper Tol", "+TOL", "Tol. +", "Tolerância superior", "Tolerância positiva"),
    "tolerancia_negativa": ("Lower Tol", "-TOL", "Tol. -", "Tolerância inferior", "Tolerância negativa"),
    "medida_1": ("MEAS 1", "Medida 1"), "medida_2": ("MEAS 2", "Medida 2"),
    "valor_medido": ("Actual", "Measured", "Valor medido"),
}
ALIAS_COLUNAS = {chave(alias): campo for campo, aliases in ALIASES.items() for alias in aliases}


def campos_xml(no, permitidos):
    """Somente campos conhecidos; tags repetidas e estruturas aninhadas são rejeitadas."""
    if no is None:
        return {}
    resultado = {}
    for item in no:
        campo = nome_local(item.tag)
        if campo not in permitidos:
            raise ErroImportacao(f"Campo XML não reconhecido: {campo}.")
        if campo in resultado or len(item):
            raise ErroImportacao(f"Campo XML duplicado ou aninhado: {campo}.")
        resultado[campo] = item.text
    return resultado


def validar_blocos(raiz, esperada, blocos):
    if nome_local(raiz.tag) != esperada:
        raise ErroImportacao(f"Este arquivo deve ter raiz XML <{esperada}>. Consulte os exemplos.")
    vistos = set()
    for no in raiz:
        nome = nome_local(no.tag)
        if nome not in blocos or nome in vistos:
            raise ErroImportacao(f"Bloco XML não reconhecido ou duplicado: {nome}.")
        vistos.add(nome)


def sanear_cabecalho(bruto, epoch1904=False):
    resultado = {}
    for campo, valor in bruto.items():
        if campo in DATAS:
            resultado[campo] = data_iso(valor, epoch1904=epoch1904)
        elif campo == "hora_ensaio":
            resultado[campo] = hora_iso(valor)
        elif campo == "temperatura_amb":
            resultado[campo] = numero(valor, unidade="celsius")
        elif campo == "resolucao_disp":
            resultado[campo] = numero(valor)
        else:
            resultado[campo] = texto(valor, identificador=campo in IDENTIFICADORES)
    resultado["fabricante_disp"] = resultado.get("fabricante_disp") or "NÃO CONSTA"
    return resultado


def extrair_padrao(bruto, epoch1904=False):
    numericos = {"valor_obtido", "incerteza", "fator_k", "resolucao", "erro"}
    resultado = {}
    for campo, valor in bruto.items():
        if campo in numericos:
            resultado[campo] = numero(valor)
        elif campo in {"ultima_calibracao", "proxima_calibracao"}:
            resultado[campo] = data_iso(valor, epoch1904=epoch1904)
        else:
            resultado[campo] = texto(valor, identificador=campo in {"codigo", "numero_serie"})
    return resultado


def extrair_insumos(planilha):
    """Guarda constantes da aba Incerteza; valores derivados nunca são copiados."""
    aba = planilha.aba("Incerteza")
    blocos = []
    bloco = None
    for linha, valores in aba.linhas().items():
        label = chave(valores.get(1))
        if label == "objeto":
            bloco = {"objeto": texto(valores.get(2)), "entradas": [], "componentes": []}
            blocos.append(bloco)
        elif bloco and label == "componentedaincerteza":
            bloco["cabecalho_unidade"] = texto(valores.get(2))
        elif bloco and label in {"comprimentoavaliadomm", "comprimentoavaliadodecimal", "numerodemedicoesrealizadas", "diferencadetemperaturaentrepadraoeobjeto"}:
            for coluna_label, coluna_valor in ((1, 3), (4, 8)):
                nome = texto(valores.get(coluna_label))
                if nome:
                    bloco["entradas"].append({"nome": nome, "valor": numero(valores.get(coluna_valor)),
                                              "depende_de_formula": (linha, coluna_valor) in aba.formulas,
                                              "origem": f"Incerteza:linha {linha}, coluna {coluna_valor}"})
        elif bloco and label and (valores.get(5) or (linha, 2) in aba.formulas) and label != "componentedaincerteza":
            campos = {"valor": 2, "divisor": 3, "coeficiente_sensibilidade": 6, "graus_liberdade": 8}
            bloco["componentes"].append({"nome": texto(valores.get(1)), "distribuicao": texto(valores.get(5)),
                                        **{c: numero(valores.get(col)) for c, col in campos.items()},
                                        "campos_por_formula": [c for c, col in campos.items() if (linha, col) in aba.formulas],
                                        "origem": f"Incerteza:linha {linha}"})
    return blocos


def extrair_dispositivo(arquivo):
    origens = {}
    if isinstance(arquivo, Planilha):
        complemento = arquivo.aba("Complemento")
        if not complemento.celulas:
            raise ErroImportacao("A planilha do dispositivo precisa conter a aba Complemento do modelo analisado.")
        bruto = {campo: complemento.valor(ref) for campo, ref in CAMPOS_COMPLEMENTO.items()}
        origens.update({campo: f"dispositivo:Complemento!{ref}" for campo, ref in CAMPOS_COMPLEMENTO.items()})
        registro = arquivo.aba("Registro")
        for campo, ref in CAMPOS_REGISTRO.items():
            if bruto.get(campo) is None or texto(bruto[campo]) is None:
                bruto[campo] = registro.valor(ref)
                origens[campo] = f"dispositivo:Registro!{ref}"
        protocolo = arquivo.aba("Protocolo")
        seletor = numero(protocolo.valor("F1"))
        if seletor is not None:
            if Decimal(seletor) not in {Decimal(i) for i in range(1, 6)}:
                raise ErroImportacao("Seletor de signatário inválido na aba Protocolo.")
            linha = int(Decimal(seletor)) * 3
            for campo, ref in (("signatario_nome", f"D{linha}"), ("signatario_cargo", f"D{linha + 1}")):
                bruto[campo] = protocolo.valor(ref)
                origens[campo] = f"dispositivo:Protocolo!{ref}"
        padroes = []
        aba = arquivo.aba("Padrões")
        linhas = aba.linhas()
        for linha, valores in linhas.items():
            if numero_sequencia(valores.get(1)) is None or texto(valores.get(3)) is None or texto(valores.get(2)) is None:
                continue
            identidade = {"sequencia": valores.get(1), "codigo": valores.get(2), "descricao": valores.get(3),
                          "modelo": valores.get(6), "fabricante": valores.get(7), "numero_serie": valores.get(8),
                          "valor_nominal": valores.get(9)}
            # A faixa metrológica segue a linha de cabeçalho U.M. logo abaixo.
            dados = linhas.get(linha + 2, {})
            identidade.update({"unidade": dados.get(1), "valor_obtido": dados.get(2), "incerteza": dados.get(3),
                               "fator_k": dados.get(4), "ultima_calibracao": dados.get(6), "numero_certificado": dados.get(7),
                               "executante": dados.get(8), "proxima_calibracao": dados.get(9),
                               "resolucao": dados.get(10), "erro": dados.get(11)})
            padrao = extrair_padrao(identidade, arquivo.epoch1904)
            faixas = []
            for l in range(linha + 2, max(linhas, default=0) + 1):
                faixa = linhas.get(l, {})
                if texto(faixa.get(1)) not in {"µm", "μm", "um", "mm"}:
                    break
                faixas.append({"unidade": texto(faixa.get(1)), "descricao": texto(faixa.get(5)),
                               **{campo: numero(faixa.get(col)) for campo, col in
                                  (("valor_obtido", 2), ("incerteza", 3), ("fator_k", 4), ("erro", 11))},
                               "origem": f"dispositivo:Padrões!linha {l}"})
            padrao["faixas_metrologicas"] = faixas
            padrao["origem"] = f"dispositivo:Padrões!linha {linha}"
            padroes.append(padrao)
        insumos = extrair_insumos(arquivo)
        formulas = sum(len(aba.formulas) for aba in arquivo.abas)
        epoch = arquivo.epoch1904
    else:
        validar_blocos(arquivo, "dispositivo", {"calibracao", "padroes", "insumos_incerteza"})
        permitido = set(CAMPOS_COMPLEMENTO) | set(CAMPOS_REGISTRO) | set(CAMPOS_RESPONSAVEIS) | set(CAMPOS_ENSAIO)
        bruto = campos_xml(filho(arquivo, "calibracao"), permitido)
        origens.update({campo: f"dispositivo:XML/calibracao/{campo}" for campo in bruto})
        padroes = []
        container = filho(arquivo, "padroes")
        campos_padrao = {"sequencia", "codigo", "descricao", "modelo", "fabricante", "numero_serie", "valor_nominal",
                         "unidade", "valor_obtido", "incerteza", "fator_k", "ultima_calibracao", "numero_certificado",
                         "executante", "proxima_calibracao", "resolucao", "erro"}
        for no in container if container is not None else []:
            if nome_local(no.tag) != "padrao":
                raise ErroImportacao("Use <padrao> dentro de <padroes>.")
            padroes.append({**extrair_padrao(campos_xml(no, campos_padrao)), "origem": "dispositivo:XML/padroes/padrao"})
        insumos = []
        container = filho(arquivo, "insumos_incerteza")
        for no in container if container is not None else []:
            if nome_local(no.tag) != "componente":
                raise ErroImportacao("Use <componente> dentro de <insumos_incerteza>.")
            campos = campos_xml(no, {"nome", "valor", "unidade", "divisor", "distribuicao", "coeficiente_sensibilidade", "graus_liberdade"})
            insumos.append({c: numero(v) if c in {"valor", "divisor", "coeficiente_sensibilidade", "graus_liberdade"} else texto(v)
                            for c, v in campos.items()})
        formulas, epoch = 0, False
    campos = set(CAMPOS_COMPLEMENTO) | set(CAMPOS_REGISTRO) | set(CAMPOS_RESPONSAVEIS) | set(CAMPOS_ENSAIO)
    cabecalho = sanear_cabecalho({campo: bruto.get(campo) for campo in sorted(campos)}, epoch)
    return cabecalho, padroes, insumos, origens, formulas


def numero_sequencia(valor):
    valor = str(valor or "")
    return int(valor) if valor.isdigit() and 0 < int(valor) < 1000 else None


def extrair_medicoes(arquivo):
    brutas = []
    cabecalho = {}
    if isinstance(arquivo, Planilha):
        for aba in arquivo.abas:
            colunas = None
            for linha, valores in aba.linhas().items():
                detectadas = {ALIAS_COLUNAS[chave(v)]: c for c, v in valores.items() if chave(v) in ALIAS_COLUNAS}
                if {"parametro_avaliado", "valor_nominal", "tolerancia_positiva", "tolerancia_negativa"} <= detectadas.keys() and (
                    "valor_medido" in detectadas or {"medida_1", "medida_2"} <= detectadas.keys()
                ):
                    colunas = detectadas
                    continue
                if colunas is None:
                    # Metadados anteriores ao cabeçalho não entram na tabela de medições.
                    metadados = {"date": "data_calibracao", "time": "hora_ensaio", "measurementplan": "programa_medicao",
                                 "partprogramname": "programa_medicao", "partno": "numero_peca", "drawingno": "numero_desenho",
                                 "metrologista": "tecnico_executante", "mmc": "codigo_maquina", "fornecedor": "fornecedor"}
                    for coluna, valor in valores.items():
                        campo = metadados.get(chave(valor))
                        if campo:
                            seguinte = next((v for c, v in sorted(valores.items()) if c > coluna and v not in (None, "")), None)
                            cabecalho[campo] = seguinte
                    continue
                dados = {campo: valores.get(coluna) for campo, coluna in colunas.items()}
                if all(texto(v) is None for v in dados.values()):
                    continue
                if texto(dados.get("parametro_avaliado")) is None:
                    raise ErroImportacao(f"Medição sem nome na aba {aba.nome}, linha {linha}.")
                brutas.append({**dados, "origem": f"medicoes:{aba.nome}!linha {linha}"})
        epoch = arquivo.epoch1904
    else:
        validar_blocos(arquivo, "medicoes", {"cabecalho", "resultados"})
        cabecalho = campos_xml(filho(arquivo, "cabecalho"), set(CAMPOS_ENSAIO) | {"data_calibracao", "codigo_dispositivo", "tecnico_executante"})
        container = filho(arquivo, "resultados")
        for i, no in enumerate(container if container is not None else [], 1):
            if nome_local(no.tag) != "medicao":
                raise ErroImportacao("Use <medicao> dentro de <resultados>.")
            brutas.append({**campos_xml(no, set(ALIASES)), "origem": f"medicoes:XML/resultados/medicao[{i}]"})
        epoch = False
    if not brutas:
        raise ErroImportacao("Nenhuma medição encontrada. Confira o relatório e seus cabeçalhos.")
    if len(brutas) > 2000:
        raise ErroImportacao("Limite de 1.000 características (duas leituras cada) ultrapassado.")
    grupos = {}
    for item in brutas:
        nome = texto(item.get("parametro_avaliado"), identificador=True)
        if nome is None:
            raise ErroImportacao("Uma característica está sem nome.")
        base = {campo: numero(item.get(campo), unidade="mm", obrigatorio=True)
                for campo in ("valor_nominal", "tolerancia_positiva", "tolerancia_negativa")}
        if Decimal(base["tolerancia_negativa"]) > Decimal(base["tolerancia_positiva"]):
            raise ErroImportacao(f"Tolerâncias invertidas em {nome}.")
        grupo = grupos.setdefault(nome, {"parametro_avaliado": nome, **base, "leituras": [], "origens": []})
        if any(Decimal(grupo[c]) != Decimal(v) for c, v in base.items()):
            raise ErroImportacao(f"Nominal ou tolerâncias divergentes entre as leituras de {nome}.")
        if "medida_1" in item or "medida_2" in item:
            if "valor_medido" in item:
                raise ErroImportacao(f"Não misture valor_medido com medida_1/medida_2 em {nome}.")
            leituras = [numero(item.get(c), unidade="mm", obrigatorio=True) for c in ("medida_1", "medida_2")]
        else:
            leituras = [numero(item.get("valor_medido"), unidade="mm", obrigatorio=True)]
        grupo["leituras"].extend(leituras)
        grupo["origens"].append(item["origem"])
    parametros = []
    for ordem, grupo in enumerate(grupos.values(), 1):
        leituras = grupo.pop("leituras")
        if len(leituras) != 2:
            raise ErroImportacao(f"{grupo['parametro_avaliado']} tem {len(leituras)} leitura(s). São exigidas exatamente duas.")
        parametros.append({"ordem_linha": ordem, **grupo, "medida_1": leituras[0], "medida_2": leituras[1], "unidade": "mm",
                           "valor_obtido_media": None, "desvio_padrao": None, "erro_indicacao": None, "status_conformidade": None})
    if len(parametros) > 1000:
        raise ErroImportacao("Limite de 1.000 características ultrapassado.")
    return parametros, sanear_cabecalho(cabecalho, epoch)


def importar(nome_dispositivo, conteudo_dispositivo, nome_medicoes, conteudo_medicoes):
    cabecalho, padroes, insumos, origens, formulas = extrair_dispositivo(ler_arquivo(nome_dispositivo, conteudo_dispositivo))
    parametros, ensaio = extrair_medicoes(ler_arquivo(nome_medicoes, conteudo_medicoes))
    # Both uploads must identify the same calibration. Never silently overwrite conflicts.
    for campo, valor in ensaio.items():
        if campo == "fabricante_disp" or valor is None:
            continue
        if cabecalho.get(campo) is not None and cabecalho[campo] != valor:
            raise ErroImportacao(f"Os arquivos divergem no campo {campo}. Confira se pertencem ao mesmo ensaio.")
        if cabecalho.get(campo) is None:
            cabecalho[campo] = valor
            origens[campo] = f"medicoes:cabecalho/{campo}"
    cabecalho.update({campo: None for campo in CAMPOS_CALCULADOS})
    cabecalho["data_importacao"] = datetime.now(timezone.utc).isoformat()
    pendencias = [{"campo": campo, "mensagem": f"Preencher: {label}."}
                 for campo, label in OBRIGATORIOS.items() if cabecalho.get(campo) is None]
    if not padroes:
        pendencias.append({"campo": "padroes", "mensagem": "Nenhum padrão de referência informado."})
    for i, padrao in enumerate(padroes):
        for campo in ("codigo", "descricao", "fabricante", "numero_certificado", "ultima_calibracao", "proxima_calibracao"):
            if padrao.get(campo) is None:
                pendencias.append({"campo": f"padroes[{i}].{campo}", "mensagem": f"Padrão {i+1}: preencher {campo}."})
    avisos = ["Média, erro, desvio padrão, conformidade e incerteza serão calculados na próxima etapa."]
    if formulas:
        avisos.append(f"{formulas} células com fórmula ou erro foram ignoradas; nenhum resultado em cache foi utilizado.")
    if not cabecalho.get("numero_certificado"):
        avisos.append("Número do certificado ausente; sua atribuição fica para a etapa de emissão.")
    if not insumos:
        avisos.append("Insumos de incerteza ausentes; a equipe deverá fornecê-los antes dos cálculos.")
    return {"versao_contrato": "1.0", "status": "pendente_complemento" if pendencias else "extraido",
            "calibracao": cabecalho, "parametros_medicao": parametros, "padroes": padroes,
            "insumos_incerteza": insumos, "origens_calibracao": origens,
            "pendencias": pendencias, "avisos": avisos,
            "processamento": {"dados_extraidos": True, "calculos_realizados": False, "pdf_gerado": False}}
