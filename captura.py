import re

from defusedxml.ElementTree import fromstring

def validar_dados_completos(dados):
    """Trata informações incompletas ou campos essenciais vazios."""
    campos_obrigatorios = ['data', 'cliente_solicitante', 'codigo', 'denominacao']
    for campo in campos_obrigatorios:
        if not dados.get(campo):
            raise ValueError(f"Informação incompleta: o campo '{campo}' está vazio na planilha.")
            
    if not dados.get('parametros'):
        raise ValueError("A planilha não contém linhas de parâmetros de medição preenchidas.")

def _ler_textos(xml_textos):
    textos = []
    if xml_textos is not None:
        raiz = fromstring(xml_textos, forbid_dtd=True)
        for item in raiz.findall("{*}si"):
            partes = item.findall("{*}t") + item.findall("{*}r/{*}t")
            textos.append("".join(t.text or "" for t in partes))
    return textos


def _ler_aba(xml, nome, textos, problemas):
    """Lê a aba uma única vez e guarda tudo em um dicionário simples.

    "valores" tem {endereço: texto limpo} só das células preenchidas;
    "erros" tem {endereço: código} das células com erro do Excel (#NAME?, #DIV/0!...).
    Funciona em xlsx comum e Strict, pois o "{*}" ignora o namespace.
    """
    raiz = fromstring(xml, forbid_dtd=True)
    if raiz.tag.split("}")[-1] != "worksheet":
        raise ValueError(f"O XML informado para '{nome}' não é uma aba do Excel.")

    valores, erros = {}, {}
    for celula in raiz.findall(".//{*}c"):
        endereco = celula.attrib["r"]
        tipo = celula.get("t")
        conteudo = celula.findtext("{*}v")

        if tipo == "e":
            erros[endereco] = conteudo
            continue
        if tipo == "s" and conteudo is not None:
            indice = int(conteudo)
            if not 0 <= indice < len(textos):
                raise ValueError("Passe também o XML de sharedStrings.xml.")
            conteudo = textos[indice]
        elif tipo == "inlineStr":
            conteudo = "".join(t.text or "" for t in celula.findall(".//{*}t"))

        if conteudo is not None and conteudo.strip():
            valores[endereco] = conteudo.strip()

    linhas = [int(re.sub(r"\D", "", endereco)) for endereco in (*valores, *erros)]
    return {
        "nome": nome,
        "valores": valores,
        "erros": erros,
        "problemas": problemas,
        "ultima_linha": max(linhas, default=0),
    }


def _texto(aba, endereco):
    """Texto da célula, ou None se vazia. Célula com erro do Excel vira None e é registrada."""
    if endereco in aba["erros"]:
        aba["problemas"].append({
            "aba": aba["nome"],
            "celula": endereco,
            "problema": f"erro do Excel: {aba['erros'][endereco]}",
        })
    return aba["valores"].get(endereco)


def _numero(aba, endereco):
    texto = _texto(aba, endereco)
    if texto is None:
        return None
    try:
        return float(texto)
    except ValueError:
        aba["problemas"].append({
            "aba": aba["nome"],
            "celula": endereco,
            "problema": f"texto onde se esperava número: {texto!r}",
        })
        return None


def _procurar_opcional(aba, coluna, rotulo, desde=1):
    """Como _procurar, mas devolve None em vez de levantar erro quando não acha o rótulo.

    Usada para campos que ficam depois de uma tabela de tamanho variável (aba
    Certificado): o deslocamento é sempre o mesmo a partir do rótulo, mas o
    rótulo em si pode estar em qualquer linha.
    """
    for linha in range(desde, aba["ultima_linha"] + 1):
        if aba["valores"].get(f"{coluna}{linha}") == rotulo:
            return linha
    return None


def _procurar(aba, coluna, rotulo, desde=1):
    """Primeira linha (>= desde) cuja célula da coluna tem exatamente esse texto."""
    linha = _procurar_opcional(aba, coluna, rotulo, desde)
    if linha is None:
        raise ValueError(f"Rótulo '{rotulo}' não encontrado na coluna {coluna} da aba {aba['nome']}.")
    return linha


def _ler_parametros(registro):
    """Tabela "Resultados obtidos (mm)" do Registro, com quantidade variável de linhas.

    Começa logo abaixo do cabeçalho "Parâmetro avaliado" e termina antes da linha
    "Incerteza de medição:". Linhas sem parâmetro (o modelo deixa " ") são ignoradas.
    """
    cabecalho = _procurar(registro, "B", "Parâmetro avaliado")
    fim = _procurar(registro, "C", "Incerteza de medição:", desde=cabecalho + 1)

    parametros = []
    for linha in range(cabecalho + 1, fim):
        parametro = _texto(registro, f"B{linha}")
        if parametro is None:
            continue
        parametros.append({
            "parametro_avaliado": parametro,
            "valor_nominal": _numero(registro, f"C{linha}"),
            "tol_mais": _numero(registro, f"D{linha}"),
            "tol_menos": _numero(registro, f"E{linha}"),
            "medida_1": _numero(registro, f"F{linha}"),
            "medida_2": _numero(registro, f"G{linha}"),
        })
    return parametros


def _ler_incerteza_medicao(registro):
    """Valor ao lado do rótulo "Incerteza de medição:" (rótulo na coluna C, valor na D)."""
    linha = _procurar(registro, "C", "Incerteza de medição:")
    return _numero(registro, f"D{linha}")


def _ler_dispersao(incerteza):
    """"Dispersão das medidas (s)" do bloco Volumétrico da aba Incerteza (valor na coluna B).

    A aba tem dois blocos com esse título; vale o primeiro, que é o que alimenta a
    incerteza final. O segundo é uma cópia com valores digitados e é ignorado.
    """
    bloco = None
    for linha in range(1, incerteza["ultima_linha"] + 1):
        rotulo = incerteza["valores"].get(f"A{linha}")
        if rotulo == "OBJETO:":  # início de um bloco; o título fica ao lado, na coluna B
            bloco = incerteza["valores"].get(f"B{linha}")
        elif rotulo == "Dispersão das medidas (s)" and bloco == "CONTROLE DIMENSIONAL - Volumétrico":
            return _numero(incerteza, f"B{linha}")
    raise ValueError("'Dispersão das medidas (s)' do bloco Volumétrico não encontrada na aba Incerteza.")


def _ler_signatario(certificado):
    """Nome e cargo de quem assina, na aba Certificado.

    Ficam sempre 4 e 5 linhas abaixo do rótulo "Técnico executante:", não em
    endereço fixo, porque a tabela de resultados (que fica acima) tem tamanho
    variável e desloca tudo que vem depois dela.
    """
    linha_tecnico = _procurar_opcional(certificado, "A", "Técnico executante:")
    if linha_tecnico is None:
        return None, None
    signatario = _texto(certificado, f"F{linha_tecnico + 4}")
    cargo_signatario = _texto(certificado, f"F{linha_tecnico + 5}")
    return signatario, cargo_signatario


def _ler_notas(certificado):
    """As notas de rodapé, na aba Certificado: até 4 linhas logo abaixo do rótulo "Notas:".

    Para na primeira linha vazia, porque nem toda planilha preenche as 4.
    """
    linha_notas = _procurar_opcional(certificado, "A", "Notas:")
    if linha_notas is None:
        return []
    notas = []
    for deslocamento in range(1, 5):
        nota = _texto(certificado, f"A{linha_notas + deslocamento}")
        if nota is None:
            break
        notas.append(nota)
    return notas


def extrair_dados(xml_registro, xml_incerteza, xml_certificado, xml_textos=None):
    """Extrai os campos do modelo usando os XMLs internos do Excel.

    Os argumentos recebem o conteúdo XML (str ou bytes), não o caminho:
      - xml_registro: aba Registro (data, cliente, técnico, instrumento, resultados, incerteza)
      - xml_incerteza: aba Incerteza (dispersão das medidas)
      - xml_certificado: aba Certificado (usada só para o texto do padrão utilizado,
        que não existe no Registro)
      - xml_textos: sharedStrings.xml, quando existir

    Fórmulas não são executadas: são lidos apenas os valores salvos.
    Células com erro do Excel viram None e são listadas em "problemas".
    """
    textos = _ler_textos(xml_textos)
    problemas = []
    registro = _ler_aba(xml_registro, "Registro", textos, problemas)
    incerteza = _ler_aba(xml_incerteza, "Incerteza", textos, problemas)
    certificado = _ler_aba(xml_certificado, "Certificado", textos, problemas)

    signatario, cargo_signatario = _ler_signatario(certificado)

    dados = {
        "data": _texto(registro, "C5"),
        "cliente_solicitante": _texto(registro, "C7"),
        "tecnico_executante": _texto(registro, "C9"),
        "padrao_utilizado": _texto(certificado, "A20"),
        "codigo": _texto(registro, "C18"),
        "denominacao": _texto(registro, "C20"),
        "fabricante": _texto(registro, "C24"),
        "numero_certificado_planilha": _texto(certificado, "F6"),
        "certificado_padrao": _texto(certificado, "A21"),
        "procedimento": _texto(certificado, "D26"),
        "metodo": _texto(certificado, "A28"),
        # observacoes e fator_k ficam na aba Registro: são posição fixa, ao
        # contrário da aba Certificado, onde a tabela de resultados tem
        # tamanho variável e desloca tudo que vem depois dela
        "observacoes": _texto(registro, "A71"),
        "fator_k": _numero(registro, "F66"),
        "signatario": signatario,
        "cargo_signatario": cargo_signatario,
        "notas": _ler_notas(certificado),
        "parametros": _ler_parametros(registro),
        "incerteza_medicao": _ler_incerteza_medicao(registro),
        "dispersao_medidas": _ler_dispersao(incerteza),
        "problemas": problemas,
    }
    validar_dados_completos(dados)
    return dados