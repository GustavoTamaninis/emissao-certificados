from defusedxml.ElementTree import fromstring


def extrair_dados(xml_certificado, xml_incerteza, xml_textos=None):
    """Extrai os campos do modelo usando os XMLs internos do Excel.

    Os argumentos recebem o conteúdo XML (str ou bytes), não o caminho.
    xml_textos é o conteúdo de sharedStrings.xml, quando existir.
    Fórmulas não são executadas: são lidos apenas os valores salvos.
    """
    certificado = fromstring(xml_certificado, forbid_dtd=True)
    incerteza = fromstring(xml_incerteza, forbid_dtd=True)
    for aba in (certificado, incerteza):
        if aba.tag.split("}")[-1] != "worksheet":
            raise ValueError("Informe os XMLs das abas do Excel.")

    textos = []
    if xml_textos is not None:
        compartilhados = fromstring(xml_textos, forbid_dtd=True)
        for item in compartilhados.findall("{*}si"):
            textos.append("".join(t.text or "" for t in item.findall(".//{*}t")))

    celulas_certificado = {c.attrib["r"]: c for c in certificado.findall(".//{*}c")}
    celulas_incerteza = {c.attrib["r"]: c for c in incerteza.findall(".//{*}c")}

    def valor(celulas, endereco):
        celula = celulas.get(endereco)
        if celula is None:
            return None

        conteudo = celula.findtext("{*}v")
        if celula.get("t") == "s" and conteudo is not None:
            indice = int(conteudo)
            if not 0 <= indice < len(textos):
                raise ValueError("Passe também o XML de sharedStrings.xml.")
            conteudo = textos[indice]
        elif celula.get("t") == "inlineStr":
            conteudo = "".join(t.text or "" for t in celula.findall(".//{*}t"))

        if conteudo is None:
            return None
        return conteudo.strip() or None

    dados = {
        "data": valor(celulas_certificado, "H6"),
        "cliente_solicitante": valor(celulas_certificado, "B8"),
        "tecnico_executante": valor(celulas_certificado, "C52"),
        "padrao_utilizado": valor(celulas_certificado, "A20"),
        "codigo": valor(celulas_certificado, "B13"),
        "denominacao": valor(celulas_certificado, "B12"),
        "fabricante": valor(celulas_certificado, "B14"),
        "incerteza_medicao": valor(celulas_certificado, "D45"),
    }

    dados["parametros"] = []
    # No modelo, os resultados ficam entre o cabeçalho B34 e a incerteza D45.
    for linha in range(35, 45):
        parametro = valor(celulas_certificado, f"B{linha}")
        if parametro is not None:
            dados["parametros"].append({
                "parametro_avaliado": parametro,
                "valor_nominal": valor(celulas_certificado, f"E{linha}"),
                "tol_mais": valor(celulas_certificado, f"F{linha}"),
                "tol_menos": valor(celulas_certificado, f"G{linha}"),
            })

    # A dispersão aparece em sete blocos; os endereços identificam cada ocorrência.
    dados["dispersao_medidas"] = {
        endereco: valor(celulas_incerteza, endereco)
        for endereco in ("B11", "B48", "B82", "B117", "B146", "B177", "B204")
    }

    return dados
