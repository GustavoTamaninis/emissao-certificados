import streamlit as st
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

import calculos
import captura
import certificado
from formatacao import formatar_numero

st.set_page_config(page_title="Validar Certificados", layout="wide")

st.title("📄 Gerador de Certificado de Calibração")
st.write("Insira uma planilha de calibração.")

arquivo = st.file_uploader("Selecione a planilha (.xlsx)", type="xlsx", accept_multiple_files=False)

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

if arquivo is not None:
    with ZipFile(arquivo) as planilha:
        textos = "xl/sharedStrings.xml"
        dados = captura.extrair_dados(
            planilha.read(caminho_da_aba(planilha, "Registro")),
            planilha.read(caminho_da_aba(planilha, "Incerteza")),
            planilha.read(caminho_da_aba(planilha, "Certificado")),
            planilha.read(textos) if textos in planilha.namelist() else None,
        )

    resultados, parecer = calculos.calcular_resultados(dados["parametros"])

    numero_certificado = st.text_input("Número do certificado", value="")

    cor_parecer = "green" if parecer == "APROVADO" else "red"
    st.markdown(f"## :{cor_parecer}[Parecer final: {parecer}]")

    linhas_tabela = [
        {
            "Parâmetro avaliado": resultado["parametro_avaliado"],
            "Valor nominal": formatar_numero(resultado["valor_nominal"]),
            "Tol. +": formatar_numero(resultado["tol_mais"]),
            "Tol. -": formatar_numero(resultado["tol_menos"]),
            "Valor obtido": formatar_numero(resultado["media"]),
            "Erro": formatar_numero(resultado["erro"]),
        }
        for resultado in resultados
    ]
    st.table(linhas_tabela)

    st.write(f"**Solicitante:** {dados['cliente_solicitante'] or ''}")
    st.write(f"**Denominação:** {dados['denominacao'] or ''}")
    st.write(f"**Código:** {dados['codigo'] or ''}")
    st.write(f"**Fabricante:** {dados['fabricante'] or ''}")
    st.write(f"**Técnico executante:** {dados['tecnico_executante'] or ''}")
    st.write(f"**Incerteza de medição:** {formatar_numero(dados['incerteza_medicao'])} mm")
    st.write(f"**k:** {formatar_numero(dados['fator_k'])}")

    pdf_bytes = (
        certificado.gerar_certificado_pdf(dados, resultados, parecer, numero_certificado)
        if numero_certificado
        else b""
    )
    nome_arquivo = f"certificado_{numero_certificado.replace('/', '-')}.pdf"

    st.download_button(
        "Gerar certificado em PDF",
        data=pdf_bytes,
        file_name=nome_arquivo,
        mime="application/pdf",
        disabled=not numero_certificado,
    )


