import io
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from model.formatacao import formatar_data, formatar_numero

LARGURA_UTIL = 18 * cm


def _texto(valor):
    """Escapa o texto para não quebrar o mini-HTML do Paragraph; vazio quando None."""
    return escape(str(valor)) if valor else ""


def gerar_certificado_pdf(dados, resultados, parecer, numero_certificado):
    """Monta o PDF do certificado de calibração e devolve os bytes (BytesIO).

    dados: dicionário vindo de captura.extrair_dados.
    resultados: lista vinda de calculos.calcular_resultados.
    parecer: "APROVADO" ou "REPROVADO", vindo de calculos.calcular_resultados.
    numero_certificado: número digitado pelo operador na tela.
    """
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
    )

    estilos = getSampleStyleSheet()
    estilo_titulo_faixa = ParagraphStyle(
        "TituloFaixa", parent=estilos["Title"], textColor=colors.white, alignment=TA_CENTER,
    )
    estilo_negrito = ParagraphStyle(
        "Negrito", parent=estilos["Normal"], fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4,
    )
    estilo_normal = ParagraphStyle("NormalEspacado", parent=estilos["Normal"], spaceAfter=2)
    estilo_parecer = ParagraphStyle(
        "Parecer",
        parent=estilos["Normal"],
        fontName="Helvetica-Bold",
        fontSize=16,
        alignment=TA_CENTER,
        textColor=colors.green if parecer == "APROVADO" else colors.red,
        spaceBefore=10,
        spaceAfter=10,
    )
    estilo_assinatura = ParagraphStyle("Assinatura", parent=estilos["Normal"], alignment=TA_CENTER)
    estilo_pequena = ParagraphStyle("Pequena", parent=estilos["Normal"], fontSize=8)
    estilo_rodape = ParagraphStyle("Rodape", parent=estilos["Normal"], alignment=TA_CENTER, fontSize=8)

    elementos = []

    # faixa superior, sem logo nem dados da empresa
    faixa = Table([[Paragraph("Certificado de Calibração", estilo_titulo_faixa)]], colWidths=[LARGURA_UTIL])
    faixa.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#003366")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    elementos.append(faixa)
    elementos.append(Spacer(1, 12))

    # número do certificado + data, em negrito
    linha_cabecalho = Table(
        [[
            Paragraph(f"<b>Certificado de Calibração nº {_texto(numero_certificado)}</b>", estilo_normal),
            Paragraph(f"<b>Data: {_texto(formatar_data(dados.get('data')))}</b>", estilo_normal),
        ]],
        colWidths=[LARGURA_UTIL / 2, LARGURA_UTIL / 2],
    )
    linha_cabecalho.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elementos.append(linha_cabecalho)
    elementos.append(Spacer(1, 6))

    elementos.append(Paragraph(f"Solicitante: {_texto(dados.get('cliente_solicitante'))}", estilo_normal))

    elementos.append(Paragraph("1. Descrição do instrumento calibrado", estilo_negrito))
    elementos.append(Paragraph(f"Denominação: {_texto(dados.get('denominacao'))}", estilo_normal))
    elementos.append(Paragraph(f"Código: {_texto(dados.get('codigo'))}", estilo_normal))
    elementos.append(Paragraph(f"Fabricante: {_texto(dados.get('fabricante'))}", estilo_normal))

    elementos.append(Paragraph("2. Padrão utilizado", estilo_negrito))
    elementos.append(Paragraph(_texto(dados.get("padrao_utilizado")), estilo_normal))
    elementos.append(Paragraph(_texto(dados.get("certificado_padrao")), estilo_normal))

    elementos.append(Paragraph("3. Procedimento de calibração", estilo_negrito))
    elementos.append(Paragraph(
        f"Nr. do procedimento / revisão: {_texto(dados.get('procedimento'))}", estilo_normal,
    ))
    elementos.append(Paragraph(_texto(dados.get("metodo")), estilo_normal))
    elementos.append(Paragraph("Os resultados estão referenciados a temperatura de 20° Celsius.", estilo_normal))

    elementos.append(Paragraph(f"Parecer final: {parecer}", estilo_parecer))

    elementos.append(Paragraph("4. Resultados (mm)", estilo_negrito))
    cabecalho_tabela = ["Parâmetro avaliado", "Valor nominal", "Tol. +", "Tol. -", "Valor obtido", "Erro"]
    linhas_tabela = [cabecalho_tabela]
    for resultado in resultados:
        linhas_tabela.append([
            resultado["parametro_avaliado"],
            formatar_numero(resultado["valor_nominal"]),
            formatar_numero(resultado["tol_mais"]),
            formatar_numero(resultado["tol_menos"]),
            formatar_numero(resultado["media"]),
            formatar_numero(resultado["erro"]),
        ])

    largura_primeira_coluna = LARGURA_UTIL * 0.3
    largura_demais_colunas = (LARGURA_UTIL - largura_primeira_coluna) / 5
    tabela_resultados = Table(
        linhas_tabela,
        colWidths=[largura_primeira_coluna] + [largura_demais_colunas] * 5,
    )
    tabela_resultados.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tabela_resultados)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("5. Incerteza de medição:", estilo_negrito))
    elementos.append(Paragraph(
        f"{formatar_numero(dados.get('incerteza_medicao'))} mm      "
        f"k = {formatar_numero(dados.get('fator_k'))}",
        estilo_normal,
    ))

    elementos.append(Paragraph("6. Observações", estilo_negrito))
    elementos.append(Paragraph(_texto(dados.get("observacoes")), estilo_normal))

    elementos.append(Paragraph(
        f"Técnico executante: {_texto(dados.get('tecnico_executante'))}", estilo_normal,
    ))

    elementos.append(Spacer(1, 24))
    elementos.append(Paragraph("_" * 40, estilo_assinatura))
    elementos.append(Paragraph(_texto(dados.get("signatario")), estilo_assinatura))
    elementos.append(Paragraph(_texto(dados.get("cargo_signatario")), estilo_assinatura))

    elementos.append(Spacer(1, 16))
    elementos.append(Paragraph("Notas:", estilo_negrito))
    for nota in dados.get("notas") or []:
        elementos.append(Paragraph(_texto(nota), estilo_pequena))

    elementos.append(Spacer(1, 16))
    elementos.append(Paragraph("Folha 01/01", estilo_rodape))

    documento.build(elementos)
    return buffer.getvalue()
