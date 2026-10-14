"""Roda extração + cálculo (+ PDF) nas planilhas de teste e no arquivo real da
Schulz, e confere se o resultado é o esperado. Não usa framework de teste (pytest
etc.) de propósito: é só uma lista de conferências, fácil de ler de cima a baixo.

As planilhas de teste (geradas por gerar_planilha_teste.py) têm o mesmo layout do
modelo em branco: 9 linhas na tabela da aba Certificado. Nelas o fator k cai no
mesmo endereço (F45) tanto no código com bug quanto no corrigido, então elas não
provam, por si só, que a correção da Parte 1 funciona. Por isso este script roda
também no arquivo real, que tem 14 características e desloca o fator k pra F50 —
é o único jeito de confirmar que a leitura por rótulo (e não por endereço fixo)
está funcionando.

O arquivo real não entra no repositório (.gitignore bloqueia *.xlsx); ajuste o
caminho em ARQUIVO_REAL se ele estiver em outro lugar na sua máquina.
"""
import sys
from pathlib import Path
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from defusedxml.ElementTree import fromstring

import calculos
import captura
import certificado
import formatacao

PASTA_TESTES = Path(__file__).resolve().parent
ARQUIVO_APROVADO = PASTA_TESTES / "planilha_teste_aprovado.xlsx"
ARQUIVO_REPROVADO = PASTA_TESTES / "planilha_teste_reprovado.xlsx"
ARQUIVO_REAL = Path(r"C:\Users\Furla\Downloads\planilha_schulz\Cal_1_5896_1.xlsx")

houve_falha = False


def checar(descricao, condicao):
    global houve_falha
    print(f"[{'OK' if condicao else 'FALHOU'}] {descricao}")
    if not condicao:
        houve_falha = True
    return condicao


def caminho_da_aba(planilha, nome):
    livro = fromstring(planilha.read("xl/workbook.xml"), forbid_dtd=True)
    rels = fromstring(planilha.read("xl/_rels/workbook.xml.rels"), forbid_dtd=True)
    destinos = {r.get("Id"): r.get("Target") for r in rels.findall("{*}Relationship")}
    for aba in livro.findall(".//{*}sheet"):
        if aba.get("name") == nome:
            id_relacao = next(v for k, v in aba.attrib.items() if k.endswith("}id"))
            destino = destinos[id_relacao]
            return destino.lstrip("/") if destino.startswith("/") else f"xl/{destino}"
    raise ValueError(f"Aba '{nome}' não encontrada na planilha.")


def carregar(caminho):
    with ZipFile(caminho) as planilha:
        textos = "xl/sharedStrings.xml"
        return captura.extrair_dados(
            planilha.read(caminho_da_aba(planilha, "Registro")),
            planilha.read(caminho_da_aba(planilha, "Incerteza")),
            planilha.read(caminho_da_aba(planilha, "Certificado")),
            planilha.read(textos) if textos in planilha.namelist() else None,
        )


def gerar_pdf_sem_erro(dados, resultados, parecer):
    try:
        pdf_bytes = certificado.gerar_certificado_pdf(dados, resultados, parecer, "TESTE")
        return len(pdf_bytes) > 0
    except Exception as erro:
        print(f"      erro ao gerar PDF: {erro}")
        return False


def verificar_cabecalho_comum(dados):
    checar('cliente_solicitante = "Usinagem II"', dados["cliente_solicitante"] == "Usinagem II")
    checar('codigo = "703.2045"', dados["codigo"] == "703.2045")
    checar(
        'tecnico_executante = "Marcos Vinícius Pereira"',
        dados["tecnico_executante"] == "Marcos Vinícius Pereira",
    )
    checar("fator_k = 2.0", dados["fator_k"] == 2.0)
    checar("incerteza_medicao = 0.001", dados["incerteza_medicao"] == 0.001)
    checar("signatario preenchido", bool(dados["signatario"]))
    checar("cargo_signatario preenchido", bool(dados["cargo_signatario"]))
    checar("observacoes preenchido", bool(dados["observacoes"]))
    checar("notas tem 4 itens", len(dados["notas"]) == 4)
    checar('data formata para "14/10/2026"', formatacao.formatar_data(dados["data"]) == "14/10/2026")
    checar("vieram 6 parâmetros", len(dados["parametros"]) == 6)
    checar(
        "medida_1 e medida_2 preenchidos nos 6 parâmetros",
        all(p["medida_1"] is not None and p["medida_2"] is not None for p in dados["parametros"]),
    )


print("=" * 70)
print("PLANILHA DE TESTE — APROVADO")
print("=" * 70)
dados_aprovado = carregar(ARQUIVO_APROVADO)
verificar_cabecalho_comum(dados_aprovado)
resultados_aprovado, parecer_aprovado, motivo_aprovado = calculos.calcular_resultados(dados_aprovado["parametros"])
checar('parecer = "APROVADO"', parecer_aprovado == "APROVADO")
checar("motivo é None quando aprovado", motivo_aprovado is None)
checar("PDF gerado sem erro (aprovado)", gerar_pdf_sem_erro(dados_aprovado, resultados_aprovado, parecer_aprovado))

print()
print("=" * 70)
print("PLANILHA DE TESTE — REPROVADO")
print("=" * 70)
dados_reprovado = carregar(ARQUIVO_REPROVADO)
verificar_cabecalho_comum(dados_reprovado)
resultados_reprovado, parecer_reprovado, motivo_reprovado = calculos.calcular_resultados(dados_reprovado["parametros"])
checar('parecer = "REPROVADO"', parecer_reprovado == "REPROVADO")

fora_da_tolerancia = {
    r["parametro_avaliado"] for r in resultados_reprovado if r["dentro_da_tolerancia"] is False
}
checar(
    '"40P_POSIÇÃO 1" e "12P_DIÂMETRO 1" fora da tolerância, só eles',
    fora_da_tolerancia == {"40P_POSIÇÃO 1", "12P_DIÂMETRO 1"},
)
checar(
    "PDF gerado sem erro (reprovado)", gerar_pdf_sem_erro(dados_reprovado, resultados_reprovado, parecer_reprovado),
)

print()
print("=" * 70)
print("ARQUIVO REAL DA SCHULZ (prova de verdade da Parte 1 — 14 características)")
print("=" * 70)
if not ARQUIVO_REAL.exists():
    print(f"[PULADO] arquivo não encontrado em {ARQUIVO_REAL}")
else:
    dados_real = carregar(ARQUIVO_REAL)
    checar("fator_k = 2.23", dados_real["fator_k"] == 2.23)
    checar("incerteza_medicao = 0.002", dados_real["incerteza_medicao"] == 0.002)
    checar("observacoes preenchido", bool(dados_real["observacoes"]))
    checar(
        'signatario = "João Carlos Gomes de Oliveira"',
        dados_real["signatario"] == "João Carlos Gomes de Oliveira",
    )
    checar('cargo_signatario = "Signatário Autorizado"', dados_real["cargo_signatario"] == "Signatário Autorizado")
    checar("notas tem 4 itens", len(dados_real["notas"]) == 4)
    checar('codigo = "703.1208"', dados_real["codigo"] == "703.1208")
    checar("14 parâmetros", len(dados_real["parametros"]) == 14)
    checar(
        "medida_1 e medida_2 preenchidos nos 14",
        all(p["medida_1"] is not None and p["medida_2"] is not None for p in dados_real["parametros"]),
    )

    resultados_real, parecer_real, motivo_real = calculos.calcular_resultados(dados_real["parametros"])
    checar('parecer = "REPROVADO"', parecer_real == "REPROVADO")
    quantidade_fora = sum(1 for r in resultados_real if r["dentro_da_tolerancia"] is False)
    checar("6 das 14 características fora da tolerância", quantidade_fora == 6)
    checar("PDF gerado sem erro (arquivo real)", gerar_pdf_sem_erro(dados_real, resultados_real, parecer_real))

print()
print("=" * 70)
print("RESULTADO FINAL:", "HÁ FALHAS" if houve_falha else "TUDO OK")
print("=" * 70)
sys.exit(1 if houve_falha else 0)
