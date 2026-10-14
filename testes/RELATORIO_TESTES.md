# Relatório de testes manuais — emissão de certificados

Como não temos acesso a uma planilha real já preenchida pelos técnicos da Schulz,
este relatório documenta os testes feitos a partir da planilha-modelo real
(`Dispositivo de controle (2) (1).xlsx`, fornecida em 2026-10-13), com linhas de
medição inseridas manualmente para simular o preenchimento do dia a dia.

Arquivos desta pasta:
- `gerar_planilha_teste.py` — gera `planilha_teste_preenchida.xlsx` a partir da
  planilha-modelo real da Schulz, inserindo 4 linhas de medição manualmente
  (edição direta do XML interno, sem usar openpyxl). Para regerar, ajuste a
  constante `ORIGINAL` no topo do script para o caminho da planilha-modelo.
- `planilha_teste_preenchida.xlsx` — resultado do script acima.
- `executar_teste_completo.py` — roda `captura.extrair_dados` →
  `calculos.calcular_resultados` → `certificado.gerar_certificado_pdf` sobre a
  planilha de teste e grava `resultado_teste.json` e `certificado_teste.pdf`.
- `resultado_teste.json` — saída completa da extração e do cálculo.
- `certificado_teste.pdf` — PDF gerado a partir desses dados.

Pra rodar de novo: `.venv\Scripts\python.exe testes\gerar_planilha_teste.py` e
depois `.venv\Scripts\python.exe testes\executar_teste_completo.py`, a partir da
raiz do projeto.

## 1. Teste de extração de campos (captura.py)

Rodado sobre a planilha real original (sem as linhas de teste), antes de inserir
qualquer dado. Encontramos 5 endereços de célula divergentes do que havia sido
especificado inicialmente — a planilha real usa um layout diferente do assumido.
Todos os campos abaixo foram corrigidos no `captura.py` e revalidados:

| Campo                | Endereço usado antes | Endereço real confirmado | Valor extraído após a correção |
|-----------------------|----------------------|---------------------------|----------------------------------|
| `observacoes`         | A54                  | **A49**                   | "*Os parâmetros avaliados foram obtidos da Ficha de Auto Controle do produto que o dispositivo controla." |
| `signatario`          | F61                  | **F56**                   | "João Carlos Gomes de Oliveira" |
| `cargo_signatario`    | F62                  | **F57**                   | "Signatário Autorizado" |
| `fator_k`             | F50                  | **F45**                   | 2.0 |
| `notas` (4 linhas)    | A64 a A67            | **A59 a A62**              | as 4 notas de rodapé sobre incerteza/reprodução |

Os endereços A64-A66 na planilha real não são notas: são o texto "Folha 01/01" e
o endereço/CNPJ da Schulz — dado sensível que o enunciado original pediu
explicitamente para excluir (só que apontando para A70/A71, que nesta planilha
não existem com esse conteúdo). Antes da correção, esse texto estava indo para
dentro do PDF via o campo `notas`; depois da correção, não aparece mais.

Os campos abaixo já estavam certos desde a primeira tentativa e não precisaram de
ajuste: `numero_certificado_planilha` (F6), `certificado_padrao` (A21),
`procedimento` (D26), `metodo` (A28).

## 2. Teste de cálculo (calculos.py) — dados inseridos manualmente

A planilha real estava com a tabela "Resultados obtidos (mm)" totalmente vazia
(é um modelo em branco). Para testar o cálculo de média/erro/tolerância e o
parecer geral, inserimos manualmente 4 linhas na aba Registro (linhas 30 a 33,
que são as 4 primeiras linhas da tabela de parâmetros), escolhidas para cobrir
os seguintes cenários:

| # | Parâmetro              | Valor nominal | Tol. + | Tol. - | Medida 1 | Medida 2 | Por que esse caso |
|---|-------------------------|---------------|--------|--------|----------|----------|---|
| 1 | Diâmetro A              | 10,0000       | 0,0100 | -0,0100| 10,0020  | 10,0030  | Força média = 10,0025, pra testar o arredondamento `ROUND_HALF_UP` na exibição (3ª casa decimal em cima de um ",5" exato) |
| 2 | Posição do furo A       | 5,0000        | 0,1000 | 0,0000 | 5,0400   | 5,0600   | Tolerância assimétrica normal (0 a +0,1), dentro da faixa |
| 3 | Posição do rebaixo B    | 8,0000        | 0,1000 | 0,0000 | 7,9400   | 7,9600   | Erro **negativo** (-0,05) com `tol_menos = 0`: deve reprovar. Se o cálculo comparasse `abs(erro) <= tol_mais` (0,05 ≤ 0,1) aprovaria **errado** |
| 4 | Diâmetro B              | 20,0000       | 0,0050 | -0,0050| 20,0120  | 20,0140  | Fora da tolerância de forma direta, sem ambiguidade (erro = 0,013 > 0,005) |

### Resultado esperado (calculado à mão) vs. obtido pelo sistema

| # | Parâmetro | Média esperada | Erro esperado | Dentro da tolerância? | Resultado do sistema |
|---|---|---|---|---|---|
| 1 | Diâmetro A | 10,0025 | 0,0025 | Sim | `media=10.0025 erro=0.0025 dentro=True` ✅ |
| 2 | Posição do furo A | 5,0500 | 0,0500 | Sim | `media=5.0500 erro=0.0500 dentro=True` ✅ |
| 3 | Posição do rebaixo B | 7,9500 | -0,0500 | **Não** | `media=7.9500 erro=-0.0500 dentro=False` ✅ |
| 4 | Diâmetro B | 20,0130 | 0,0130 | **Não** | `media=20.0130 erro=0.0130 dentro=False` ✅ |

**Parecer geral esperado:** REPROVADO (pelo menos um parâmetro fora: #3 e #4).
**Parecer geral obtido:** `REPROVADO` ✅

O caso #3 confirma que a tolerância assimétrica está sendo tratada corretamente
— uma implementação com `abs(erro) <= tol_mais` teria aprovado esse parâmetro
por engano.

## 3. Teste de formatação (formatacao.py)

```
formatar_numero(Decimal("10.0025"))  -> "10,003"   (round() nativo daria "10.002", errado)
formatar_numero(Decimal("0.0025"))   -> "0,003"
formatar_data("2026-08-26T16:47:10.00000000465661175") -> "26/08/2026"
formatar_data("00/00/0000")          -> "00/00/0000"  (texto cru, não quebra)
```

Confirma o uso de `Decimal` + `ROUND_HALF_UP` em vez do `round()` nativo, e que
datas no modelo em branco (`00/00/0000`) não quebram a formatação.

## 4. Teste de geração de PDF (certificado.py)

PDF gerado com sucesso (`certificado_teste.pdf`, ~3,7 KB) a partir dos dados
acima, número de certificado "TESTE-001/2026". Itens confirmados na revisão
visual do PDF:
- Faixa superior sem logo, sem endereço/CNPJ da Schulz.
- Parecer final "REPROVADO" em destaque (vermelho), antes da tabela.
- Tabela de resultados com as 6 colunas na ordem pedida, bordas e cabeçalho com
  fundo cinza.
- Campos com placeholder da planilha em branco ("----", "00/00/0000") aparecem
  como estão, sem quebrar e sem "N/A".
- Notas de rodapé corretas (sem o texto de endereço/CNPJ da Schulz).

## 5. O que ainda não foi validado

- Não testamos com uma planilha realmente preenchida por um técnico da Schulz
  (não temos acesso a uma). Os valores de medição usados aqui são sintéticos,
  inseridos manualmente editando o XML interno do `.xlsx` (mesma técnica que o
  `captura.py` usa pra ler, sem depender de openpyxl).
- Não testamos o fluxo pela interface do Streamlit com upload manual — só pelo
  mesmo caminho que o `app.py` usa internamente (`captura.extrair_dados` →
  `calculos.calcular_resultados` → `certificado.gerar_certificado_pdf`).
- Se os endereços de célula mudarem em versões futuras do modelo da Schulz, os
  campos vão silenciosamente vir vazios ou errados — não há validação de
  planilha inválida (decisão do enunciado: fora de escopo desta entrega).
