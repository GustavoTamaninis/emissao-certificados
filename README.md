# Captura dos campos de calibração

```powershell
python -m pip install -r requirements.txt
```

```python
from captura import extrair_dados

dados = extrair_dados(xml_certificado, xml_incerteza, xml_textos)
```

Os argumentos são conteúdos XML em `str` ou `bytes`, não caminhos de arquivos.
O formato é o XML interno do modelo `Dispositivo de controle.xlsx`, com as
abas Certificado e Incerteza. Não é um formato genérico de outro sistema.

Para obter esses XMLs da planilha anexada:

```python
from zipfile import ZipFile
from captura import extrair_dados

with ZipFile("Dispositivo de controle.xlsx") as planilha:
    dados = extrair_dados(
        planilha.read("xl/worksheets/sheet3.xml"),  # Certificado
        planilha.read("xl/worksheets/sheet2.xml"),  # Incerteza
        planilha.read("xl/sharedStrings.xml"),
    )
```

Os caminhos acima correspondem ao anexo; podem mudar em outras planilhas.
`xml_textos` é opcional quando os campos não usam textos compartilhados.

O retorno contém somente os campos pedidos, a lista `parametros` com as quatro
colunas e `dispersao_medidas` com as sete ocorrências do modelo, identificadas
pelas células. Os valores continuam como textos, preservando a precisão;
células ausentes ou vazias retornam `None`.

As dispersões são: `B11` (linear), `B48` (volumétrico), `B82` (ângulo),
`B117` (segundo bloco volumétrico), `B146` (projetor linear), `B177`
(projetor angular) e `B204` (micrômetro). A função não escolhe um bloco.

As fórmulas não são calculadas. Seus valores salvos podem estar desatualizados;
erros e marcadores do modelo, como `#NAME?`, `----` e `00/00/0000`, são
preservados. O anexo não tem parâmetros preenchidos, portanto a lista extraída
dele é vazia. A função não gera certificados nem sanitiza os dados.
