# emissao-certificados

Aplicativo em Streamlit para extrair os dados de uma planilha de calibração (.xlsx) e gerar o certificado de calibração.

## Pré-requisitos

- [Python 3.13+](https://www.python.org/downloads/)

## Instalação

1. Crie um ambiente virtual na raiz do projeto:

   ```powershell
   python -m venv .venv
   ```

2. Ative o ambiente virtual:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

3. Instale as dependências listadas no `requirements.txt`:

   ```powershell
   pip install -r requirements.txt
   ```

## Como executar

Com o ambiente virtual ativado:

```powershell
streamlit run app.py
```

## Dependências

As bibliotecas usadas pelo projeto estão no `requirements.txt`. Sempre que uma nova dependência for adicionada ao código, atualize esse arquivo rodando:

```powershell
pip freeze > requirements.txt
```
# Projeto de Calibração - AV3
