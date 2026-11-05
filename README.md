# emissao-certificados

Aplicação para extrair os dados de uma planilha de calibração (.xlsx) e gerar o certificado de calibração.

## Pré-requisitos

- [Python 3.13](https://www.python.org/downloads/)

Para verificar se a versão correta está instalada:

**Windows**

```powershell
py -3.13 --version
```

**Linux**

```bash
python3.13 --version
```

> No Debian/Ubuntu, pode ser necessário instalar também o pacote do módulo de ambientes virtuais: `sudo apt install python3.13-venv`.

## Instalação

1. Crie um ambiente virtual na raiz do projeto:

   **Windows**

   ```powershell
   py -3.13 -m venv .venv
   ```

   **Linux**

   ```bash
   python3.13 -m venv .venv
   ```

2. Ative o ambiente virtual:

   **Windows (PowerShell)**

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

   **Windows (Prompt de Comando)**

   ```bat
   .venv\Scripts\activate.bat
   ```

   **Linux**

   ```bash
   source .venv/bin/activate
   ```

   > Se o PowerShell bloquear a execução do script, rode uma vez:
   > `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

3. Com o ambiente ativado, instale as dependências listadas no `requirements.txt` (o comando é o mesmo nos dois sistemas):

   ```bash
   python -m pip install -r requirements.txt
   ```

## Como executar

Com o ambiente virtual ativado (o comando é o mesmo no Windows e no Linux):

```bash
python -m streamlit run app.py
```

## Dependências

As bibliotecas usadas pelo projeto estão no `requirements.txt`. Sempre que uma nova dependência for adicionada ao código, atualize esse arquivo, com o ambiente virtual ativado:

```bash
python -m pip freeze > requirements.txt
```
