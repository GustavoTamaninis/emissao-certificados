@echo off
setlocal
REM Executa sempre a partir da pasta do script, mesmo com espacos no caminho
cd /d "%~dp0"

TITLE Inicializador de Aplicacao Schulz S.A. - AV 05
cls
echo ============================================================
echo    INICIALIZANDO APLICACAO SCHULZ S.A. - AV 05
echo    Emissao de Certificados de Calibracao
echo ============================================================
echo.

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "PORTA=8501"

REM ------------------------------------------------------------
REM [1/4] Checagem de dependencias: Python 3.10+ no PATH
REM ------------------------------------------------------------
echo [1/4] Verificando instalacao do Python 3.10+...

REM Tenta o comando "python"; se nao existir (ou for o atalho da
REM Microsoft Store), tenta o inicializador "py" do Windows.
set "PYTHON="
python -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=python"
) else (
    py -3 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PYTHON=py -3"
)

if not defined PYTHON (
    echo.
    echo [ERRO] Python nao foi encontrado no PATH do sistema.
    echo        Instale o Python 3.10 ou superior em https://www.python.org/downloads/
    echo        e marque a opcao "Add python.exe to PATH" durante a instalacao.
    goto :erro
)

%PYTHON% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 (
    echo.
    for /f "delims=" %%v in ('%PYTHON% --version 2^>^&1') do echo [ERRO] Versao encontrada: %%v
    echo        Este projeto exige Python 3.10 ou superior.
    echo        Atualize o Python em https://www.python.org/downloads/
    goto :erro
)

for /f "delims=" %%v in ('%PYTHON% --version 2^>^&1') do echo [OK] %%v detectado no sistema.

REM O Windows limita caminhos a 260 caracteres e o Streamlit instala arquivos
REM com caminhos longos. Em pastas muito profundas a instalacao falharia, entao
REM avisa antes (a menos que o suporte a caminhos longos esteja ativado).
if "%CD:~110,1%"=="" goto :caminho_ok
reg query "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /v LongPathsEnabled 2>nul | findstr /r /c:"0x1$" >nul
if not errorlevel 1 goto :caminho_ok
echo.
echo [ERRO] O caminho da pasta do projeto e longo demais para o Windows:
echo        %CD%
echo        Mova a pasta para um local mais curto, por exemplo C:\schulz,
echo        e execute este arquivo novamente.
goto :erro
:caminho_ok
echo.

REM ------------------------------------------------------------
REM [2/4] Criacao e ativacao do ambiente virtual
REM ------------------------------------------------------------
echo [2/4] Preparando o ambiente virtual (%VENV_DIR%)...

REM Se a pasta existir mas estiver quebrada (ex.: copiada de outro
REM computador), ela e apagada e criada de novo.
if exist "%VENV_PY%" (
    "%VENV_PY%" -c "import sys" >nul 2>&1
    if errorlevel 1 (
        echo        Ambiente virtual invalido encontrado. Recriando...
        rmdir /s /q "%VENV_DIR%"
    )
)

if not exist "%VENV_PY%" (
    echo        Criando ambiente virtual. Isso pode levar alguns segundos...
    %PYTHON% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo.
        echo [ERRO] Nao foi possivel criar o ambiente virtual.
        goto :erro
    )
    echo [OK] Ambiente virtual criado com sucesso.
) else (
    echo [OK] Ambiente virtual ja existente.
)

REM CALL e obrigatorio: sem ele o script terminaria junto com o activate.bat
call "%VENV_DIR%\Scripts\activate.bat"
echo [OK] Ambiente virtual ativado.
echo.

REM ------------------------------------------------------------
REM [3/4] Instalacao das bibliotecas do requirements.txt
REM ------------------------------------------------------------
echo [3/4] Instalando bibliotecas do requirements.txt...

if not exist "requirements.txt" (
    echo.
    echo [ERRO] Arquivo requirements.txt nao encontrado na pasta do projeto.
    goto :erro
)

python -m pip install --upgrade pip --quiet --disable-pip-version-check >nul 2>&1
python -m pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar as bibliotecas.
    echo        Verifique sua conexao com a internet e a mensagem acima.
    goto :erro
)
echo [OK] Bibliotecas instaladas.
echo.

REM ------------------------------------------------------------
REM [4/4] Inicializacao dos servicos
REM ------------------------------------------------------------
echo [4/4] Iniciando a aplicacao...

REM Se a porta 8501 estiver ocupada por outro programa, usa a proxima livre
:procurar_porta
netstat -ano | findstr /r /c:":%PORTA% .*LISTENING" >nul 2>&1
if errorlevel 1 goto :porta_livre
echo        Porta %PORTA% ocupada, tentando a proxima...
set /a PORTA+=1
if %PORTA% LSS 8520 goto :procurar_porta
echo.
echo [ERRO] Nenhuma porta livre encontrada entre 8501 e 8519.
goto :erro

:porta_livre

REM Servidor em uma janela propria. O modo headless evita que o Streamlit
REM pare pedindo e-mail no primeiro uso; o navegador e aberto por este script.
start "Servidor Schulz S.A. - porta %PORTA%" cmd /k python -m streamlit run app.py --server.port %PORTA% --server.headless true --browser.gatherUsageStats false

set "URL=http://localhost:%PORTA%"
echo        Aguardando o servidor responder em %URL% ...

where curl >nul 2>&1
if errorlevel 1 (
    REM Sem curl (Windows antigo): espera fixa antes de abrir o navegador
    ping -n 9 127.0.0.1 >nul
    goto :abrir_navegador
)

set /a TENTATIVAS=0
:aguardar_servidor
curl -s -f -o nul "%URL%/_stcore/health" >nul 2>&1
if not errorlevel 1 goto :abrir_navegador
set /a TENTATIVAS+=1
if %TENTATIVAS% GEQ 60 (
    echo.
    echo [ERRO] O servidor nao respondeu em 60 segundos.
    echo        Veja a mensagem de erro na janela "Servidor Schulz S.A.".
    goto :erro
)
REM Pausa de ~1 segundo (ping funciona mesmo quando o timeout nao funciona)
ping -n 2 127.0.0.1 >nul
goto :aguardar_servidor

:abrir_navegador
start "" "%URL%"
echo [OK] Servidor ativo e interface aberta no navegador padrao.
echo.
echo ============================================================
echo    APLICACAO EM EXECUCAO
echo    Endereco: %URL%
echo.
echo    Para encerrar, feche a janela "Servidor Schulz S.A.".
echo ============================================================
echo.
pause
endlocal
exit /b 0

:erro
echo.
echo A inicializacao foi interrompida.
pause
endlocal
exit /b 1
