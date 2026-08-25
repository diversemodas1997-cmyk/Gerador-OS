@echo off
REM ============================================================
REM  Instala o "cracha da fabrica" (ca.crt) neste computador.
REM  Faz o https://193.168.0.200 abrir SEM o aviso de
REM  "a conexao nao e particular". E so uma vez por maquina.
REM
REM  Como usar: clique com o botao DIREITO neste arquivo e
REM  escolha "Executar como administrador". (Se der dois
REM  cliques, ele mesmo pede a permissao de administrador.)
REM ============================================================

title Instalar cracha da fabrica (ca.crt)

REM --- Endereco do servidor da fabrica ---
set "SERVIDOR=https://193.168.0.200"
set "DEST=%TEMP%\ca.crt"

REM --- Se ainda nao for administrador, reabre pedindo a permissao ---
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Pedindo permissao de administrador...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
  exit /b
)

echo.
echo ============================================================
echo   Instalando o cracha da fabrica neste computador
echo ============================================================
echo.

echo [1/2] Baixando o cracha de %SERVIDOR%/ca.crt ...
powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Add-Type -TypeDefinition 'using System.Net;using System.Security.Cryptography.X509Certificates;public class TrustAllBat:ICertificatePolicy{public bool CheckValidationResult(ServicePoint s,X509Certificate c,WebRequest r,int p){return true;}}'; [System.Net.ServicePointManager]::CertificatePolicy=New-Object TrustAllBat; Invoke-WebRequest -Uri '%SERVIDOR%/ca.crt' -OutFile '%DEST%' -UseBasicParsing -TimeoutSec 15 } catch { exit 1 }"

if not exist "%DEST%" (
  REM Nao baixou pela rede. Tenta um ca.crt ja salvo em Downloads.
  if exist "%USERPROFILE%\Downloads\ca.crt" (
    copy /y "%USERPROFILE%\Downloads\ca.crt" "%DEST%" >nul
  )
)

if not exist "%DEST%" (
  echo.
  echo NAO consegui obter o cracha.
  echo   - Confira se esta maquina esta na rede da fabrica.
  echo   - Ou baixe manualmente: abra %SERVIDOR%/ca.crt no navegador,
  echo     salve em Downloads e rode este arquivo de novo.
  echo.
  pause
  exit /b 1
)

echo [2/2] Instalando na lista de autoridades confiaveis do Windows ...
certutil -addstore -f Root "%DEST%" >nul
if %errorlevel% neq 0 (
  echo.
  echo FALHOU ao instalar. Rode este arquivo com o botao direito
  echo   ^> "Executar como administrador".
  echo.
  pause
  exit /b 1
)

del "%DEST%" >nul 2>&1

echo.
echo ============================================================
echo   PRONTO! O cracha foi instalado.
echo.
echo   Agora FECHE e reabra o navegador (Chrome ou Edge) e acesse
echo   %SERVIDOR%
echo   O cadeado deve aparecer sem nenhum aviso.
echo ============================================================
echo.
pause
