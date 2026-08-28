@echo off
REM ===================================================================
REM  Gerador-OS - instalar o certificado da fabrica nesta maquina.
REM
REM  POR QUE ISTO EXISTE
REM  O servidor da fabrica atende em https://193.168.0.200 com um
REM  certificado assinado por uma autoridade PROPRIA ("Gerador-OS - CA
REM  da fabrica"), que nenhum computador conhece de nascenca. Sem
REM  instalar essa autoridade, o navegador abre a tela vermelha de
REM  "site nao seguro" e o login nem chega a acontecer.
REM
REM  Instalar e copiar UM arquivo para a lista de autoridades confiaveis
REM  do Windows. Nao muda mais nada na maquina, e vale para o Chrome, o
REM  Edge e para o Windows inteiro (o Firefox tem lista propria - ver o
REM  aviso no fim).
REM
REM  SEM ACENTO DE PROPOSITO: arquivo .cmd com acento sai embaralhado em
REM  maquina com outra configuracao de idioma, e este roda em maquina
REM  que ninguem conhece.
REM
REM  COMO USAR: clique duas vezes. Ele pede a permissao de
REM  administrador sozinho (a lista de autoridades e do computador, nao
REM  do usuario) e diz no fim se deu certo.
REM ===================================================================
setlocal EnableDelayedExpansion
title Gerador-OS - instalar o certificado
color 0F

REM A impressao digital do CA de verdade. E por ela que o fim confere se
REM o que entrou na lista foi ESTE certificado, e nao outro com o mesmo
REM nome de arquivo.
set "DIGITAL=3581656E6C0EABC4FD52989652B59A56715AF356"
REM O servidor pode ser alcancado por dois caminhos: o cabo de rede (o de
REM sempre) e o Wi-Fi (o contorno de 28/08/2026, com o cabo fora). O mesmo
REM certificado vale para os dois. O instalador testa os dois e diz qual
REM responde nesta maquina - dizer "nao respondeu" quando o outro caminho
REM funciona seria mandar a pessoa procurar um problema que nao existe.
set "SERVIDOR=https://193.168.0.200"
set "ALTERNATIVO=https://192.168.1.158"

echo.
echo   ================================================
echo    GERADOR-OS - certificado do servidor da fabrica
echo   ================================================
echo.

REM ---- 1. achar o certificado ao lado deste arquivo -------------------
set "CERT=%~dp0ca.crt"
if not exist "%CERT%" (
  echo   [ERRO] Nao achei o arquivo "ca.crt" nesta pasta.
  echo.
  echo   Os DOIS arquivos tem de estar juntos:
  echo       instalar-certificado.cmd
  echo       ca.crt
  echo.
  echo   Copie a pasta inteira, nao so este arquivo.
  echo.
  pause
  exit /b 1
)

REM ---- 2. permissao de administrador ---------------------------------
REM Depois de achar o arquivo, nao antes: quem copiou so o .cmd merece a
REM mensagem do que falta, e nao uma janela de permissao que nao leva a nada.
REM `net session` so funciona como administrador; e o teste mais antigo e
REM mais confiavel que existe no Windows para isto.
net session >nul 2>&1
if errorlevel 1 (
  echo   Preciso da permissao de administrador para gravar na lista de
  echo   autoridades do computador. Vou pedir agora - clique em SIM.
  echo.
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
  if errorlevel 1 (
    echo   Nao consegui pedir a permissao.
    echo   Clique com o botao DIREITO neste arquivo e escolha
    echo   "Executar como administrador".
    echo.
    pause
  )
  exit /b
)

REM ---- 3. instalar ---------------------------------------------------
echo   Instalando...
echo.
certutil -addstore -f Root "%CERT%" >nul 2>&1
if errorlevel 1 (
  echo   [ERRO] O Windows recusou a instalacao.
  echo   Tente de novo com o botao direito ^> "Executar como administrador".
  echo.
  pause
  exit /b 1
)

REM ---- 4. conferir que foi ESTE certificado que entrou ----------------
REM Instalar sem conferir e supor. A impressao digital e o unico jeito
REM de saber que a lista tem o certificado certo, e nao um homonimo.
certutil -verifystore Root %DIGITAL% >nul 2>&1
if errorlevel 1 (
  echo   [ERRO] O certificado nao aparece na lista depois de instalado.
  echo   Chame o Junior - alguma coisa fora do comum aconteceu aqui.
  echo.
  pause
  exit /b 1
)
echo   OK - certificado instalado e conferido.

REM ---- 5. atalho na area de trabalho de todo mundo --------------------
REM Um .url simples: nao depende de qual navegador e o padrao, e some com
REM um delete se nao quiserem.
set "ATALHO=%PUBLIC%\Desktop\Gerador-OS.url"
>"%ATALHO%" echo [InternetShortcut]
>>"%ATALHO%" echo URL=%SERVIDOR%/
if exist "%ATALHO%" (echo   OK - atalho "Gerador-OS" criado na area de trabalho.)

REM ---- 6. o servidor esta respondendo agora? --------------------------
REM Instalado o certificado, a pergunta seguinte e sempre "e agora, abre?".
REM Melhor responder aqui do que deixar a pessoa descobrir no navegador.
echo.
echo   Testando o servidor...
set "ACHOU="
call :testar "%SERVIDOR%"    && set "ACHOU=%SERVIDOR%"
if not defined ACHOU call :testar "%ALTERNATIVO%" && set "ACHOU=%ALTERNATIVO%"

if defined ACHOU (
  echo   OK - o servidor respondeu em %ACHOU%
  if not "%ACHOU%"=="%SERVIDOR%" (
    echo.
    echo   [AVISO] Respondeu pelo endereco ALTERNATIVO, nao pelo de sempre.
    echo   Use %ACHOU% enquanto o de sempre nao voltar.
    >"%ATALHO%" echo [InternetShortcut]
    >>"%ATALHO%" echo URL=%ACHOU%/
  )
  goto fim
)

echo   [ATENCAO] O certificado esta instalado, mas o servidor NAO respondeu
echo   em nenhum dos dois enderecos.
echo.
echo   Isso nao e problema desta maquina. Quase sempre e uma destas:
echo     - o cabo de rede do servidor esta fora do lugar;
echo     - esta maquina esta em outra rede (Wi-Fi de visitante, p.ex.^);
echo     - o servidor esta desligado.
echo.
echo   Avise o Junior.
goto fim

REM Devolve 0 quando o endereco respondeu 200. Sub-rotina para nao repetir a
REM linha do powershell duas vezes - repetida, uma seria corrigida sem a outra.
:testar
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $r = Invoke-WebRequest -Uri '%~1/index.html' -TimeoutSec 6 -UseBasicParsing; if ($r.StatusCode -eq 200) { exit 0 } else { exit 1 } } catch { exit 1 }" >nul 2>&1
exit /b %errorlevel%

:fim
echo.
echo   ------------------------------------------------
echo    Pronto. Abra o atalho "Gerador-OS" na area de
echo    trabalho, ou digite  %SERVIDOR%
echo    (se o de sempre nao responder: %ALTERNATIVO%^)
echo.
echo    Se esta maquina usa FIREFOX, ele tem uma lista
echo    propria de certificados e precisa de um passo a
echo    mais - avise o Junior. Chrome e Edge ja estao
echo    prontos.
echo   ------------------------------------------------
echo.
pause
