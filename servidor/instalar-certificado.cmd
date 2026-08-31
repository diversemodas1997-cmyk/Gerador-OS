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
REM  O NOME vem primeiro: um numero pertence a UMA rede, um nome nao. O
REM  Windows resolve o nome da maquina sozinho (LLMNR/NetBIOS) na rede em
REM  que esta maquina estiver, entao ele acerta o endereco certo sem que
REM  ninguem escolha. Os IPs ficam abaixo como rede de seguranca.
REM  O nome antigo fica logo abaixo do novo enquanto houver maquina que
REM  ainda nao viu o reboot da renomeacao. O certificado cobre os dois.
set "PORNOME=https://GERADOR-OS"
set "PORNOMEANTIGO=https://DESKTOP-SOV61AF"
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
REM Ate 31/08/2026 aqui se escrevia um .url com UM endereco fixo. Foi
REM exatamente isso que deixou a fabrica parada numa segunda-feira: o
REM servidor trocou de rede e o atalho continuou apontando para o endereco
REM morto, com "tempo esgotado" por toda explicacao.
REM Agora vai um LANCADOR: ele procura o servidor entre os caminhos
REM conhecidos e abre o primeiro que responde. Quem troca de rede -- o
REM servidor ou esta maquina -- nao precisa avisar ninguem.
set "PASTA=%ProgramData%\Gerador-OS"
set "LANCADOR=%PASTA%\abrir-gerador-os.cmd"
set "ATALHO=%PUBLIC%\Desktop\Gerador-OS.lnk"
set "URLVELHO=%PUBLIC%\Desktop\Gerador-OS.url"
set "TEMLANCADOR="

if exist "%~dp0abrir-gerador-os.cmd" (
  if not exist "%PASTA%" mkdir "%PASTA%" >nul 2>&1
  copy /y "%~dp0abrir-gerador-os.cmd" "%LANCADOR%" >nul 2>&1
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%ATALHO%'); $s.TargetPath='%LANCADOR%'; $s.WindowStyle=7; $s.Description='Abre o Gerador-OS procurando o servidor na rede'; $s.Save()" >nul 2>&1
)
if exist "%ATALHO%" (
  set "TEMLANCADOR=1"
  echo   OK - atalho "Gerador-OS" criado: ele PROCURA o servidor.
  REM O .url antigo, de endereco fixo, seria a mesma armadilha de novo.
  if exist "%URLVELHO%" del /q "%URLVELHO%" >nul 2>&1
) else (
  REM Pacote antigo, sem o lancador ao lado: fica o de antes, que ao menos
  REM abre alguma coisa. Melhor um atalho velho do que nenhum.
  >"%URLVELHO%" echo [InternetShortcut]
  >>"%URLVELHO%" echo URL=%SERVIDOR%/
  set "ATALHO=%URLVELHO%"
  echo   OK - atalho "Gerador-OS" criado na area de trabalho.
)

REM ---- 6. o servidor esta respondendo agora? --------------------------
REM Instalado o certificado, a pergunta seguinte e sempre "e agora, abre?".
REM Melhor responder aqui do que deixar a pessoa descobrir no navegador.
echo.
echo   Testando o servidor...
set "ACHOU="
call :testar "%PORNOME%"     && set "ACHOU=%PORNOME%"
if not defined ACHOU call :testar "%PORNOMEANTIGO%" && set "ACHOU=%PORNOMEANTIGO%"
if not defined ACHOU call :testar "%SERVIDOR%"    && set "ACHOU=%SERVIDOR%"
if not defined ACHOU call :testar "%ALTERNATIVO%" && set "ACHOU=%ALTERNATIVO%"

if defined ACHOU (
  echo   OK - o servidor respondeu em %ACHOU%
  if defined TEMLANCADOR (
    REM Com o lancador, saber QUAL respondeu e informacao, nao decisao: ele
    REM refaz esta busca a cada clique, e por isso acerta tambem amanha.
    echo   O atalho refaz esta busca a cada clique - nao fica preso a este.
  ) else (
    if not "%ACHOU%"=="%SERVIDOR%" (
      echo.
      echo   [AVISO] Respondeu pelo endereco ALTERNATIVO, nao pelo de sempre.
      echo   Use %ACHOU% enquanto o de sempre nao voltar.
      >"%ATALHO%" echo [InternetShortcut]
      >>"%ATALHO%" echo URL=%ACHOU%/
    )
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
echo    trabalho: ele procura o servidor sozinho, pelo
echo    cabo ou pelo Wi-Fi, e nao precisa ser trocado
echo    quando o servidor mudar de rede.
echo.
echo    Se precisar digitar:  %PORNOME%
echo    (ou %SERVIDOR% / %ALTERNATIVO%^)
echo.
echo    Se esta maquina usa FIREFOX, ele tem uma lista
echo    propria de certificados e precisa de um passo a
echo    mais - avise o Junior. Chrome e Edge ja estao
echo    prontos.
echo   ------------------------------------------------
echo.
pause
