@echo off
setlocal enabledelayedexpansion
title Gerador-OS
chcp 65001 >nul 2>&1

REM ======================================================================
REM  Abre o Gerador-OS procurando o servidor, em vez de decorar um numero.
REM
REM  POR QUE ISTO EXISTE
REM  Um atalho comum guarda UM endereco. O servidor da fabrica tem dois --
REM  o cabo (193.168.0.200) e o Wi-Fi (192.168.1.158) -- e troca de um para
REM  o outro quando o cabo cai. Em 31/08/2026 isso deixou a fabrica inteira
REM  sem o app numa segunda-feira: o atalho apontava para um endereco que
REM  ninguem mais atendia, e a mensagem era "tempo esgotado", que nao diz
REM  nada sobre o motivo.
REM
REM  Este arquivo tenta os caminhos em ordem e abre o PRIMEIRO que responde.
REM  Quem troca de rede -- o servidor ou esta maquina -- nao precisa avisar
REM  ninguem nem mexer em atalho nenhum.
REM
REM  A ORDEM IMPORTA
REM  O NOME vem primeiro de proposito. Um numero pertence a UMA rede; um
REM  nome, nao. O Windows resolve o nome da maquina sozinho (LLMNR/NetBIOS)
REM  na rede em que esta maquina estiver, entao o nome acerta o endereco
REM  certo sem que ninguem escolha. Os dois IPs ficam abaixo dele como rede
REM  de seguranca, para o caso de a resolucao de nome estar desligada por
REM  politica nesta maquina.
REM
REM  O certificado cobre os tres (reemitido em 31/08 com a MESMA autoridade,
REM  entao nenhuma maquina precisou reinstalar o ca.crt).
REM
REM  PARA MUDAR OS CAMINHOS: edite a linha ALVOS abaixo, so ela.
REM ======================================================================

set "ALVOS=DESKTOP-SOV61AF 193.168.0.200 192.168.1.158"

echo.
echo   Gerador-OS - procurando o servidor...
echo.

set "ACHADO="
for %%A in (%ALVOS%) do (
  if not defined ACHADO (
    REM  -k ignora o certificado SO NESTA SONDAGEM: aqui a pergunta e "tem
    REM  alguem atendendo?", e nao "e de confianca?". Quem valida o
    REM  certificado de verdade e o navegador, logo abaixo -- e ai o
    REM  cadeado tem de aparecer limpo.
    curl.exe -s -k --max-time 2 -o NUL "https://%%A/" >nul 2>&1
    if !errorlevel! equ 0 (
      set "ACHADO=%%A"
      echo     %%A ... responde
    ) else (
      echo     %%A ... nao responde
    )
  )
)

if not defined ACHADO goto :naoachei

echo.
echo   Abrindo https://!ACHADO!/
echo.

REM O Chrome em modo aplicativo abre sem barra de endereco, com cara de
REM programa e nao de site. Se nao houver Chrome, o navegador padrao serve.
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"

if exist "%CHROME%" (
  start "" "%CHROME%" --app=https://!ACHADO!/
) else (
  start "" "https://!ACHADO!/"
)
exit /b 0

:naoachei
echo.
echo   ==================================================================
echo    NAO ACHEI O SERVIDOR.
echo   ==================================================================
echo.
echo   Nenhum destes respondeu:
for %%A in (%ALVOS%) do echo       %%A
echo.
echo   O que costuma ser:
echo.
echo     1. Esta maquina esta em outra rede que o servidor.
echo        Confira o Wi-Fi / o cabo desta maquina.
echo.
echo     2. O servidor esta desligado, ou o Docker ainda esta subindo.
echo        Ele leva uns 3 minutos depois de ligar.
echo.
echo     3. O servidor trocou de rede e o firewall dele voltou a
echo        classifica-la como Publica -- ai ele so atende a si mesmo.
echo        No servidor, rodar como administrador:
echo            servidor\liberar-portas-firewall.ps1
echo.
pause
exit /b 1
