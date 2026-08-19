<#
  Tira da inicializacao do Windows o que nao tem o que fazer num servidor, e
  poe um teto de memoria no WSL.

  POR QUE EXISTE:
  em 13/08/2026 a fabrica esperou 15 minutos pelo programa numa manha em que
  NADA estava quebrado. O desligamento da vespera foi limpo, o vigia agiu na
  hora certa - a maquina e que estava sufocada: 99 MB de RAM livres de 8 GB, e
  16,9 GB de 19,6 comprometidos. Nesse estado o motor do Docker levou 5m18s
  para responder (o normal e ~40 s), a propria API do Docker devolveu erro 500
  em comandos triviais, e o "up -d" da pilha do app FALHOU as 07:32.

  Nao adianta apertar os tempos do vigia se a maquina nao tem ar. Isto aqui e o
  conserto na raiz, e nao mexe em nada enquanto a fabrica trabalha: as duas
  mudancas so valem no PROXIMO logon.

  SEGUNDA PASSADA, 19/08/2026: a primeira nao bastou. Seis dias depois a
  maquina amanhecia com 211 MB livres e o motor do Docker levava 432 s (e 1220 s
  em 18/08). O que a primeira passada nao alcancava e que aqui entra:
  a inicializacao do USUARIO ja estava enxuta, mas a da MAQUINA nao — Adobe
  Creative Cloud (~300 MB em 9 processos), Radeon Software (~287 MB, incluindo
  o gravador de tela ReLive) e os monitores da impressora Brother sobem por
  HKLM e por tarefa agendada, onde o -Desfazer da primeira versao nunca olhou.
  Sao ~680 MB, todo logon, num PC de 8 GB. Essa parte exige administrador.

  O TETO DO WSL FICA EM 4 GB, e isso foi medido, nao chutado: os 12 conteineres
  somam 2,0 GB (o maior e o kong, com 556 MB). Baixar o teto para 3 GB nao
  devolveria nada ao Windows — a VM ja trabalha perto disso — e cobraria swap
  justamente na hora do arranque.

  O QUE NAO E MEXIDO, DE PROPOSITO:
    GoogleDriveFS  - os dois backups diarios gravam em J:\Meu Drive. Sem ele,
                     a fabrica fica sem backup, que e pior que ficar lenta.
    Audaces 360    - e software de trabalho, com licenca; nao e enfeite.
    CodeMeter      - e a licenca do Audaces. Sem ele o Audaces nao abre.
    PrintNode      - impressao; e o que poe a etiqueta no papel.
    Seagull Drivers- driver da impressora de etiquetas, pelo mesmo motivo.
    RotinasEscritorio-AutoStart - programa do proprio Junior.
    Docker Desktop - ja esta desligado, e assim tem de ficar: quem abre o Docker
                     e o vigia-docker.ps1, e dois donos da mesma decisao brigam.
    Defender/Topaz - seguranca. Lentidao nao se conserta desligando protecao.

  Rodar (sem administrador, na conta do servidor):
    .\servidor\enxugar-inicializacao.ps1
  Rodar a parte da MAQUINA (pede o UAC uma vez, e o de clicar "Sim"):
    .\servidor\enxugar-inicializacao.ps1 -Maquina
  Fechar agora o que ja esta aberto, sem esperar o proximo logon:
    .\servidor\enxugar-inicializacao.ps1 -FecharAgora
  Voltar tudo como estava (as duas partes):
    .\servidor\enxugar-inicializacao.ps1 -Desfazer -Maquina
#>
[CmdletBinding()]
param(
  [switch] $Desfazer,
  # A parte da maquina (HKLM e tarefas agendadas). Separada porque pede UAC, e
  # a parte do usuario tem de continuar rodando sem ninguem clicar em nada.
  [switch] $Maquina,
  # Fecha agora o que foi tirado da inicializacao, em vez de esperar o proximo
  # logon. Sao agentes e paineis: nenhum deles tem trabalho salvo para perder.
  [switch] $FecharAgora,
  # Teto da VM do WSL. O padrao do WSL e "metade da RAM", que nesta maquina da
  # os mesmos 4 GB - o ganho real nao e o teto, e o autoMemoryReclaim: sem ele
  # o WSL NUNCA devolve ao Windows a memoria que os conteineres pararam de usar.
  [string] $MemoriaWsl = '4GB'
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\enxugar-inicializacao.log'

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  for ($i = 0; $i -lt 5; $i++) {
    try { Add-Content -Path $Log -Value $linha -Encoding utf8; break }
    catch { Start-Sleep -Milliseconds 200 }
  }
  Write-Host $linha
}

# Cada linha e uma entrada do Run do usuario, com o motivo de sair. O motivo fica
# aqui e nao no commit porque quem for reverter isto daqui a um ano vai abrir
# este arquivo, nao o historico do git.
$Alvos = @(
  @{ nome = 'Discord';                                                    porque = 'mensageiro; 518 MB em tres processos' }
  @{ nome = 'MicrosoftEdgeAutoLaunch_65E153DD33F59AC373DA09E8D899FD40';   porque = 'abre o Edge sozinho no logon' }
  @{ nome = 'CanvaAutoLaunchAvailabilityCheckAgent';                      porque = 'agente do Canva, so verifica atualizacao' }
  @{ nome = 'electron.app.BlueStacks Services';                           porque = 'emulador de Android' }
  @{ nome = 'AMDNoiseSuppression';                                        porque = 'filtro de ruido de microfone' }
)

$Run      = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$Aprovado = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'

# O mesmo mecanismo que o Gerenciador de Tarefas usa: um valor de 12 bytes em
# StartupApproved\Run, onde o primeiro byte manda. 2 = habilitado, 3 = desligado
# pelo usuario. A entrada em ...\Run continua existindo e intacta - e por isso
# que desfazer e trivial, e por isso que olhar so o Run engana (foi o que
# escondeu o AutoStart do Docker em 11/08/2026).
$LIGADO   = [byte[]](2,0,0,0,0,0,0,0,0,0,0,0)
$DESLIGADO= [byte[]](3,0,0,0,0,0,0,0,0,0,0,0)

if (-not (Test-Path $Aprovado)) { New-Item -Path $Aprovado -Force | Out-Null }

$mexeu = 0
foreach ($a in $Alvos) {
  $existeNoRun = $null -ne (Get-ItemProperty -Path $Run -Name $a.nome -ErrorAction SilentlyContinue)
  if (-not $existeNoRun) { Anotar ("pulado (nao esta no Run): " + $a.nome); continue }

  $atual = (Get-ItemProperty -Path $Aprovado -Name $a.nome -ErrorAction SilentlyContinue).($a.nome)
  $querido = if ($Desfazer) { $LIGADO } else { $DESLIGADO }
  # Ausente em StartupApproved significa HABILITADO - nao ha o que desfazer.
  if ($Desfazer -and $null -eq $atual) { continue }
  if ($null -ne $atual -and $atual[0] -eq $querido[0]) { continue }

  try {
    Set-ItemProperty -Path $Aprovado -Name $a.nome -Value $querido -Type Binary -ErrorAction Stop
    if ($Desfazer) { Anotar ("religado na inicializacao: " + $a.nome) }
    else           { Anotar ("fora da inicializacao: " + $a.nome + " (" + $a.porque + ")") }
    $mexeu++
  } catch {
    Anotar ("FALHA em " + $a.nome + ": " + $_.Exception.Message)
  }
}

# ------------------------------------------------------------------- .wslconfig
$wslconfig = Join-Path $env:USERPROFILE '.wslconfig'
$conteudo = @"
# Escrito por servidor\enxugar-inicializacao.ps1 - ver o cabecalho daquele arquivo.
[wsl2]
memory=$MemoriaWsl
swap=2GB

[experimental]
# Sem isto, a memoria que os conteineres liberam fica presa na VM ate o WSL ser
# reiniciado - e o WSL desta maquina so reinicia quando o servidor e desligado.
autoMemoryReclaim=gradual
"@

if ($Desfazer) {
  if (Test-Path $wslconfig) {
    $guarda = Join-Path $Raiz ('servidor\tls\wslconfig-removido-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Copy-Item $wslconfig $guarda -Force
    Remove-Item $wslconfig -Force
    Anotar "removido o .wslconfig (copia guardada em servidor\tls)"
    $mexeu++
  }
} else {
  $precisa = $true
  if (Test-Path $wslconfig) {
    $antes = Get-Content $wslconfig -Raw -ErrorAction SilentlyContinue
    if ($antes -eq $conteudo) { $precisa = $false }
    else {
      $guarda = Join-Path $Raiz ('servidor\tls\wslconfig-anterior-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
      Copy-Item $wslconfig $guarda -Force
      Anotar "havia um .wslconfig diferente; copia guardada em servidor\tls"
    }
  }
  if ($precisa) {
    # UTF8 sem BOM: o WSL le este arquivo, e o BOM na primeira linha faz a secao
    # [wsl2] ser ignorada em silencio - o arquivo parece certo e nao vale nada.
    [IO.File]::WriteAllText($wslconfig, $conteudo, (New-Object Text.UTF8Encoding($false)))
    Anotar ("gravado " + $wslconfig + " (memory=" + $MemoriaWsl + ", autoMemoryReclaim=gradual)")
    $mexeu++
  }
}

# ------------------------------------------------- a parte da maquina (pede UAC)
# POR QUE E SEPARADA: mexer em HKLM e em tarefa agendada exige administrador, e
# a parte de cima nao pode passar a exigir clique nenhum — ela e a que se roda
# de novo depois de qualquer reinstalacao.
#
# POR QUE MOVER O VALOR, e nao usar o StartupApproved como na parte do usuario:
# la o mecanismo e o mesmo do Gerenciador de Tarefas e da para conferir na tela.
# Aqui nao daria: nao ha como saber se o Windows honrou o byte sem deslogar, e
# deslogar neste PC derruba o Docker e a fabrica junto. Mover o valor para
# HKLM:\SOFTWARE\Gerador-OS\Inicializacao-removida e deterministico — o que nao
# esta no Run nao sobe — e o -Desfazer devolve cada valor a chave de onde veio.
$Run64  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$Run32  = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
$Guarda = 'HKLM:\SOFTWARE\Gerador-OS\Inicializacao-removida'

$AlvosMaquina = @(
  @{ chave = $Run64; onde = 'Run64'; nome = 'AdobeAAMUpdater-1.0'; porque = 'so procura atualizacao da Adobe' }
  @{ chave = $Run32; onde = 'Run32'; nome = 'Adobe Creative Cloud'; porque = 'agente do Creative Cloud; e ele que puxa o UI Helper, 130 MB em 5 processos' }
  @{ chave = $Run32; onde = 'Run32'; nome = 'Adobe CCXProcess';     porque = 'outro agente da Adobe' }
  @{ chave = $Run32; onde = 'Run32'; nome = 'SunJavaUpdateSched';   porque = 'so procura atualizacao do Java' }
  @{ chave = $Run32; onde = 'Run32'; nome = 'ControlCenter4';       porque = 'painel de digitalizacao da Brother; a impressora imprime sem ele' }
  @{ chave = $Run32; onde = 'Run32'; nome = 'BrStsMon00';           porque = 'monitor de status da Brother, pelo mesmo motivo' }
)

$TarefasMaquina = @(
  @{ nome = 'StartCN';                   porque = 'painel do Radeon Software (RadeonSoftware + AMDRSSrcExt, ~180 MB)' }
  @{ nome = 'StartDVR';                  porque = 'gravador de tela ReLive da AMD (AMDRSServ, 96 MB) — num servidor' }
  @{ nome = 'Adobe Acrobat Update Task'; porque = 'atualizador do Acrobat' }
)

$souAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($Maquina -and -not $souAdmin) {
  # Subir sozinho, uma vez. O UAC desta maquina esta em "pedir consentimento":
  # aparece uma janela para clicar "Sim", sem senha.
  Anotar 'a parte da maquina precisa de administrador — pedindo o UAC'
  # NAO chamar de $args: e variavel automatica do PowerShell.
  $argumentos = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Maquina')
  if ($Desfazer)    { $argumentos += '-Desfazer' }
  if ($FecharAgora) { $argumentos += '-FecharAgora' }
  try {
    $p = Start-Process powershell.exe -ArgumentList $argumentos -Verb RunAs -PassThru -WindowStyle Hidden
    $p.WaitForExit()
    Anotar ('a passagem com administrador terminou (codigo ' + $p.ExitCode + ')')
  } catch {
    Anotar ('FALHA ao pedir o UAC: ' + $_.Exception.Message + ' — rodar a mao numa janela de administrador')
  }
  exit 0
}

if ($Maquina) {
  if (-not (Test-Path "$Guarda\Run64")) { New-Item -Path "$Guarda\Run64" -Force | Out-Null }
  if (-not (Test-Path "$Guarda\Run32")) { New-Item -Path "$Guarda\Run32" -Force | Out-Null }

  foreach ($a in $AlvosMaquina) {
    $cofre = Join-Path $Guarda $a.onde
    try {
      if ($Desfazer) {
        $valor = (Get-ItemProperty -Path $cofre -Name $a.nome -ErrorAction SilentlyContinue).($a.nome)
        if ($null -eq $valor) { continue }
        Set-ItemProperty -Path $a.chave -Name $a.nome -Value $valor -Type String -ErrorAction Stop
        Remove-ItemProperty -Path $cofre -Name $a.nome -ErrorAction SilentlyContinue
        Anotar ('devolvido a inicializacao: ' + $a.nome)
        $mexeu++
      } else {
        $valor = (Get-ItemProperty -Path $a.chave -Name $a.nome -ErrorAction SilentlyContinue).($a.nome)
        if ($null -eq $valor) { Anotar ('pulado (nao esta no Run da maquina): ' + $a.nome); continue }
        Set-ItemProperty -Path $cofre -Name $a.nome -Value $valor -Type String -ErrorAction Stop
        Remove-ItemProperty -Path $a.chave -Name $a.nome -ErrorAction Stop
        Anotar ('fora da inicializacao da maquina: ' + $a.nome + ' (' + $a.porque + ')')
        $mexeu++
      }
    } catch {
      Anotar ('FALHA em ' + $a.nome + ': ' + $_.Exception.Message)
    }
  }

  foreach ($t in $TarefasMaquina) {
    $tarefa = Get-ScheduledTask -TaskName $t.nome -ErrorAction SilentlyContinue
    if (-not $tarefa) { Anotar ('pulada (nao existe): ' + $t.nome); continue }
    try {
      if ($Desfazer) {
        if ($tarefa.State -ne 'Disabled') { continue }
        Enable-ScheduledTask -TaskName $t.nome -ErrorAction Stop | Out-Null
        Anotar ('tarefa religada: ' + $t.nome)
        $mexeu++
      } else {
        if ($tarefa.State -eq 'Disabled') { continue }
        Disable-ScheduledTask -TaskName $t.nome -ErrorAction Stop | Out-Null
        Anotar ('tarefa desligada: ' + $t.nome + ' (' + $t.porque + ')')
        $mexeu++
      }
    } catch {
      Anotar ('FALHA na tarefa ' + $t.nome + ': ' + $_.Exception.Message)
    }
  }
}

# --------------------------------------------------------------- fechar agora
# Tirar da inicializacao so vale amanha. Se a maquina esta sufocada HOJE, isto
# fecha o que ja esta aberto. Sao agentes, paineis e atualizadores: nao ha
# documento aberto para perder. O Adobe Desktop Service entra na lista porque
# sem ele os "Creative Cloud Helper" ficam orfaos e voltam a subir sozinhos.
#
# MEDIDO EM 19/08/2026: fechou ~650 MB e o livre pulou de 208 para 631 MB. Mas
# o RadeonSoftware e o AMDRSServ VOLTARAM em segundos, relancados por um cmd.exe
# que sai do servico "AMD External Events Utility" — voltaram menores (87 MB no
# lugar de 287), e nao ha o que fazer daqui sem mexer em servico de driver de
# video, o que nao vale o risco. Quem decide o caso deles e o logon de amanha,
# com as tarefas StartCN/StartDVR ja desligadas.
if ($FecharAgora -and -not $Desfazer) {
  $fechar = @(
    'RadeonSoftware', 'AMDRSServ', 'AMDRSSrcExt',
    'Creative Cloud', 'Creative Cloud Helper', 'Creative Cloud UI Helper',
    'CCXProcess', 'Adobe Desktop Service', 'AdobeIPCBroker', 'Adobe Crash Processor',
    'AdobeCollabSync', 'jusched', 'BrStMonW', 'BrCcUxSys', 'M365Copilot'
  )
  $antes = [int]((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB)
  $soltos = 0
  foreach ($nome in $fechar) {
    $ps = Get-Process -Name $nome -ErrorAction SilentlyContinue
    if (-not $ps) { continue }
    $mb = [math]::Round((($ps | Measure-Object WorkingSet64 -Sum).Sum) / 1MB)
    try {
      $ps | Stop-Process -Force -ErrorAction Stop
      $soltos += $mb
      $mexeu++
    } catch {
      Anotar ('nao consegui fechar ' + $nome + ': ' + $_.Exception.Message)
    }
  }
  Start-Sleep -Seconds 3
  $depois = [int]((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB)
  Anotar ("fechados agora: ~$soltos MB em processos; livre passou de $antes MB para $depois MB")
}

if ($mexeu -eq 0) { Anotar 'nada a fazer: ja estava como se queria' }
else { Anotar 'pronto. As duas mudancas valem no PROXIMO logon; nada foi interrompido agora.' }
exit 0
