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

  O QUE NAO E MEXIDO, DE PROPOSITO:
    GoogleDriveFS  - os dois backups diarios gravam em J:\Meu Drive. Sem ele,
                     a fabrica fica sem backup, que e pior que ficar lenta.
    Audaces 360    - e software de trabalho, com licenca; nao e enfeite.
    Docker Desktop - ja esta desligado, e assim tem de ficar: quem abre o Docker
                     e o vigia-docker.ps1, e dois donos da mesma decisao brigam.
    Defender/Topaz - seguranca. Lentidao nao se conserta desligando protecao.

  Rodar (sem administrador, na conta do servidor):
    .\servidor\enxugar-inicializacao.ps1
  Voltar tudo como estava:
    .\servidor\enxugar-inicializacao.ps1 -Desfazer
#>
[CmdletBinding()]
param(
  [switch] $Desfazer,
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

if ($mexeu -eq 0) { Anotar 'nada a fazer: ja estava como se queria' }
else { Anotar 'pronto. As duas mudancas valem no PROXIMO logon; nada foi interrompido agora.' }
exit 0
