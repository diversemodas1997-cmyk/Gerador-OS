<#
  Conta, numa janelinha, como foi o arranque do servidor naquela manha.

  POR QUE EXISTE:
  o vigia ja anota tudo o que interessa - quanto o motor do Docker demorou,
  quanta memoria havia livre, se alguma pilha falhou. So que ninguem abre
  servidor\tls\vigia-docker.log de manha, e o log so e lido DEPOIS que a fabrica
  ja passou meia hora parada. Este script inverte isso: as 08:05 a maquina
  mostra o resumo sem ninguem pedir.

  Os dois numeros que importam andam juntos: o TEMPO do motor e a MEMORIA LIVRE.
  Em 13/08/2026 o motor levou 5m18s com 99 MB livres; em 19/08, 432 s com
  211 MB. Depois do enxugamento de 19/08 (ver enxugar-inicializacao.ps1) o
  esperado e mais ar e menos tempo - e e esta janelinha que diz se pegou.

  A SEGUNDA PERGUNTA DA MANHA, desde 18/09/2026: o Audaces tem licenca?
  A licenca do Audaces venceu em 04/09/2026 e a fabrica so descobriu no dia 17,
  quando alguem foi abrir um encaixe. Treze dias de silencio - porque o unico
  aviso de vencimento que o programa da e um popup que ele mostrou UMA vez, em
  08/08, numa maquina que trabalha sozinha com autologon e ninguem olha.
  E quando falta licenca, nenhuma tela diz "licenca": o Encaixe diz "Erro fatal:
  Problemas de comunicacao com servidor: localhost", que manda todo mundo cacar
  rede, porta e firewall. Custou uma manha inteira.
  Aqui a pergunta e feita todo dia, e em portugues: o servidor de licenca do
  Audaces Go responde em 127.0.0.1:5556? Enquanto responder, esta e uma linha
  mansa no meio do relatorio; no dia em que parar, e um ATENCAO com a data da
  ultima licenca servida, que e o que diz ha quanto tempo a coisa esta parada.

  A janela se fecha sozinha em 2 minutos. Um aviso que fica esperando clique
  num servidor e um aviso que trava o servidor.

  Registrar (uma vez):
    .\servidor\relatorio-da-manha.ps1 -Agendar
  Ver agora, sem esperar:
    .\servidor\relatorio-da-manha.ps1
  Tirar:
    Unregister-ScheduledTask -TaskName 'Gerador-OS Relatorio da Manha' -Confirm:$false
#>
[CmdletBinding()]
param(
  [string] $Hora = '08:05',
  [switch] $Agendar,
  # Sem janela: so escreve no log. E o que a tarefa usa quando ninguem esta
  # olhando a tela, e o que serve para conferir o historico depois.
  [switch] $Calado
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\relatorio-da-manha.log'
$LogVigia = Join-Path $Raiz 'servidor\tls\vigia-docker.log'

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  for ($i = 0; $i -lt 5; $i++) {
    try { Add-Content -Path $Log -Value $linha -Encoding utf8; break }
    catch { Start-Sleep -Milliseconds 200 }
  }
  Write-Host $linha
}

# ------------------------------------------------------- a licenca do Audaces
# O visor da bandeja (pserver.exe) fala com o servidor de licenca por HTTP em
# 127.0.0.1:5556 - e a mesma porta que o proprio programa usa, a cada 30 s.
# Se ela responde, ha licenca sendo servida. Se recusa conexao, nao ha, e dai em
# diante nenhuma tela do Audaces diz a palavra "licenca".
# Devolve as linhas do relatorio, ou nada quando nao ha Audaces nesta maquina.
function ConferirLicencaDoAudaces {
  $pasta = Join-Path $env:LOCALAPPDATA 'Audaces\pserver'
  if (-not (Test-Path $pasta)) { return @() }

  $responde = $false
  try {
    $cliente   = New-Object Net.Sockets.TcpClient
    $tentativa = $cliente.ConnectAsync('127.0.0.1', 5556)
    $responde  = ($tentativa.Wait(1500) -and $cliente.Connected)
    $cliente.Close()
  } catch { $responde = $false }

  # ---------------------------------------------------------- nao responde
  if (-not $responde) {
    # Quando parou esta no go.db, o banco do proprio servidor: uma linha por
    # sessao de licenca, com inicio e fim. Ler as datas cruas de um arquivo de
    # 8 KB basta - abrir SQLite aqui exigiria carregar driver so para isto.
    $ultima = ''
    $go = Join-Path $pasta 'go.db'
    if (Test-Path $go) {
      try {
        $cru   = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($go))
        $datas = @([regex]::Matches($cru, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}') | ForEach-Object { $_.Value })
        if ($datas.Count -gt 0) { $ultima = ($datas | Sort-Object)[-1] }
      } catch { }
    }
    $recado = 'ATENCAO - o Audaces esta SEM LICENCA: o servidor de licenca nao responde.'
    if ($ultima) {
      $quando = [datetime]::ParseExact($ultima, 'yyyy-MM-dd HH:mm:ss', $null)
      $parado = [int]((Get-Date) - $quando).TotalDays
      $recado += ' A ultima licenca servida foi em ' + $quando.ToString('dd/MM/yyyy') + ", ha $parado dia(s)."
    }
    return @(
      $recado
      '  O Encaixe vai dizer "Erro fatal: Problemas de comunicacao com servidor:'
      '  localhost". Nao e rede: e licenca. Renovar com a Audaces.'
    )
  }

  # ------------------------------------------------------------- responde
  # Com licenca viva o /info traz a data de vencimento. O formato pode mudar de
  # versao para versao, entao aqui e melhor nao achar a data do que errar o dia:
  # sem data legivel, a linha apenas diz que a licenca esta no ar.
  $vence = $null
  try {
    $info = Invoke-WebRequest -Uri 'http://127.0.0.1:5556/info' -TimeoutSec 4 -UseBasicParsing
    $achou = [regex]::Match([string]$info.Content, '(?i)"[^"]*expir[^"]*"\s*:\s*"([^"]{8,40})"')
    if ($achou.Success) {
      # -as devolve nulo quando nao entende, em vez de estourar. TryParse aqui
      # nao serve: o [ref] de uma variavel vazia e erro em PowerShell, e o erro
      # cairia no catch fingindo que a data nao existe.
      $vence = $achou.Groups[1].Value -as [datetime]
    }
  } catch { }

  if (-not $vence) { return @('Audaces: licenca no ar.') }

  $faltam = [int]($vence - (Get-Date)).TotalDays
  # Trinta dias e o prazo para renovar sem parar a fabrica; dai para baixo o
  # aviso sobe de tom todo dia, porque um aviso dado uma vez so ja falhou aqui.
  if ($faltam -gt 30) {
    return @('Audaces: licenca no ar, vence em ' + $vence.ToString('dd/MM/yyyy') + ".")
  }
  return @(
    "ATENCAO - a licenca do Audaces vence em $faltam dia(s), em " + $vence.ToString('dd/MM/yyyy') + '.'
    '  Renovar antes: vencida, ela derruba o Encaixe e a mensagem fala em rede,'
    '  nao em licenca.'
  )
}

# ------------------------------------------------------------------- agendar
if ($Agendar) {
  # Pelo wscript, como o vigia: a tarefa chamando powershell.exe direto pisca
  # uma janela preta na cara de quem estiver trabalhando.
  $vbs = Join-Path $Raiz 'servidor\tls\relatorio-da-manha.vbs'
  $ps1 = Join-Path $PSScriptRoot 'relatorio-da-manha.ps1'
  $conteudoVbs = 'CreateObject("Wscript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""' + $ps1 + '""", 0, False'
  [IO.File]::WriteAllText($vbs, $conteudoVbs, (New-Object Text.ASCIIEncoding))

  $acao    = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $vbs + '"')
  $gatilho = New-ScheduledTaskTrigger -Daily -At $Hora
  # Tres minutos de folga: se o PC ligou tarde, o relatorio ainda sai.
  $opcoes  = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
  Register-ScheduledTask -TaskName 'Gerador-OS Relatorio da Manha' -Action $acao -Trigger $gatilho `
    -Settings $opcoes -Description 'Mostra como foi o arranque do servidor naquela manha.' -Force | Out-Null
  Anotar ("tarefa registrada para as $Hora, todo dia")
  exit 0
}

# -------------------------------------------------------------------- montar
$hoje    = Get-Date -Format 'yyyy-MM-dd'
$linhas  = @()
if (Test-Path $LogVigia) {
  $linhas = @(Get-Content $LogVigia -ErrorAction SilentlyContinue | Where-Object { $_ -like "$hoje*" })
}

$partes = @()

$motor = @($linhas | Where-Object { $_ -match 'motor do Docker respondendo' })[-1]
if ($motor) {
  # A linha nova traz "(levou 432 s, 211 MB livres)"; a antiga so os segundos.
  if ($motor -match 'levou (\d+) s(?:, (-?\d+) MB livres)?') {
    $seg = [int]$matches[1]
    $mem = $matches[2]
    $comoFoi = if ($seg -le 90) { 'normal' } elseif ($seg -le 240) { 'devagar' } else { 'MUITO devagar' }
    $texto = "O motor do Docker levou $seg s ($comoFoi)."
    if ($mem) { $texto += " Havia $mem MB de memoria livre nessa hora." }
    $partes += $texto
  } else {
    $partes += ($motor -replace '^\S+\s+\S+\s+', '')
  }
} elseif ($linhas.Count -gt 0) {
  $partes += 'O motor ja estava de pe: o vigia nao precisou abrir o Docker.'
} else {
  $partes += 'O vigia nao escreveu nada hoje - sinal de que nao houve o que fazer.'
}

$subida = @($linhas | Where-Object { $_ -match 'tudo de pe' })[-1]
if ($subida) { $partes += 'Os conteineres subiram e o banco respondeu.' }

$ruim = @($linhas | Where-Object { $_ -match 'FALHA|ATENCAO' })
if ($ruim.Count -gt 0) {
  $partes += ''
  $partes += "ATENCAO - o vigia registrou $($ruim.Count) linha(s) de problema hoje:"
  foreach ($r in $ruim | Select-Object -Last 3) {
    $curta = $r -replace '^\S+\s+\S+\s+', ''
    if ($curta.Length -gt 200) { $curta = $curta.Substring(0, 200) + '...' }
    $partes += "  - $curta"
  }
}

# A licenca do Audaces entra antes do "Agora": e a unica linha do relatorio que
# nao fala do Gerador-OS, e e a que ninguem ia procurar sozinho.
$licenca = @(ConferirLicencaDoAudaces)
if ($licenca.Count -gt 0) {
  $partes += ''
  $partes += $licenca
}

# O estado de agora, que e o que a pessoa vai querer saber em seguida.
$deP = 0
try { $deP = @(& docker ps --format '{{.Names}}' 2>$null).Count } catch { }
$livre = -1
try { $livre = [int]((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB) } catch { }
$partes += ''
$partes += "Agora: $deP conteineres de pe, $livre MB de memoria livre."

$resumo = ($partes -join "`r`n")
Anotar ($resumo -replace "`r`n", ' | ')

if (-not $Calado) {
  # Popup do WScript: fecha sozinho no tempo dado. Um MessageBox comum ficaria
  # esperando clique para sempre, e numa tarefa agendada isso e uma tarefa
  # travada ate alguem chegar na maquina.
  try {
    (New-Object -ComObject WScript.Shell).Popup($resumo, 120, 'Servidor Gerador-OS - a manha de hoje', 64) | Out-Null
  } catch {
    Anotar ('nao consegui mostrar a janela: ' + $_.Exception.Message)
  }
}
exit 0
