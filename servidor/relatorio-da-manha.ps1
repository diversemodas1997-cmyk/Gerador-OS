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
