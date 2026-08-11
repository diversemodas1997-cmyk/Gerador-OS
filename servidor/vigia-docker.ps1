<#
  Vigia do servidor da fabrica: garante que o Docker e os conteineres estejam de pe.

  POR QUE EXISTE:
  em 11/08/2026 a fabrica ficou sem o programa a manha inteira porque o PC foi
  desligado no fim do expediente anterior e, ao religar, o Docker Desktop nao
  subiu. O "restart: always" dos conteineres NAO salva nesse caso — ele so vale
  depois que o motor do Docker esta rodando; sem motor, nada sobe.

  E ha uma segunda armadilha, que e o motivo de este script mandar "up -d" em vez
  de so abrir o Docker: conteiner parado DE PROPOSITO (o que o desligar-servidor.ps1
  faz, para o Postgres nao ser arrancado no meio de uma escrita) fica marcado como
  parado, e o "restart: always" respeita isso e NAO o levanta. Quem levanta e aqui.

  So grava no log quando FAZ alguma coisa ou quando falha. Rodando de 5 em 5
  minutos, um log tagarela viraria 288 linhas por dia e ninguem leria a linha que
  importa.

  Registrar (uma vez, no servidor):
    .\servidor\vigia-docker.ps1 -Agendar

  Conferir na mao, sem esperar o agendador:
    .\servidor\vigia-docker.ps1
#>
[CmdletBinding()]
param(
  [string] $Docker       = 'C:\supabase\docker',
  [int]    $EsperaMin    = 5,
  [int]    $CadaMinutos  = 5,
  [switch] $Agendar
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\vigia-docker.log'   # tls/ e ignorado pelo git

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  try { Add-Content -Path $Log -Value $linha -Encoding utf8 } catch { }
  Write-Host $linha
}

# O agendador nao herda o PATH da sessao interativa em toda situacao; procurar o
# docker.exe pelo caminho conhecido evita um "nao achei o comando" travestido de
# "o Docker esta fora do ar", que mandaria o vigia reabrir o Docker a cada 5 min.
function Achar-Docker {
  $c = Get-Command docker.exe -ErrorAction SilentlyContinue
  if ($c) { return $c.Source }
  $p = 'C:\Program Files\Docker\Docker\resources\bin\docker.exe'
  if (Test-Path $p) { return $p }
  return $null
}

function Motor-Responde($exe) {
  if (-not $exe) { return $false }
  & $exe info --format '{{.ServerVersion}}' 2>$null | Out-Null
  return ($LASTEXITCODE -eq 0)
}

# ------------------------------------------------------------------- agendar
if ($Agendar) {
  $nome = 'Gerador-OS Vigia Docker'
  $ps1  = Join-Path $PSScriptRoot 'vigia-docker.ps1'

  # Janela preta piscando de 5 em 5 minutos na cara de quem usa o PC faria
  # qualquer um desativar a tarefa em uma semana. O wscript com modo 0 abre o
  # PowerShell sem janela nenhuma — nem o flash de um quadro.
  $vbs = Join-Path $Raiz 'servidor\tls\vigia-oculto.vbs'
  $cmd = 'powershell -NoProfile -ExecutionPolicy Bypass -File "' + $ps1 + '" -Docker "' + $Docker + '"'
  $conteudo = 'CreateObject("Wscript.Shell").Run "' + ($cmd -replace '"', '""') + '", 0, False'
  Set-Content -Path $vbs -Value $conteudo -Encoding ASCII

  $acao = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $vbs + '"') -WorkingDirectory $Raiz

  # Dois gatilhos: um no logon (o caso de religar o PC) e um repetindo para
  # sempre (o caso de o Docker morrer no meio do expediente). O atraso de 1 min
  # no logon e para nao brigar com o Windows enquanto ele ainda esta montando a
  # sessao — inclusive o J: do Google Drive.
  $noLogon = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
  $noLogon.Delay = 'PT1M'

  $repetindo = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(1) `
                 -RepetitionInterval (New-TimeSpan -Minutes $CadaMinutos)

  # "Somente com o usuario conectado", de proposito e pelo mesmo motivo do
  # backup-diario.ps1: o Docker Desktop no Windows so existe DENTRO da sessao.
  # Fora dela nao ha o que abrir.
  $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

  try {
    Unregister-ScheduledTask -TaskName $nome -Confirm:$false -ErrorAction SilentlyContinue
    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger @($noLogon, $repetindo) `
      -Settings $conf `
      -Description 'Abre o Docker Desktop e levanta os conteineres do Gerador-OS se estiverem fora.' `
      -ErrorAction Stop | Out-Null
    Anotar "tarefa '$nome' registrada: no logon (+1 min) e a cada $CadaMinutos min"
  } catch {
    Anotar "FALHA ao registrar a tarefa: $($_.Exception.Message)"
    exit 1
  }
  exit 0
}

# -------------------------------------------------------------------- rodar
#
# CUIDADO ao renomear: esta variavel NAO pode se chamar $docker. O parametro
# $Docker acima guarda a PASTA do Supabase, e no PowerShell $docker e $Docker sao
# a MESMA variavel. Pior: como o parametro e [string], atribuir o objeto do
# Get-Command nele nao da erro — o PowerShell converte para o texto "docker.exe"
# calado, .Source vira vazio e a pasta do Supabase se perde. Foi assim que a
# primeira versao deste script quebrou, e so no ramo de recuperacao, que e o
# unico que interessa.
$dockerExe = Achar-Docker
if (-not $dockerExe) { Anotar 'FALHA: nao achei o docker.exe — o Docker Desktop foi desinstalado?'; exit 1 }

$agi = $false   # so vira log se o vigia tiver mesmo feito algo

# 1) O motor esta de pe?
if (-not (Motor-Responde $dockerExe)) {
  $agi = $true
  Anotar 'motor do Docker fora do ar — abrindo o Docker Desktop'

  $appExe = 'C:\Program Files\Docker\Docker\Docker Desktop.exe'
  if (-not (Test-Path $appExe)) { Anotar "FALHA: nao achei $appExe"; exit 1 }
  try { Start-Process $appExe -ErrorAction Stop } catch { Anotar "FALHA ao abrir: $($_.Exception.Message)"; exit 1 }

  $limite = (Get-Date).AddMinutes($EsperaMin)
  while ((Get-Date) -lt $limite -and -not (Motor-Responde $dockerExe)) { Start-Sleep -Seconds 10 }

  if (-not (Motor-Responde $dockerExe)) {
    Anotar "FALHA: o motor nao respondeu em $EsperaMin min. A proxima passagem tenta de novo."
    exit 1
  }
  Anotar 'motor do Docker respondendo'
}

# 2) Os conteineres que a fabrica precisa estao rodando?
$precisa = @('gerador-os-web', 'supabase-db', 'supabase-kong', 'supabase-rest', 'supabase-auth')
$rodando = @(& $dockerExe ps --format '{{.Names}}' 2>$null)
$faltando = @($precisa | Where-Object { $rodando -notcontains $_ })

if ($faltando.Count -gt 0) {
  $agi = $true
  Anotar ('conteineres fora: ' + ($faltando -join ', ') + ' — levantando')

  # O Supabase primeiro: o nginx do app so tem o que servir depois que a API existe.
  $pilhas = @(
    @{ nome = 'supabase'; arquivo = (Join-Path $Docker 'docker-compose.yml') },
    @{ nome = 'app';      arquivo = (Join-Path $PSScriptRoot 'docker-compose.app.yml') }
  )
  foreach ($p in $pilhas) {
    if (-not (Test-Path $p.arquivo)) { Anotar ("FALHA: nao achei " + $p.arquivo); exit 1 }
    $saida = & $dockerExe compose -f $p.arquivo up -d 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
      $motivo = ($saida.Trim() -replace '\s+', ' ')
      if ($motivo.Length -gt 300) { $motivo = $motivo.Substring(0, 300) + '…' }
      Anotar ("FALHA ao levantar a pilha '" + $p.nome + "': " + $motivo)
      exit 1
    }
  }

  # O Postgres pode levar minutos recuperando o WAL depois de um desligamento
  # brusco. Dizer "ok" antes disso seria mentira: o site ja responde, mas as
  # telas voltam vazias, que e exatamente o susto que se quer evitar.
  $limite = (Get-Date).AddMinutes($EsperaMin)
  $banco = $false
  while ((Get-Date) -lt $limite) {
    $r = & $dockerExe exec supabase-db pg_isready -U postgres 2>&1 | Out-String
    if ($r -match 'accepting connections') { $banco = $true; break }
    Start-Sleep -Seconds 10
  }
  if ($banco) { Anotar 'tudo de pe: banco aceitando conexoes' }
  else { Anotar "ATENCAO: conteineres levantados, mas o banco nao respondeu em $EsperaMin min"; exit 1 }
}

if (-not $agi) { exit 0 }   # silencio: estava tudo certo, nao ha o que contar
exit 0
