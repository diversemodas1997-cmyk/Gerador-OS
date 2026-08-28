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

  E ha uma terceira, descoberta em 12/08/2026: o Docker Desktop tambem nao sobe
  quando os JSONs de configuracao dele estao zerados por um desligamento sujo.
  Nesse caso abrir o app de novo a cada 5 minutos nao adianta nada — por isso o
  vigia agora confere e repoe esses arquivos antes de abrir o Docker.

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
  # Quanto esperar o MOTOR do Docker. Eram 5 min, e em 13/08/2026 ele levou
  # 5m18s: o vigia escapou de desistir por 33 segundos. A máquina tem 8 GB e
  # arranca com a memória no talo, então esse tempo varia de manhã para manhã.
  # Desistir custa caro (mais 5 min até a próxima passagem, com a fábrica
  # parada) e esperar não custa nada: o loop sai no segundo em que o motor
  # responde. Separado do $EsperaMin de propósito — o banco tem outro ritmo, e
  # somar os dois tem de caber no limite de 20 min da tarefa agendada.
  [int]    $EsperaMotorMin = 10,
  [int]    $EsperaMin    = 5,
  [int]    $CadaMinutos  = 5,
  [switch] $Agendar
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\vigia-docker.log'   # tls/ e ignorado pelo git

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  # Tentar de novo em vez de engolir a falha: em 12/08/2026 varias passagens
  # escreveram no log ao mesmo tempo, o arquivo estava travado, o catch vazio
  # comeu o erro e a linha que explicava a pane simplesmente nao apareceu.
  for ($i = 0; $i -lt 5; $i++) {
    try { Add-Content -Path $Log -Value $linha -Encoding utf8; break }
    catch { Start-Sleep -Milliseconds 200 }
  }
  Write-Host $linha
}

# Um desligamento sujo (queda de luz, botao de forca) pode deixar estes JSONs do
# Docker cheios de bytes zero: o NTFS salvou o tamanho do arquivo, mas nao o
# conteudo. O Docker Desktop entao abre, o motor NAO sobe e a janela nao diz o
# porque — foi o que segurou a fabrica na manha de 12/08/2026. Conferir aqui,
# antes de abrir o Docker, custa milissegundos e evita a manha inteira parada.
function Consertar-Configs {
  $padroes = @{
    (Join-Path $env:USERPROFILE '.docker\daemon.json')         = "{`r`n  `"builder`": {`r`n    `"gc`": {`r`n      `"defaultKeepStorage`": `"20GB`",`r`n      `"enabled`": true`r`n    }`r`n  },`r`n  `"experimental`": false`r`n}"
    (Join-Path $env:USERPROFILE '.docker\windows-daemon.json') = "{`r`n  `"experimental`": false`r`n}"
    (Join-Path $env:APPDATA 'Docker\features-overrides.json')  = '{}'
    (Join-Path $env:APPDATA 'Docker\settings-store.json')      = $null   # o Docker recria
    (Join-Path $env:APPDATA 'Docker\login-info.json')          = $null   # so cache de login
  }

  $consertou = $false
  foreach ($arquivo in $padroes.Keys) {
    if (-not (Test-Path $arquivo)) { continue }
    $estragado = $false
    try {
      $bytes = [IO.File]::ReadAllBytes($arquivo)
      if ($bytes -contains 0) { $estragado = $true }
      elseif ($bytes.Length -gt 0) {
        try { [Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json | Out-Null }
        catch { $estragado = $true }
      }
    } catch { continue }
    if (-not $estragado) { continue }

    $nome = Split-Path $arquivo -Leaf
    $guarda = Join-Path $Raiz 'servidor\tls\configs-corrompidas'
    New-Item -ItemType Directory -Force -Path $guarda | Out-Null
    try {
      Copy-Item $arquivo (Join-Path $guarda ($nome + '.' + (Get-Date -Format 'yyyyMMdd-HHmmss'))) -Force
      if ($null -eq $padroes[$arquivo]) { Remove-Item $arquivo -Force }
      else { [IO.File]::WriteAllText($arquivo, $padroes[$arquivo]) }
      Anotar "config do Docker corrompida (zerada): $nome — reposta pelo padrao"
      $consertou = $true
    } catch {
      Anotar "FALHA ao repor $nome : $($_.Exception.Message)"
    }
  }
  return $consertou
}

# Quando o motor nao volta, dizer so "nao respondeu" nao ajuda ninguem. O Docker
# escreve o motivo real no log dele; trazer essa linha para ca e a diferenca
# entre reinstalar no escuro e saber qual arquivo consertar.
function Motivo-Do-Docker {
  $log = Join-Path $env:LOCALAPPDATA 'Docker\log\host\com.docker.backend.exe.log'
  if (-not (Test-Path $log)) { return $null }
  try {
    $linha = Get-Content $log -Tail 400 -ErrorAction Stop |
             Where-Object { $_ -match 'backend crashed|fatal|panic:' } |
             Select-Object -Last 1
    if (-not $linha) { return $null }
    $limpa = ($linha -replace '\s+', ' ').Trim()
    if ($limpa.Length -gt 300) { $limpa = $limpa.Substring(0, 300) + '…' }
    return $limpa
  } catch { return $null }
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

# NAO trocar por 'docker info' (28/08/2026). Parece a pergunta certa, e era o que
# estava aqui, mas o 'info' monta o retrato do CLIENTE tambem: ele EXECUTA os 15
# cli-plugins instalados, um por um, so para ler a versao de cada um. Medido nesta
# maquina com o motor saudavel: 'info' 455 ms contra 72 ms deste 'version'.
#
# E o primeiro plugin da fila e o docker-agent.exe — justamente o que fica ate
# 60 s pendurado no arranque (ver Desligar-Agente-IA). Ou seja: a pergunta "o
# motor ja subiu?" ficava presa no unico programa que ainda nao tinha subido. Em
# 28/08 o motor estava de pe as 07:16:18 e o vigia so conseguiu ver as 07:18:13 —
# 115 s anotados como demora do motor que eram demora da PERGUNTA.
#
# 'version' pergunta a versao AO SERVIDOR: se ele responde, o motor esta de pe.
function Motor-Responde($exe) {
  if (-not $exe) { return $false }
  & $exe version --format '{{.Server.Version}}' 2>$null | Out-Null
  return ($LASTEXITCODE -eq 0)
}

# O "Docker AI Agent" nao serve para nada nesta maquina — a fabrica roda nginx e
# Postgres locais — e cobra caro em toda manha. O log do Docker de 28/08:
#
#   10:16:20  starting docker agent api process
#   10:16:20  pinging agent api server at \\.\pipe\dockerAgent
#   ... 18 tentativas, todas falhando ...
#   10:17:21  agent api process stopped as requested        (60,6 s)
#   10:17:28  boot probe to https://registry-1.docker.io/v2/ failed: context deadline
#   10:17:30  docker agent boot timed out
#
# Setenta segundos por manha para uma funcao que falha sempre — e que ainda
# segurava o 'docker info' de quem estivesse perguntando pelo motor.
#
# Aqui e o lugar certo de desligar: este trecho so roda com o motor fora do ar e
# ANTES de abrir o Docker Desktop, entao a gravacao nao e sobrescrita por ele ao
# fechar. E se uma atualizacao do Docker religar a funcao, a manha seguinte
# desliga de novo sozinha.
function Desligar-Agente-IA {
  $arq = Join-Path $env:APPDATA 'Docker\settings-store.json'
  if (-not (Test-Path $arq)) { return $false }
  try {
    $cfg = Get-Content $arq -Raw -Encoding UTF8 | ConvertFrom-Json
  } catch { return $false }   # corrompido e problema do Consertar-Configs, nao deste
  if ($null -ne $cfg.EnableDockerAI -and -not $cfg.EnableDockerAI) { return $false }
  try {
    if ($null -eq $cfg.PSObject.Properties['EnableDockerAI']) {
      $cfg | Add-Member -NotePropertyName 'EnableDockerAI' -NotePropertyValue $false
    } else { $cfg.EnableDockerAI = $false }
    [IO.File]::WriteAllText($arq, ($cfg | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
    Anotar 'agente de IA do Docker desligado (economiza ~70 s de arranque)'
    return $true
  } catch {
    Anotar "aviso: nao consegui desligar o agente de IA do Docker: $($_.Exception.Message)"
    return $false
  }
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
  # sempre (o caso de o Docker morrer no meio do expediente).
  #
  # SEM ATRASO NO LOGON (28/08/2026). Havia um 'Delay = PT1M' aqui, para "nao
  # brigar com o Windows enquanto ele ainda monta a sessao — inclusive o J: do
  # Google Drive". Mas o vigia nao le o J: nem nenhuma outra pasta de rede: ele
  # confere cinco JSONs em C:, abre o Docker Desktop e chama o compose. O
  # minuto protegia contra um trabalho que ele nao faz.
  #
  # E custava caro na unica hora que importa. Em 28/08 a maquina ligou as
  # 07:13:52 e o vigia so abriu o Docker as 07:15:47 — quase dois minutos com o
  # servidor ligado e ninguem pedindo o motor, que e a parte mais demorada do
  # arranque (144 s naquela manha). Abrir o Docker JUNTO com o resto do logon
  # sobrepoe as duas esperas em vez de somar.
  $noLogon = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

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
    Anotar "tarefa '$nome' registrada: no logon (sem atraso) e a cada $CadaMinutos min"
  } catch {
    Anotar "FALHA ao registrar a tarefa: $($_.Exception.Message)"
    exit 1
  }
  exit 0
}

# -------------------------------------------------------------------- rodar
#
# UMA passagem de cada vez. O agendador tem a opcao "IgnoreNew", mas ela nao vale
# aqui: a tarefa chama o wscript, que dispara o PowerShell e morre no mesmo
# instante. Para o Windows a tarefa ja acabou, e a cada 5 min ele solta mais uma.
# Com o motor fora do ar, cada passagem fica ate 5 min esperando — e as passagens
# se empilham. Em 12/08/2026 tres delas mandaram "compose up" no mesmo projeto ao
# mesmo tempo, uma travou a outra e nenhum conteiner subiu por 25 minutos. Esta
# tranca e o que garante que so uma trabalhe; as outras saem caladas.
$tranca = New-Object System.Threading.Mutex($false, 'Local\GeradorOS-VigiaDocker')
$minhaVez = $false
try { $minhaVez = $tranca.WaitOne(0) }
catch [System.Threading.AbandonedMutexException] { $minhaVez = $true }   # a anterior morreu segurando
if (-not $minhaVez) { exit 0 }

try {
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

    if (Consertar-Configs) { $agi = $true }   # antes de abrir: config zerada nao deixa o motor subir
    Desligar-Agente-IA | Out-Null              # tambem antes de abrir, ou o Docker regrava ao fechar

    $appExe = 'C:\Program Files\Docker\Docker\Docker Desktop.exe'
    if (-not (Test-Path $appExe)) { Anotar "FALHA: nao achei $appExe"; exit 1 }
    try { Start-Process $appExe -ErrorAction Stop } catch { Anotar "FALHA ao abrir: $($_.Exception.Message)"; exit 1 }

    $comecou = Get-Date
    $limite = $comecou.AddMinutes($EsperaMotorMin)
    while ((Get-Date) -lt $limite -and -not (Motor-Responde $dockerExe)) { Start-Sleep -Seconds 10 }

    if (-not (Motor-Responde $dockerExe)) {
      $porque = Motivo-Do-Docker
      if ($porque) { Anotar "FALHA: o motor nao respondeu em $EsperaMotorMin min. O Docker reclamou: $porque" }
      else         { Anotar "FALHA: o motor nao respondeu em $EsperaMotorMin min. A proxima passagem tenta de novo." }
      exit 1
    }
    # Anotar quanto demorou: e o unico numero que diz se a maquina esta ficando
    # mais lenta a cada manha. 40 s e o normal; 5 min ja e memoria no limite.
    # Junto com o tempo vai a memoria livre: sao os dois numeros que andam
    # juntos. Em 13/08/2026 o motor levou 5m18s com 99 MB livres, e sem o
    # segundo numero no log gastei a manha procurando defeito no Docker. Se um
    # dia o arranque voltar a demorar, o log ja diz se foi falta de ar.
    # (atribuir direto de um try/catch so vale do PowerShell 7 em diante; aqui e 5.1)
    $livre = -1
    try { $livre = [int]((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB) } catch { }
    Anotar ('motor do Docker respondendo (levou ' + [int]((Get-Date) - $comecou).TotalSeconds + ' s, ' + $livre + ' MB livres)')
  }

  # 2) Os conteineres que a fabrica precisa estao rodando?
  #
  # O 'supabase-storage' entrou em 19/08/2026, e custou uma manha de desenhos
  # invisiveis: na noite anterior ele parou com calma junto com o resto, o vigia
  # subiu os outros e, como ele NAO estava nesta lista, o log do vigia disse
  # "tudo de pe" com o storage fora. A OS abria, a lista de OS abria, so as
  # imagens dos desenhos vinham quebradas — e ninguem liga isso a um conteiner.
  # A regra que ficou: entra nesta lista tudo que, faltando, o usuario percebe.
  $precisa = @('gerador-os-web', 'supabase-db', 'supabase-kong', 'supabase-rest',
               'supabase-auth', 'supabase-storage', 'supabase-edge-functions')
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
        # Cortar os avisos ANTES de truncar em 300 caracteres. O compose abre a
        # saida com o aviso do volume "supabase_deno-cache" (existe, mas nao foi
        # ele quem criou) toda vez, e so ele ja passa dos 300 — de 14/08 a
        # 18/08/2026 TODA falha desta linha foi registrada como se fosse esse
        # aviso, e o erro de verdade nunca chegou ao log. O aviso e inofensivo;
        # o que interessa vem depois dele.
        $util = ($saida -split "`n" | Where-Object { $_ -notmatch 'level=warning' }) -join ' '
        if (-not $util.Trim()) { $util = $saida }
        $motivo = ($util.Trim() -replace '\s+', ' ')
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
}
finally {
  try { $tranca.ReleaseMutex() } catch { }
  $tranca.Dispose()
}
