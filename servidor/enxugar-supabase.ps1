<#
  Tira da pilha do Supabase o que a fabrica nao usa, e desafoga o gateway.

  POR QUE EXISTE:
  o enxugamento de 19/08/2026 (enxugar-inicializacao.ps1) atacou o WINDOWS -
  Adobe, Radeon, monitores de impressora - e deu resultado: o motor do Docker
  caiu de 1220 s (18/08) para 106 s (27/08). So que o aviso da manha de 27/08
  mostrou a outra metade do problema: havia 910 MB livres na hora do arranque e,
  com os 12 conteineres de pe, sobraram 254 MB. O peso agora esta DENTRO da
  pilha, e nao mais na inicializacao do Windows.

  O QUE FOI MEDIDO (27/08/2026, docker stats, maquina de 8 GB / WSL com teto 4):
    kong      592 MB   <- o maior, e e nginx com 8 workers num PC de 8 nucleos
    studio    264 MB   <- o PAINEL de administracao do Supabase
    storage   246 MB
    realtime  246 MB
    pooler    232 MB   <- supavisor, conexao Postgres direta
    db        231 MB
    meta      131 MB   <- API que o painel usa; so ele usa
    imgproxy  125 MB   <- corte/redimensionamento de imagem do Storage
    functions 105 MB
    rest       72 MB
    auth       31 MB
    web        17 MB

  O QUE ESTE SCRIPT TIRA, E POR QUE CADA UM PODE SAIR:
    studio + meta (395 MB) - o painel do Supabase. A fabrica nao abre painel de
      banco; quem precisa dele e a manutencao, de vez em quando. Fica SOB
      DEMANDA: -Painel sobe os dois, -FecharPainel derruba.
    pooler / supavisor (232 MB) - poll de conexoes Postgres na 5432/6543.
      O programa fala com o banco por PostgREST (kong -> rest), nunca por
      conexao direta; nenhum outro conteiner depende dele.
    imgproxy (125 MB) - so serve para o Storage entregar imagem redimensionada
      (/render/image/...). O programa pede o desenho pelo endereco publico
      direto (/object/public/desenhos/...); nao ha uma chamada de transformacao
      no codigo inteiro.

  E O QUE ELE CONSERTA DE GRACA, QUE VALE TANTO QUANTO A MEMORIA:
    · O KONG ESPERAVA O STUDIO. No compose de fabrica o gateway tem
      `depends_on: studio: condition: service_healthy` — ou seja, a porta de
      entrada de TODA a fabrica so abria depois que o painel de administracao
      (um Next.js, com start_period de 20 s) se declarasse saudavel. Sem o
      studio na frente, o kong sobe direto atras do banco.
    · KONG COM 8 WORKERS. O nginx do kong abre um worker por nucleo, e cada um
      carrega os buffers de 160k declarados no compose. Para dez pessoas numa
      fabrica, 1 worker sobra — e sao centenas de MB a menos.

  O QUE NAO SAI, DE PROPOSITO:
    db, kong, rest, auth, storage, realtime - e a fabrica funcionando.
    functions - e a funcao de criar usuario (servidor\funcoes); e do admin, mas
      quando ele precisa, precisa na hora.

  COMO USAR
    Ver o que mudaria, sem mexer em nada:
      .\servidor\enxugar-supabase.ps1 -Simular
    Aplicar (derruba os quatro conteineres e reinicia o kong; ~15 s de porta
    fechada, entao fora do horario da fabrica):
      .\servidor\enxugar-supabase.ps1
    Subir o painel do Supabase quando precisar dele:
      .\servidor\enxugar-supabase.ps1 -Painel
      .\servidor\enxugar-supabase.ps1 -FecharPainel
    Voltar tudo como veio de fabrica:
      .\servidor\enxugar-supabase.ps1 -Desfazer

  O ORIGINAL FICA GUARDADO em docker-compose.yml.antes-do-enxugamento, ao lado
  do proprio compose. -Desfazer e uma copia de arquivo, nao uma tentativa de
  desfazer edicao por edicao.
#>
[CmdletBinding()]
param(
  [string] $Docker = 'C:\supabase\docker',
  [switch] $Simular,
  [switch] $Desfazer,
  [switch] $Painel,
  [switch] $FecharPainel
)

$ErrorActionPreference = 'Stop'
$Compose  = Join-Path $Docker 'docker-compose.yml'
$Guardado = Join-Path $Docker 'docker-compose.yml.antes-do-enxugamento'

function Falar($t) { Write-Host $t }
function DockerExe {
  $c = Get-Command docker -ErrorAction SilentlyContinue
  if ($c) { return $c.Source }
  throw 'docker nao encontrado no PATH'
}

if (-not (Test-Path $Compose)) { throw "compose nao encontrado: $Compose" }
$docker = DockerExe

# ------------------------------------------------------------------ painel
# Sobe (ou derruba) studio + meta sem tocar em mais nada. So faz sentido depois
# de aplicado; antes disso eles ja estao de pe o tempo todo.
if ($Painel -or $FecharPainel) {
  if ($FecharPainel) {
    & $docker rm -f supabase-studio supabase-meta 2>&1 | Out-Null
    Falar 'painel do Supabase fechado (studio e meta fora)'
  } else {
    & $docker compose -f $Compose --profile painel up -d studio meta
    Falar 'painel do Supabase de pe — http://localhost:8000 (usuario e senha do .env)'
  }
  exit 0
}

# ----------------------------------------------------------------- desfazer
if ($Desfazer) {
  if (-not (Test-Path $Guardado)) { throw "nao ha copia guardada em $Guardado" }
  Copy-Item $Guardado $Compose -Force
  Falar 'compose original restaurado'
  & $docker compose -f $Compose up -d
  exit 0
}

# ------------------------------------------------------------------ aplicar
$texto = Get-Content $Compose -Raw

# As ancoras sao trechos LITERAIS do compose de fabrica. Se a Supabase mudar o
# arquivo numa atualizacao, uma delas deixa de casar — e o script para e diz
# qual, em vez de gravar um YAML pela metade.
$edicoes = @(
  @{
    nome = 'kong deixa de esperar o studio'
    de   = "    depends_on:`n      studio:`n        condition: service_healthy`n"
    para = ''
  },
  @{
    nome = 'kong com 1 worker de nginx (em vez de 1 por nucleo)'
    de   = "      KONG_ROUTER_FLAVOR: expressions`n"
    para = "      KONG_ROUTER_FLAVOR: expressions`n      # Um worker basta para a fabrica; o padrao abre um por nucleo, e cada`n      # um carrega os buffers de 160k declarados abaixo. Ver enxugar-supabase.ps1.`n      KONG_NGINX_WORKER_PROCESSES: `"1`"`n"
  },
  @{
    nome = 'studio sob demanda (perfil painel)'
    de   = "  studio:`n    container_name: supabase-studio`n"
    para = "  studio:`n    container_name: supabase-studio`n    profiles: [painel]`n"
  },
  @{
    nome = 'meta sob demanda (so o painel usa)'
    de   = "  meta:`n    container_name: supabase-meta`n"
    para = "  meta:`n    container_name: supabase-meta`n    profiles: [painel]`n"
  },
  @{
    nome = 'pooler fora (ninguem conecta direto no Postgres)'
    de   = "  supavisor:`n    container_name: supabase-pooler`n"
    para = "  supavisor:`n    container_name: supabase-pooler`n    profiles: [pooler]`n"
  },
  @{
    nome = 'imgproxy fora (nao ha transformacao de imagem no programa)'
    de   = "  imgproxy:`n    container_name: supabase-imgproxy`n"
    para = "  imgproxy:`n    container_name: supabase-imgproxy`n    profiles: [imagens]`n"
  },
  @{
    nome = 'storage deixa de esperar o imgproxy'
    de   = "      imgproxy:`n        condition: service_started`n"
    para = ''
  }
)

$feitas = @(); $jaEstavam = @(); $faltando = @()
foreach ($e in $edicoes) {
  if ($texto.Contains($e.de) -and ($e.para -eq '' -or -not $texto.Contains($e.para))) {
    $texto = $texto.Replace($e.de, $e.para)
    $feitas += $e.nome
  } elseif ($e.para -ne '' -and $texto.Contains($e.para)) {
    $jaEstavam += $e.nome
  } elseif ($e.para -eq '' -and -not $texto.Contains($e.de)) {
    $jaEstavam += $e.nome
  } else {
    $faltando += $e.nome
  }
}

if ($faltando.Count) {
  Falar 'NAO APLIQUEI NADA — o compose nao esta como o script espera:'
  $faltando | ForEach-Object { Falar "  · $_" }
  Falar 'Provavel atualizacao do Supabase. Conferir o docker-compose.yml a mao.'
  exit 1
}

if ($Simular) {
  Falar 'o que MUDARIA:'
  $feitas    | ForEach-Object { Falar "  + $_" }
  $jaEstavam | ForEach-Object { Falar "  = $_ (ja estava)" }
  exit 0
}

if (-not $feitas.Count) { Falar 'nada a fazer — o compose ja esta enxuto'; exit 0 }

if (-not (Test-Path $Guardado)) { Copy-Item $Compose $Guardado -Force; Falar "original guardado em $Guardado" }

# Grava sem BOM: o compose e lido por um parser YAML, nao pelo PowerShell.
[IO.File]::WriteAllText($Compose, $texto, (New-Object Text.UTF8Encoding($false)))
$feitas | ForEach-Object { Falar "  + $_" }

# O compose so vale se o proprio docker aceitar o arquivo. Se nao aceitar,
# volta o original NA HORA — um compose quebrado aqui e a fabrica sem servidor
# no dia seguinte.
$conf = & $docker compose -f $Compose config -q 2>&1 | Out-String
if ($LASTEXITCODE -ne 0) {
  Copy-Item $Guardado $Compose -Force
  Falar 'compose recusado pelo docker — original restaurado. Saida:'
  Falar $conf
  exit 1
}

Falar 'derrubando o que saiu da pilha...'
& $docker rm -f supabase-studio supabase-meta supabase-pooler supabase-imgproxy 2>&1 | Out-Null
Falar 'aplicando (o kong reinicia para pegar o worker unico)...'
& $docker compose -f $Compose up -d --remove-orphans
Falar ''
Falar 'pronto. Conferir com:  docker stats --no-stream'
Falar 'O teto do WSL (.wslconfig) continua em 4 GB de proposito: baixe para 3 GB'
Falar 'so depois de ver, por alguns dias, quanto a VM passou a usar de fato.'
