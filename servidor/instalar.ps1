<#
  Instalador do servidor da fabrica - Gerador-OS.

  Faz sozinho os passos 2 a 7 e 9 do README: clona o Supabase, gera e grava as
  chaves, sobe os containers, cria as tabelas, gera o certificado HTTPS, migra
  os dados do backup, sobe o nginx e importa as imagens dos desenhos.

  O QUE ELE NAO FAZ (e por que):
    - instalar o Docker Desktop: exige tela, administrador e reiniciar;
    - criar as contas de login: cada pessoa tem a sua, e voce escolhe as senhas;
    - instalar o certificado nas maquinas: e um comando por maquina, com
      administrador.
  Ele diz, no fim, exatamente o que sobrou para voce.

  COMO RODAR, no PowerShell, dentro da pasta do Gerador-OS no servidor:

      .\servidor\instalar.ps1 -IP 192.168.0.50

  PODE RODAR DE NOVO. Cada etapa confere se ja foi feita e pula. Chaves nao sao
  regeradas por acidente - trocar o JWT_SECRET invalidaria a chave que ja esta
  configurada em todas as maquinas.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string] $IP,                                   # IP fixo do servidor na rede
  [string] $PastaSupabase = "C:\supabase",        # onde clonar o Supabase
  [string] $Backup = "",                          # backup a migrar (vazio = o mais recente)
  [string] $PastaDesenhos = "",                   # vazio = procura sozinho (ver abaixo)
  [switch] $RefazerChaves                         # forca gerar chaves novas
)

$ErrorActionPreference = "Stop"
$RaizApp = Split-Path -Parent $PSScriptRoot
$passo = 0

function Titulo($texto) {
  $script:passo++
  Write-Host ""
  Write-Host ("[{0}] {1}" -f $script:passo, $texto) -ForegroundColor Cyan
}
function Ok($texto)    { Write-Host ("    ok  " + $texto) -ForegroundColor Green }
function Pulo($texto)  { Write-Host ("    ja  " + $texto) -ForegroundColor DarkGray }
function Aviso($texto) { Write-Host ("    !   " + $texto) -ForegroundColor Yellow }
function Parar($texto) { Write-Host ""; Write-Host ("ERRO: " + $texto) -ForegroundColor Red; exit 1 }

function LerEnv($arquivo, $chave) {
  if (-not (Test-Path $arquivo)) { return $null }
  foreach ($linha in Get-Content $arquivo) {
    if ($linha -match ("^\s*" + [regex]::Escape($chave) + "\s*=\s*(.+?)\s*$")) { return $Matches[1] }
  }
  return $null
}

# ---------------------------------------------------------------- pre-requisitos
Titulo "Conferindo o que precisa estar instalado"
foreach ($prog in @("docker", "node", "git")) {
  if (-not (Get-Command $prog -ErrorAction SilentlyContinue)) {
    Parar ("'$prog' nao encontrado. Instale antes de continuar. " +
           "O Docker Desktop precisa estar ABERTO e com o motor rodando.")
  }
  Ok $prog
}
docker info | Out-Null
if (-not $?) { Parar "O Docker esta instalado mas nao esta respondendo. Abra o Docker Desktop e espere aparecer 'Engine running'." }
Ok "Docker respondendo"

if ($IP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { Parar "-IP nao parece um endereco IPv4: $IP" }
$meus = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -eq $IP })
if (-not $meus) {
  Aviso "O IP $IP nao esta nesta maquina. Se ainda vai fixa-lo, tudo bem; se digitou errado, pare agora (Ctrl+C)."
  Start-Sleep -Seconds 5
}

# ------------------------------------------------------------------- Supabase
Titulo "Supabase local em $PastaSupabase"
$PastaDocker = Join-Path $PastaSupabase "docker"
if (Test-Path (Join-Path $PastaDocker "docker-compose.yml")) {
  Pulo "ja clonado"
} else {
  # a pasta-mae so precisa ser criada quando ela mesma nao existe; com
  # -PastaSupabase C:\supabase a mae e "C:\", e o Windows recusa cria-la
  $PastaMae = Split-Path -Parent $PastaSupabase
  if ($PastaMae -and -not (Test-Path $PastaMae)) {
    New-Item -ItemType Directory -Force -Path $PastaMae | Out-Null
  }
  git clone --depth 1 https://github.com/supabase/supabase $PastaSupabase
  if ($LASTEXITCODE -ne 0) { Parar "falha ao clonar o Supabase" }
  Ok "clonado"
}
$EnvArq = Join-Path $PastaDocker ".env"
if (-not (Test-Path $EnvArq)) {
  Copy-Item (Join-Path $PastaDocker ".env.example") $EnvArq
  Ok ".env criado a partir do exemplo"
}

# --------------------------------------------------------------------- chaves
Titulo "Chaves do servidor"
# A chave de demonstracao e, por definicao, a que vem no .env.example - comparar
# com ele acerta sempre. A versao anterior procurava um prefixo fixo da chave
# demo e passou a errar quando o Supabase mudou a ordem dos campos do JWT: ela
# concluia "ja geradas" sobre um .env recem-copiado do exemplo, e o servidor
# subia com as credenciais publicas do repositorio deles. Confere tambem o
# JWT_SECRET, porque sao as duas pontas do mesmo par.
$ExemploArq   = Join-Path $PastaDocker ".env.example"
$AnonAtual    = LerEnv $EnvArq "ANON_KEY"
$AnonExemplo  = LerEnv $ExemploArq "ANON_KEY"
$SegredoAtual = LerEnv $EnvArq "JWT_SECRET"
$SegredoExemplo = LerEnv $ExemploArq "JWT_SECRET"
$JaGerado = ($AnonAtual -ne $null -and $AnonAtual.Length -gt 100 -and
             $AnonAtual -ne $AnonExemplo -and $SegredoAtual -ne $SegredoExemplo)
if ($JaGerado -and -not $RefazerChaves) {
  Pulo "ja geradas (use -RefazerChaves so se souber o que isso implica)"
} else {
  if ($JaGerado) { Aviso "REGERANDO as chaves: todas as maquinas terao de ser reconfiguradas com a chave nova." }
  $saida = & node (Join-Path $PSScriptRoot "gerar-chaves.js") --escrever-env $EnvArq
  if ($LASTEXITCODE -ne 0) { Parar "falha ao gerar as chaves" }
  Ok "geradas e gravadas no .env"
}
$AnonKey = LerEnv $EnvArq "ANON_KEY"
$SrvKey  = LerEnv $EnvArq "SERVICE_ROLE_KEY"
if (-not $AnonKey -or -not $SrvKey) { Parar "nao consegui ler ANON_KEY/SERVICE_ROLE_KEY de $EnvArq" }

# ------------------------------------------------------------------ containers
Titulo "Subindo o Supabase (a primeira vez baixa varios GB e demora)"
Push-Location $PastaDocker
docker compose up -d
$falhou = ($LASTEXITCODE -ne 0)
Pop-Location
if ($falhou) { Parar "docker compose up falhou" }

Write-Host "    esperando a API responder..." -NoNewline
$pronto = $false
foreach ($tentativa in 1..90) {
  Start-Sleep -Seconds 2
  $codigo = 0
  try {
    $r = Invoke-WebRequest -Uri "http://localhost:8000/rest/v1/" -Headers @{ apikey = $AnonKey } `
         -UseBasicParsing -TimeoutSec 5
    $codigo = [int] $r.StatusCode
  } catch [System.Net.WebException] {
    # Qualquer resposta HTTP ja prova que o Kong esta de pe - e o Invoke-WebRequest
    # trata 4xx como excecao, entao o codigo tem de ser lido daqui. Importa em
    # especial o 403: o Supabase passou a servir a raiz /rest/v1/ como rota so de
    # admin, e a chave anonima e recusada ali mesmo com tudo funcionando. Antes
    # disto a espera ia ate o fim e a instalacao parava dizendo que a API nao
    # respondeu, quando ela estava respondendo o tempo inteiro.
    if ($_.Exception.Response) { $codigo = [int] $_.Exception.Response.StatusCode }
  } catch { }
  if ($codigo -ge 200 -and $codigo -lt 500) { $pronto = $true; break }
  Write-Host "." -NoNewline
}
Write-Host ""
if (-not $pronto) { Parar "a API nao respondeu em 3 minutos. Veja 'docker compose logs' em $PastaDocker" }
Ok "API de pe em http://localhost:8000"

# --------------------------------------------------------------------- tabelas
Titulo "Criando as tabelas"
$schema = Join-Path $PSScriptRoot "schema.sql"
# Copia o arquivo para dentro do container em vez de mandar pelo cano: assim os
# acentos dos comentarios nao dependem da codificacao do PowerShell.
docker cp $schema supabase-db:/tmp/schema.sql
if ($LASTEXITCODE -ne 0) { Parar "nao consegui copiar o schema para o container supabase-db" }
docker exec supabase-db psql -U postgres -d postgres -v ON_ERROR_STOP=1 -f /tmp/schema.sql
if ($LASTEXITCODE -ne 0) { Parar "o schema.sql falhou. A mensagem do psql esta acima." }
Ok "tabelas, permissoes, realtime e bucket dos desenhos"

# A tranca dos papeis: `set_user_role` confere no BANCO se quem chamou e admin.
# Esconder a tela no navegador nao vale de nada - a pagina roda na maquina de
# quem usa, e o papel e uma variavel de JavaScript.
$sqlPapeis = Join-Path (Split-Path -Parent $PSScriptRoot) "sql\supabase-admin-roles.sql"
if (Test-Path $sqlPapeis) {
  docker cp $sqlPapeis supabase-db:/tmp/papeis.sql | Out-Null
  docker exec supabase-db psql -U postgres -d postgres -f /tmp/papeis.sql | Out-Null
  Ok "set_user_role instalada (quem muda papel e conferido no servidor)"
} else {
  Aviso "nao achei sql\supabase-admin-roles.sql - o botao de promover a admin nao vai funcionar"
}

# --------------------------------------------------------------------- funcoes
Titulo "Funcoes do servidor"
# Criar conta pelo navegador exigiria ligar o auto-cadastro, e a chave anonima
# esta dentro de toda pagina aberta na fabrica - qualquer pessoa na rede criaria
# a propria conta e leria tudo. A funcao abaixo usa a chave de servico, que fica
# so no servidor, e confere la dentro se quem chamou e admin.
$origemFuncoes = Join-Path $PSScriptRoot "funcoes"
$destinoFuncoes = Join-Path $PastaSupabase "docker\volumes\functions"
if (Test-Path $origemFuncoes) {
  foreach ($f in Get-ChildItem $origemFuncoes -Directory) {
    $alvo = Join-Path $destinoFuncoes $f.Name
    New-Item -ItemType Directory -Force -Path $alvo | Out-Null
    Copy-Item (Join-Path $f.FullName "*") $alvo -Recurse -Force
    Ok ("funcao " + $f.Name)
  }
} else { Pulo "nenhuma funcao para instalar" }

# Com a criacao de contas feita pelo admin, o auto-cadastro deixa de ter uso -
# e ligado ele deixaria qualquer um na rede criar conta e ler a fabrica inteira.
$dis = LerEnv $EnvArq "DISABLE_SIGNUP"
if ($dis -ne "true") {
  $texto = Get-Content $EnvArq -Raw
  if ($texto -match '(?m)^DISABLE_SIGNUP=') { $texto = $texto -replace '(?m)^DISABLE_SIGNUP=.*$', 'DISABLE_SIGNUP=true' }
  else { $texto = $texto.TrimEnd() + "`nDISABLE_SIGNUP=true`n" }
  Set-Content -Path $EnvArq -Value $texto -Encoding utf8
  Push-Location $PastaDocker
  docker compose up -d auth | Out-Null
  Pop-Location
  Ok "auto-cadastro desligado (contas so pelo admin, dentro do programa)"
} else { Pulo "auto-cadastro ja desligado" }

# ---------------------------------------------------------------------- dados
Titulo "Migrando os dados"
if (-not $Backup) {
  $ultimo = Get-ChildItem (Join-Path $RaizApp "backups") -Filter "BACKUP-COMPLETO-*.json" -ErrorAction SilentlyContinue |
            Sort-Object Name | Select-Object -Last 1
  if ($ultimo) { $Backup = $ultimo.FullName }
}
if (-not $Backup -or -not (Test-Path $Backup)) {
  Aviso "nenhum backup encontrado - pule esta etapa e rode migrar-do-backup.js a mao depois"
} else {
  Write-Host "    usando $(Split-Path -Leaf $Backup)"
  & node (Join-Path $PSScriptRoot "migrar-do-backup.js") --url "http://localhost:8000" --key $SrvKey --arq $Backup
  if ($LASTEXITCODE -ne 0) {
    Aviso "a migracao nao gravou (o servidor ja tinha dados?). Nada foi sobrescrito."
  } else { Ok "dados no servidor" }
}

# ----------------------------------------------------------------- certificado
Titulo "Certificado HTTPS para $IP"
$PastaTls = Join-Path $PSScriptRoot "tls"
if ((Test-Path (Join-Path $PastaTls "servidor.crt")) -and -not $RefazerChaves) {
  Pulo "ja existe (apague servidor\tls para refazer)"
} else {
  & node (Join-Path $PSScriptRoot "gerar-certificado.js") --ip $IP
  if ($LASTEXITCODE -ne 0) { Parar "falha ao gerar o certificado" }
}

# ----------------------------------------------------------------------- nginx
Titulo "Publicando o app em https://$IP"

# Como cada maquina descobre com quem falar, sem ninguem preencher endereco e
# chave computador por computador. So a chave ANONIMA entra aqui: ela ja viaja
# dentro de todo navegador que abre o app, entao publica-la nao concede nada.
# Fica em servidor\tls (ignorado pelo git) por ser desta instalacao, nao do
# programa.
$LocalJson = Join-Path $PSScriptRoot "tls\servidor-local.json"
@{ url = "https://$IP"; key = $AnonKey } | ConvertTo-Json -Compress |
  Set-Content -Path $LocalJson -Encoding utf8 -NoNewline
Ok "endereco publicado para as maquinas se configurarem sozinhas"

Push-Location $RaizApp
docker compose -f (Join-Path "servidor" "docker-compose.app.yml") up -d
$falhou = ($LASTEXITCODE -ne 0)
Pop-Location
if ($falhou) { Parar "nao consegui subir o servidor web" }
Ok "no ar"

# -------------------------------------------------------------------- firewall
Titulo "Liberando as portas 80 e 443 no firewall"
foreach ($porta in @(80, 443)) {
  $nome = "Gerador-OS $porta"
  if (Get-NetFirewallRule -DisplayName $nome -ErrorAction SilentlyContinue) { Pulo "$nome" }
  else {
    try {
      New-NetFirewallRule -DisplayName $nome -Direction Inbound -LocalPort $porta `
        -Protocol TCP -Action Allow | Out-Null
      Ok $nome
    } catch { Aviso "nao consegui criar a regra $nome (rode o PowerShell como administrador)" }
  }
}

# -------------------------------------------------------------------- desenhos
Titulo "Imagens dos desenhos tecnicos"
$pd = $PastaDesenhos
if (-not $pd) {
  # Achada por padrao, e nao escrita aqui de propriedade: o nome tem acento, e o
  # PowerShell 5.1 le .ps1 como ANSI quando o arquivo nao tem BOM - um acento
  # cravado no codigo viraria lixo e a pasta "nao existiria".
  # A de riscos de corte ("-grades de corte") fica de fora: sao PDFs, nao os
  # desenhos do cadastro.
  $achada = Get-ChildItem $RaizApp -Directory |
            Where-Object { $_.Name -like "Desenhos*" -and $_.Name -notlike "*grades*" } |
            Select-Object -First 1
  if ($achada) { $pd = $achada.FullName }
} elseif (-not [System.IO.Path]::IsPathRooted($pd)) { $pd = Join-Path $RaizApp $pd }
if (-not (Test-Path $pd)) {
  Aviso "pasta '$pd' nao encontrada - as folhas de OS abrirao sem desenho ate voce importar"
} else {
  & node (Join-Path $PSScriptRoot "importar-desenhos-da-pasta.js") --pasta $pd `
    --url "http://localhost:8000" --key $SrvKey
  if ($LASTEXITCODE -ne 0) { Aviso "a importacao dos desenhos nao terminou; veja a mensagem acima" }
}

# --------------------------------------------------------------------- resumo
$resumo = @"
=====================================================================
 SERVIDOR DA FABRICA PRONTO
=====================================================================
 Endereco do app .... https://$IP
 Painel do Supabase . http://localhost:8000  (so aqui no servidor)
 Chave ANON ......... $AnonKey

 FALTA VOCE FAZER, nesta ordem:

 1) CONTAS. Painel -> Authentication -> Users -> Add user. Uma por
    pessoa: sem conta ninguem entra, e sem entrar nao ve nada.
    Depois, o papel de admin para a sua, no SQL Editor:

      INSERT INTO user_roles (user_id, role)
      SELECT id, 'admin' FROM auth.users WHERE email = 'SEU-EMAIL'
      ON CONFLICT (user_id) DO UPDATE SET role = 'admin';

 2) CERTIFICADO em cada computador. Copie servidor\tls\ca.crt para a
    maquina e rode o PowerShell como ADMINISTRADOR:

      Import-Certificate -FilePath .\ca.crt -CertStoreLocation Cert:\LocalMachine\Root

    Feche e reabra o navegador depois.

 3) APONTAR cada maquina. Abra https://$IP, va em Configuracoes ->
    Servidor da fabrica e preencha:
      Endereco: https://$IP
      Chave:    a ANON acima
    "Testar conexao", depois "Salvar e recarregar". A barra lateral
    deve mostrar "Servidor da fabrica".

 4) IP FIXO para esta maquina (ou reserva no roteador). Se o IP mudar,
    todas as maquinas param de achar o servidor e o certificado deixa
    de valer.

 5) LOGON AUTOMATICO do Windows e no-break. O Docker Desktop so sobe
    dentro de uma sessao logada: sem isso, um reinicio de madrugada
    deixa a fabrica sem sistema pela manha.

 CONFERIR: https://$IP/vendor-check.html deve mostrar 7 linhas "ok".
=====================================================================
"@
Write-Host $resumo -ForegroundColor White
# Guardado tambem em arquivo, na pasta que o git ignora.
New-Item -ItemType Directory -Force -Path $PastaTls | Out-Null
$resumo | Out-File -FilePath (Join-Path $PastaTls "resumo-instalacao.txt") -Encoding utf8
Write-Host "(uma copia deste resumo ficou em servidor\tls\resumo-instalacao.txt)" -ForegroundColor DarkGray
