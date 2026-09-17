<#
  Publica o app SO nos enderecos da fabrica, em vez de em todos.

  O QUE SE GANHA
  Escrito "443:443" no compose, o Docker prende a porta em TODOS os enderecos da
  maquina (0.0.0.0 e [::]), inclusive nos loopbacks que outros programas criam
  para si -- o Warsaw/Topaz, dos bancos, levanta um em 54.232.189.113 e precisa
  da 443 nele. Hoje os dois convivem, porque um bind especifico ganha do curinga
  naquele endereco. E uma disputa que nao precisa existir, e nao um incendio.

  POR QUE ISTO E UM SCRIPT, E COM CONFERENCIA ANTES
  Em 17/09/2026 a troca foi tentada a mao, duas vezes, e as duas derrubaram o app
  por cerca de um minuto cada. O motivo nao era o compose:

      failed to bind host port 127.0.0.1:32787/tcp: address already in use

  Para cada publicacao em endereco de REDE o Docker Desktop reserva um auxiliar
  seu em 127.0.0.1:32xxx, e o segundo auxiliar da mesma porta de container bate
  no primeiro. Medido fora da producao, com containers descartaveis em portas
  livres, depois do reinicio de 17/09 (tres tentativas de cada, sempre igual):

      .200 + 127.0.0.1   na 80 e na 443   ->  sobe
      .200 + .158        na 80 e na 443   ->  address already in use
      .200 + .158 + 127.0.0.1             ->  address already in use
      .200 duas vezes na mesma porta 80   ->  address already in use

  Ou seja: cabe UM SO endereco de rede por porta de container, e 127.0.0.1 nao
  conta. Nao e a faixa de portas do Windows -- a porta pedida esta livre e fica
  fora da faixa dinamica, que aqui comeca em 49152 -- nem estado sujo do motor:
  o reinicio de 17/09/2026 nao mudou nada disto.

  O QUE ISTO CUSTA
  Os tres enderecos de uma vez (.200, .158 e 127.0.0.1) nao cabem. Restringir a
  443 hoje obriga a escolher entre servir o app pelo cabo ou pelo Wi-Fi -- quem
  entra pelo endereco que ficar de fora perde a pagina. Por isso o padrao deste
  script continua sendo os tres, que reprovam na conferencia: e para a escolha
  ser feita a mao, sabendo o que se perde, e nao por um script decidindo.
  Enquanto a escolha nao vale a pena, deixar o curinga e o certo -- a disputa
  pela 443 nao esta acontecendo: o Warsaw tem o endereco dele e funciona.

  Por isso o script COMECA testando, com um container descartavel, a MESMA forma
  que vai escrever no compose -- os enderecos pedidos, na 80 e na 443 -- e so
  mexe na producao se ela subir.

  USO

    .\servidor\restringir-443.ps1
        So confere. Nao toca em nada. Diz se a maquina aceita a mudanca hoje.

    .\servidor\restringir-443.ps1 -Aplicar
        Confere, e SO SE PASSAR aplica: escreve o compose, down, up, e verifica
        os tres enderecos. Se a verificacao falhar, volta o compose de antes e
        sobe de novo, sozinho.

    .\servidor\restringir-443.ps1 -Aplicar -Enderecos '193.168.0.200','127.0.0.1'
        A UNICA forma que passa na conferencia hoje: o app fica servido pelo cabo
        e por localhost, e PARA de responder pelo Wi-Fi (192.168.1.158). So faz
        sentido no dia em que alguem precisar da 443 num endereco de rede que
        hoje o curinga ocupa; enquanto ninguem precisa, nao rode.

  QUANDO RODAR
  Fora do expediente, depois do backup completo do dia e depois do reinicio
  (.\servidor\desligar-servidor.ps1 -Reiniciar). O app fica fora alguns segundos
  entre o `down` e o `up`.

  O QUE FICA DE RISCO DEPOIS DE APLICADO
  Com endereco fixo no lugar do curinga, o container SO SOBE SE O ENDERECO
  EXISTIR. Se a maquina voltar de um reboot com o Wi-Fi desconectado, o
  192.168.1.158 nao existe, o `restart: always` nao salva e o app nao volta.
  A saida rapida, nesse caso, e rodar de novo com -Enderecos sem o .158.
#>
[CmdletBinding()]
param(
  [switch]   $Aplicar,
  [string[]] $Enderecos = @('193.168.0.200', '192.168.1.158', '127.0.0.1'),
  [string]   $Compose   = 'servidor\docker-compose.app.yml',
  [int]      $Paciencia = 40
)

# 'Continue', e nao 'Stop': no PowerShell 5.1, a saida de ERRO de um programa
# externo (o docker) vira ErrorRecord, e com 'Stop' o script morre ali mesmo --
# engolindo justamente a mensagem que explica o que fazer a seguir. Quem decide
# se deu certo aqui e o $LASTEXITCODE, que e o que o docker de fato responde.
$ErrorActionPreference = 'Continue'
$Raiz = Split-Path -Parent $PSScriptRoot
Set-Location $Raiz
$ComposePath = Join-Path $Raiz $Compose

function Diz($t)   { Write-Host $t }
function Bom($t)   { Write-Host "  ok    $t" -ForegroundColor Green }
function Mau($t)   { Write-Host "  FALHA $t" -ForegroundColor Red }

# ---------------------------------------------------------------- 1. conferir
Diz ''
Diz '1) OS ENDERECOS EXISTEM NESTA MAQUINA?'
$faltando = @()
foreach ($ip in $Enderecos) {
  $tem = Get-NetIPAddress -IPAddress $ip -ErrorAction SilentlyContinue
  if ($tem) { Bom $ip } else { Mau "$ip  -- nao existe aqui"; $faltando += $ip }
}
if ($faltando.Count) {
  Diz ''
  Diz "PARADO. Sem $($faltando -join ', ') o container nao sobe."
  Diz 'Se o Wi-Fi voltou desconectado, ou rode fixar-ip-wifi.ps1, ou repita este'
  Diz "script com  -Enderecos '193.168.0.200','127.0.0.1'"
  exit 1
}

Diz ''
Diz '2) ESTA MAQUINA ACEITA A PUBLICACAO QUE ESTE SCRIPT VAI ESCREVER?'
Diz '   (container descartavel, nas mesmas portas de container -- 80 e 443 --'
Diz '    que a producao usa, mas em portas de host livres)'
# O TESTE PRECISA TER A FORMA DO QUE VAI SER ESCRITO, e nao uma forma parecida.
# Ate 17/09/2026 ele subia dois binds no MESMO endereco apontando os dois para a
# porta 80 do container -- uma forma que nao existe no compose e que falha por
# conta propria. Media coisa errada: reprovava tambem a variante que FUNCIONA
# (cabo + 127.0.0.1), a mesma que o cabecalho recomenda.
#
# O QUE DE FATO LIMITA, medido nesta maquina em 17/09/2026 com containers
# descartaveis: cabe UM SO endereco que nao seja 127.0.0.1 por porta de
# container. Para cada publicacao em endereco de rede o Docker Desktop reserva
# um auxiliar seu em 127.0.0.1:32xxx, e o segundo do mesmo par bate no primeiro:
#
#   .200 + 127.0.0.1  na 80 e na 443      -> sobe (3 tentativas)
#   .200 + .158       na 80 e na 443      -> 'address already in use' (2)
#   .200 + .158 + 127.0.0.1               -> 'address already in use' (3)
#   .200 duas vezes na mesma porta 80     -> 'address already in use' (3)
#
# Nao e estado sujo do motor: o reinicio de 17/09 nao mudou nada disto.
$nome = 'teste-bind-443'
$lixo = Join-Path $env:TEMP 'restringir-443-docker.txt'
& docker rm -f $nome 2>$lixo | Out-Null
$publica = @()
foreach ($ip in $Enderecos) { $publica += '-p'; $publica += "${ip}:18080:80" }
foreach ($ip in $Enderecos) { $publica += '-p'; $publica += "${ip}:18443:443" }
& docker run -d --rm --name $nome @publica nginx:alpine 2>$lixo | Out-Null
$passou = ($LASTEXITCODE -eq 0)
# -Raw num arquivo vazio devolve $null, e .Trim() em $null grita: o teste que
# PASSA nao escreve nada aqui, e era ele quem levantava o erro vermelho.
$bruto  = Get-Content $lixo -Raw -ErrorAction SilentlyContinue
$motivo = if ($bruto) { $bruto.Trim() } else { '' }
& docker rm -f $nome 2>$lixo | Out-Null
if (-not $passou) {
  $deRede = @($Enderecos | Where-Object { $_ -ne '127.0.0.1' })
  Mau "o Docker recusa esta publicacao: $($Enderecos -join ', ')"
  if ($motivo) { Diz ''; Diz "   $motivo" }
  Diz ''
  if ($deRede.Count -gt 1) {
    Diz "PARADO, e nada foi tocado. Sao $($deRede.Count) enderecos de rede na mesma"
    Diz 'porta de container, e aqui cabe um so. Repita com um deles:'
    Diz ''
    Diz "    .\servidor\restringir-443.ps1 -Aplicar -Enderecos '$($deRede[0])','127.0.0.1'"
    Diz ''
    Diz "O custo de escolher: o endereco que ficar de fora ($($deRede[1..($deRede.Count-1)] -join ', '))"
    Diz 'para de servir o app -- quem entra por ele perde a pagina.'
  } else {
    Diz 'PARADO, e nada foi tocado. Nem com um endereco de rede so o Docker aceita'
    Diz 'a publicacao hoje. Deixe como esta: a disputa pela 443 nao esta'
    Diz 'acontecendo (o Warsaw tem o endereco dele e funciona). Tentar de novo so'
    Diz 'depois de atualizar o Docker Desktop.'
  }
  exit 1
}
Bom "a publicacao em $($Enderecos -join ', ') sobe -- da para aplicar"

if (-not $Aplicar) {
  Diz ''
  Diz 'Conferencia apenas. Rode de novo com -Aplicar para valer.'
  exit 0
}

# ---------------------------------------------------------------- 2. aplicar
Diz ''
Diz '3) APLICANDO'
$backup = "$ComposePath.antes-de-restringir-443"
Copy-Item $ComposePath $backup -Force
Bom "compose de antes guardado em $(Split-Path -Leaf $backup)"

$linhas = @()
foreach ($ip in $Enderecos) { $linhas += "      - `"${ip}:80:80`"" }
foreach ($ip in $Enderecos) { $linhas += "      - `"${ip}:443:443`"" }
$blocoNovo = "    ports:`n" + ($linhas -join "`n") + "`n"

$txt = Get-Content $ComposePath -Raw
$alvo = "    ports:`r`n      - `"80:80`"              # só redireciona para o 443`r`n      - `"443:443`"`r`n"
if ($txt -notlike "*$alvo*") {
  $alvo = $alvo -replace "`r`n", "`n"
}
if ($txt -notlike "*$alvo*") {
  Mau 'nao achei o bloco de portas original no compose -- nada foi mudado'
  exit 1
}
($txt -replace [regex]::Escape($alvo), $blocoNovo) | Set-Content $ComposePath -Encoding utf8 -NoNewline
Bom 'compose reescrito'

& docker compose -f $Compose config 2>$lixo | Out-Null
if ($LASTEXITCODE -ne 0) {
  Mau 'o compose novo nao e YAML valido -- voltando'
  Copy-Item $backup $ComposePath -Force
  exit 1
}
Bom 'YAML valido'

Diz '   down (solta as portas de verdade -- `up -d` sozinho nao solta)'
& docker compose -f $Compose down 2>$lixo | Out-Null
Start-Sleep -Seconds 5
Diz '   up'
& docker compose -f $Compose up -d 2>$lixo | Out-Null
$subiu = ($LASTEXITCODE -eq 0)
$motivoUp = if (Test-Path $lixo) { (Get-Content $lixo -Raw).Trim() } else { '' }

# ---------------------------------------------------------------- 3. verificar
$ok = $subiu
if ($subiu) {
  Diz ''
  Diz '4) VERIFICANDO'
  Start-Sleep -Seconds 4
  foreach ($ip in $Enderecos) {
    $t = Test-NetConnection -ComputerName $ip -Port 443 -WarningAction SilentlyContinue
    if ($t.TcpTestSucceeded) { Bom "$ip responde na 443" } else { Mau "$ip NAO responde"; $ok = $false }
  }
} else {
  Mau 'o container nao subiu'
  if ($motivoUp) { Diz "   $motivoUp" }
}

if (-not $ok) {
  Diz ''
  Diz 'VOLTANDO ATRAS -- o app e mais importante que a mudanca.'
  Copy-Item $backup $ComposePath -Force
  & docker compose -f $Compose up -d --force-recreate 2>$lixo | Out-Null
  Start-Sleep -Seconds 4
  foreach ($ip in $Enderecos) {
    $t = Test-NetConnection -ComputerName $ip -Port 443 -WarningAction SilentlyContinue
    if ($t.TcpTestSucceeded) { Bom "$ip voltou" } else { Mau "$ip AINDA FORA -- olhe agora" }
  }
  exit 1
}

Diz ''
Bom 'aplicado. O app atende so nos enderecos da fabrica.'
Diz '   Confira tambem de OUTRA maquina antes de considerar fechado.'
Diz "   Para desfazer:  Copy-Item '$backup' '$ComposePath' -Force"
Diz "                   docker compose -f $Compose up -d --force-recreate"
