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

  Para cada publicacao com endereco FIXO, o Docker Desktop reserva uma porta
  interna sua em 127.0.0.1:32xxx -- e o Windows desta maquina recusa. Reproduzido
  fora da producao, com containers descartaveis em portas livres:

      1 publicacao com endereco fixo  ->  funciona
      2 ou mais                       ->  falha, sempre

  A porta que ele pede esta LIVRE (conferida uma a uma) e fica FORA da faixa
  dinamica do Windows, que aqui comeca em 49152. E desencontro entre o Docker
  Desktop e a configuracao de portas da maquina, nao erro de configuracao nossa.

  Como o app precisa de 80 E 443, qualquer versao da mudanca precisa de pelo
  menos dois binds. Por isso o script COMECA testando, com um container
  descartavel, se esta maquina aceita dois binds AGORA -- e so mexe na producao
  se aceitar. A aposta e que um reinicio limpe o estado do motor do Docker.

  USO

    .\servidor\restringir-443.ps1
        So confere. Nao toca em nada. Diz se a maquina aceita a mudanca hoje.

    .\servidor\restringir-443.ps1 -Aplicar
        Confere, e SO SE PASSAR aplica: escreve o compose, down, up, e verifica
        os tres enderecos. Se a verificacao falhar, volta o compose de antes e
        sobe de novo, sozinho.

    .\servidor\restringir-443.ps1 -Aplicar -Enderecos '193.168.0.200','127.0.0.1'
        O mesmo, sem o Wi-Fi -- e o que fazer se a maquina voltar de um reboot
        com o Wi-Fi desconectado (ja aconteceu).

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
Diz '2) ESTA MAQUINA ACEITA DOIS BINDS COM ENDERECO FIXO?'
Diz '   (container descartavel, em portas livres -- nao toca na producao)'
$nome = 'teste-bind-443'
$lixo = Join-Path $env:TEMP 'restringir-443-docker.txt'
& docker rm -f $nome 2>$lixo | Out-Null
$ip1 = $Enderecos[0]
& docker run -d --rm --name $nome -p "${ip1}:18080:80" -p "${ip1}:18443:80" nginx:alpine 2>$lixo | Out-Null
$passou = ($LASTEXITCODE -eq 0)
$motivo = if (Test-Path $lixo) { (Get-Content $lixo -Raw).Trim() } else { '' }
& docker rm -f $nome 2>$lixo | Out-Null
if (-not $passou) {
  Mau 'o Docker ainda recusa dois binds com endereco fixo'
  if ($motivo) { Diz ''; Diz "   $motivo" }
  Diz ''
  Diz 'PARADO, e nada foi tocado. O reinicio nao resolveu o estado do Docker.'
  Diz 'Deixe como esta: a disputa pela 443 nao esta acontecendo hoje (o Warsaw'
  Diz 'tem o endereco dele e funciona). Tentar de novo so depois de atualizar o'
  Diz 'Docker Desktop, ou de acertar a faixa de portas dinamicas do Windows.'
  exit 1
}
Bom 'dois binds com endereco fixo funcionam -- da para aplicar'

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
