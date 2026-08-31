# Fixa o IP do Wi-Fi desta maquina em 192.168.1.158.
#
# Por que: enquanto o cabo de rede esta fora, o servidor so atende pelo Wi-Fi,
# e o endereco vinha do DHCP do roteador -- podia mudar sozinho a cada
# renovacao. O 192.168.1.158 ja esta gravado no SAN do certificado e na linha
# 39 do instalar-certificado.cmd, entao fixar ESTE numero nao obriga a
# reemitir nada.
#
# Rode como ADMINISTRADOR. Se qualquer passo falhar, o script devolve a placa
# para DHCP sozinho -- nao deixa a maquina fora da rede.
#
# Para desfazer a qualquer momento:
#   Set-NetIPInterface -InterfaceAlias 'Wi-Fi' -Dhcp Enabled
#   Set-DnsClientServerAddress -InterfaceAlias 'Wi-Fi' -ResetServerAddresses
#   ipconfig /renew

$ErrorActionPreference = 'Stop'

$PLACA    = 'Wi-Fi'
$IP       = '192.168.1.158'
$PREFIXO  = 24
$GATEWAY  = '192.168.1.1'
$DNS      = '192.168.1.1'

function Passo($t) { Write-Host "`n>> $t" -ForegroundColor Cyan }
function Ok($t)    { Write-Host "   OK - $t" -ForegroundColor Green }
function Aviso($t) { Write-Host "   [AVISO] $t" -ForegroundColor Yellow }

# ---- 0. sou administrador? ------------------------------------------------
$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Host "Este script precisa ser rodado como ADMINISTRADOR." -ForegroundColor Red
  Write-Host "Clique com o botao direito no PowerShell e escolha 'Executar como administrador'." -ForegroundColor Red
  exit 1
}

function VoltarParaDhcp {
  Aviso "Devolvendo a placa para DHCP..."
  try {
    Set-NetIPInterface -InterfaceAlias $PLACA -Dhcp Enabled -ErrorAction SilentlyContinue
    Set-DnsClientServerAddress -InterfaceAlias $PLACA -ResetServerAddresses -ErrorAction SilentlyContinue
    ipconfig /renew "$PLACA" | Out-Null
    Aviso "Placa devolvida para DHCP. A maquina deve voltar a rede em alguns segundos."
  } catch {
    Write-Host "   NAO consegui devolver para DHCP. Rode a mao:" -ForegroundColor Red
    Write-Host "   Set-NetIPInterface -InterfaceAlias '$PLACA' -Dhcp Enabled" -ForegroundColor Red
  }
}

# ---- 1. como esta agora (para conferencia) --------------------------------
Passo "Como a placa '$PLACA' esta agora"
Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 |
  Select-Object IPAddress, PrefixLength, PrefixOrigin | Format-Table -AutoSize | Out-String | Write-Host

# ---- 2. o endereco esta livre? --------------------------------------------
# Se outro aparelho ja responde no .158, fixar aqui criaria conflito de IP.
# Melhor descobrir antes de mexer do que depois.
Passo "Conferindo se $IP esta livre"
$eu = (Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
if ($eu -contains $IP) {
  Ok "$IP ja e desta maquina (concessao atual do DHCP) - livre por definicao."
} else {
  if (Test-Connection -ComputerName $IP -Count 2 -Quiet -ErrorAction SilentlyContinue) {
    Write-Host "   $IP JA ESTA EM USO por outro aparelho. Abortando." -ForegroundColor Red
    Write-Host "   Nada foi alterado." -ForegroundColor Red
    exit 1
  }
  Ok "$IP nao respondeu - livre."
}

# ---- 3. fixar ---------------------------------------------------------------
Passo "Fixando $IP/$PREFIXO, gateway $GATEWAY, DNS $DNS"
try {
  Set-NetIPInterface -InterfaceAlias $PLACA -Dhcp Disabled

  # Tira os enderecos IPv4 que vieram do DHCP antes de por o fixo, senao a
  # placa fica com os dois e o Windows escolhe um deles na hora de responder.
  Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
    Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

  Get-NetRoute -InterfaceAlias $PLACA -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
    Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue

  New-NetIPAddress -InterfaceAlias $PLACA -IPAddress $IP -PrefixLength $PREFIXO -DefaultGateway $GATEWAY | Out-Null
  Set-DnsClientServerAddress -InterfaceAlias $PLACA -ServerAddresses $DNS
  Ok "endereco fixado."
} catch {
  Write-Host "   FALHOU: $($_.Exception.Message)" -ForegroundColor Red
  VoltarParaDhcp
  exit 1
}

Start-Sleep -Seconds 4

# ---- 4. conferir que a rede continua de pe --------------------------------
# Fixar IP errado tira a maquina do ar. Nao adianta dizer "pronto" sem provar
# que o gateway responde e que o app abre.
Passo "Conferindo"
$falhas = @()

$origem = (Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
           Where-Object { $_.IPAddress -eq $IP }).PrefixOrigin
if ($origem -eq 'Manual') { Ok "a placa esta com $IP, origem Manual (fixo)." }
else { $falhas += "o endereco nao ficou como Manual (origem: $origem)" }

if (Test-Connection -ComputerName $GATEWAY -Count 2 -Quiet -ErrorAction SilentlyContinue) {
  Ok "roteador ($GATEWAY) responde."
} else { $falhas += "o roteador $GATEWAY nao responde" }

try {
  Resolve-DnsName 'github.com' -ErrorAction Stop | Out-Null
  Ok "DNS resolvendo (internet de pe - backup e espelho seguem funcionando)."
} catch { $falhas += "o DNS nao resolveu - a maquina esta sem internet" }

foreach ($u in @("https://localhost", "https://$IP")) {
  try {
    $r = Invoke-WebRequest -Uri $u -TimeoutSec 10 -UseBasicParsing
    Ok "$u responde (HTTP $($r.StatusCode))."
  } catch { $falhas += "$u nao respondeu" }
}

# ---- 5. veredito ------------------------------------------------------------
Write-Host ""
if ($falhas.Count -eq 0) {
  Write-Host "PRONTO. O Wi-Fi desta maquina agora e sempre $IP." -ForegroundColor Green
  Write-Host "O endereco nao muda mais sozinho, e o certificado ja o cobre." -ForegroundColor Green
} else {
  Write-Host "DEU ERRADO:" -ForegroundColor Red
  $falhas | ForEach-Object { Write-Host "   - $_" -ForegroundColor Red }
  VoltarParaDhcp
  exit 1
}
