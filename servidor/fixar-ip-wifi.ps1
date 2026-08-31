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
$MASCARA  = '255.255.255.0'
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
    & netsh @('interface','ipv4','set','address',"name=$PLACA",'source=dhcp')    | Out-Null
    & netsh @('interface','ipv4','set','dnsservers',"name=$PLACA",'source=dhcp') | Out-Null

    # O DHCP nao entrega o endereco na hora. Sem esperar, a mensagem final
    # mente: diz "voltou" enquanto a placa ainda esta em 169.254 -- que e ficar
    # sem rede. Foi o que aconteceu no cabo em 31/08.
    $voltou = $false
    foreach ($i in 1..12) {
      Start-Sleep -Seconds 5
      $a = (Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.PrefixOrigin -eq 'Dhcp' }).IPAddress
      if ($a) { Aviso "placa de volta no DHCP, endereco $a."; $voltou = $true; break }
    }
    if (-not $voltou) {
      Write-Host "   A PLACA NAO PEGOU ENDERECO DO DHCP. A maquina esta SEM REDE." -ForegroundColor Red
      Write-Host "   Tente a mao:  ipconfig /release   e depois   ipconfig /renew" -ForegroundColor Red
    }
  } catch {
    Write-Host "   NAO consegui devolver para DHCP. Rode a mao:" -ForegroundColor Red
    Write-Host "   netsh interface ipv4 set address name=`"$PLACA`" source=dhcp" -ForegroundColor Red
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

# Por que netsh e nao New-NetIPAddress: em 31/08, no cabo, o cmdlet devolveu
# "Inconsistent parameters PolicyStore PersistentStore and Dhcp Enabled" (erro
# 87) -- o `Set-NetIPInterface -Dhcp Disabled` acima nao pega enquanto a placa
# ainda segura o endereco que o DHCP deu. E o cmdlet entrou pela METADE: gravou
# o endereco e nao gravou o gateway. Aqui isso seria pior do que no cabo: sem
# gateway no Wi-Fi, esta maquina fica sem internet. O netsh faz endereco,
# mascara e gateway numa operacao so, e desliga o DHCP junto.
$saida = (& netsh @('interface','ipv4','set','address',"name=$PLACA",'static',$IP,$MASCARA,$GATEWAY) | Out-String).Trim()
if ($saida) {
  Write-Host "   FALHOU ao fixar o endereco: $saida" -ForegroundColor Red
  VoltarParaDhcp
  exit 1
}
& netsh @('interface','ipv4','set','dnsservers',"name=$PLACA",'static',$DNS,'primary','validate=no') | Out-Null
Ok "endereco fixado."

# O Windows testa o endereco na rede antes de usa-lo (deteccao de duplicado).
# Conferir enquanto ele esta 'Tentative' da falso negativo.
Passo "Esperando o endereco ficar valido"
$pronto = $false
foreach ($i in 1..10) {
  Start-Sleep -Seconds 3
  $e = Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
       Where-Object { $_.IPAddress -eq $IP }
  if ($e -and $e.AddressState -eq 'Preferred') { Ok "endereco valido (Preferred)."; $pronto = $true; break }
}
if (-not $pronto) { Aviso "o endereco nao chegou a 'Preferred' - as conferencias abaixo dirao se serve." }

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
