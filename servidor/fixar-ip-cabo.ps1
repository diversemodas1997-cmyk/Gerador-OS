# Fixa o IP do CABO desta maquina em 193.168.0.200.
#
# Por que: o .200 nunca foi fixo de verdade. A placa estava em DHCP e o .200
# era uma concessao que se renovava enquanto a maquina ficava ligada. O cabo
# saiu na sexta 28/08, a concessao venceu no fim de semana, e na segunda 31/08
# o roteador entregou 193.168.0.9. Toda a fabrica aponta para o .200 -- atalho,
# certificado, ca.crt ja instalado -- entao a fabrica ficou sem o app.
#
# Fixar ESTE numero nao obriga a reemitir nem a reinstalar nada: o .200 ja esta
# no SAN do certificado (junto com 192.168.1.158, 127.0.0.1 e localhost).
#
# Rode como ADMINISTRADOR. Se qualquer conferencia falhar, o script devolve a
# placa para DHCP sozinho -- nao deixa o servidor fora da rede.
#
# Para desfazer a qualquer momento:
#   netsh interface ipv4 set address name="Ethernet 3" source=dhcp
#   netsh interface ipv4 set dnsservers name="Ethernet 3" source=dhcp
#
# DEPOIS DISTO, falta um passo que nao e daqui: reservar o .200 por MAC no
# roteador (MAC B4-2E-99-F4-B6-8B), para o roteador nunca entregar este
# endereco a outro aparelho enquanto o servidor estiver desligado.

param(
  # A placa do cabo. Vazio = descobre sozinha qual placa com fio esta com link.
  #
  # Por que virou parametro em 31/08/2026: com o cabo sem dar link nem em
  # autonegociacao, a saida mais provavel passou a ser um adaptador
  # USB-Ethernet -- e ele entra no Windows com OUTRO nome ('Ethernet 4',
  # 'Realtek USB GbE...'). Sem isto, o script continuaria procurando uma placa
  # que talvez nunca mais tenha link, e diria "cabo sem link" com o cabo
  # funcionando na placa nova, ao lado.
  [string] $Placa
)

$ErrorActionPreference = 'Continue'

# Descobre a placa COM FIO que esta com link. Exclui Wi-Fi (que tem casa
# propria, no fixar-ip-wifi.ps1) e as virtuais do Docker/Hyper-V, que estao
# sempre "conectadas" e roubariam a escolha.
function AcharPlacaDoCabo {
  $c = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
       Where-Object { $_.MediaConnectionState -eq 'Connected' -and
                      # LOOPBACK FICA DE FORA, e isto custou caro em 31/08/2026:
                      # o 'Microsoft KM-TEST Loopback Adapter' (o do Audaces) e
                      # reportado como FISICO, aparece 'Connected' a 1,2 Gbps e
                      # por isso ganhou a ordenacao por velocidade. O .200 foi
                      # parar nele -- que nao esta em fio nenhum --, o loopback
                      # perdeu o proprio endereco (54.232.189.113, de onde a
                      # licenca do CAD depende), e todas as conferencias
                      # PASSARAM: daqui de dentro respondia tudo. A fabrica e
                      # que nao alcancaria.
                      $_.InterfaceDescription -notmatch 'Wireless|Wi-Fi|802\.11|Loopback|Virtual|TAP|VPN' -and
                      $_.Name -notmatch '^(vEthernet|Wi-Fi|Topaz|Loopback)' } |
       # Ordena por velocidade so para desempatar entre placas de verdade.
       Sort-Object -Property @{ E = { $_.LinkSpeed } } -Descending |
       Select-Object -First 1
  if ($c) { return $c.Name }
  # Nenhuma com link: fica com a de sempre, para as mensagens fazerem sentido.
  return 'Ethernet 3'
}

$PLACA    = if ($Placa) { $Placa } else { AcharPlacaDoCabo }
$IP       = '193.168.0.200'
$MASCARA  = '255.255.255.0'
$GATEWAY  = '193.168.0.1'
$DNS      = @('132.255.216.26','132.255.216.27')   # os que o DHCP ja entregava

function Passo($t) { Write-Host "`n>> $t" -ForegroundColor Cyan }
function Ok($t)    { Write-Host "   OK - $t" -ForegroundColor Green }
function Aviso($t) { Write-Host "   [AVISO] $t" -ForegroundColor Yellow }
function Erro($t)  { Write-Host "   $t" -ForegroundColor Red }

# ---- 0. sou administrador? ------------------------------------------------
$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Erro "Este script precisa ser rodado como ADMINISTRADOR."
  exit 1
}

function VoltarParaDhcp {
  Aviso "Devolvendo a placa para DHCP..."
  & netsh @('interface','ipv4','set','address',"name=$PLACA",'source=dhcp')    | Out-Null
  & netsh @('interface','ipv4','set','dnsservers',"name=$PLACA",'source=dhcp') | Out-Null
  # O DHCP nao entrega o endereco na hora. Sem esperar, a mensagem final mente:
  # diz "voltou" enquanto a placa ainda esta em 169.254 -- que e ficar sem rede.
  $voltou = $false
  foreach ($i in 1..12) {
    Start-Sleep -Seconds 5
    $a = (Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
          Where-Object { $_.PrefixOrigin -eq 'Dhcp' }).IPAddress
    if ($a) { Aviso "placa de volta no DHCP, endereco $a."; $voltou = $true; break }
  }
  if (-not $voltou) {
    Erro "A PLACA NAO PEGOU ENDERECO DO DHCP. O servidor esta SEM REDE."
    Erro "Tente a mao:  ipconfig /release   e depois   ipconfig /renew"
  }
}

# ---- 1. como esta agora (para conferencia) --------------------------------
Write-Host ""
Write-Host "Placa do cabo: $PLACA$(if (-not $Placa) { '  (descoberta automaticamente)' })" -ForegroundColor Cyan

Passo "Como a placa '$PLACA' esta agora"
Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
  Select-Object IPAddress, PrefixLength, PrefixOrigin, AddressState |
  Format-Table -AutoSize | Out-String | Write-Host

# ---- 2. o endereco esta livre? --------------------------------------------
# Se outro aparelho ja responde no .200, fixar aqui criaria conflito de IP e
# derrubaria os dois. Melhor descobrir antes de mexer do que depois.
Passo "Conferindo se $IP esta livre"
$meus = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
if ($meus -contains $IP) {
  Ok "$IP ja e desta maquina - livre por definicao."
} elseif (Test-Connection -ComputerName $IP -Count 2 -Quiet -ErrorAction SilentlyContinue) {
  Erro "$IP JA ESTA EM USO por outro aparelho. Abortando - nada foi alterado."
  exit 1
} else {
  Ok "$IP nao respondeu - livre."
}

# ---- 3. fixar ---------------------------------------------------------------
Passo "Fixando $IP / $MASCARA, gateway $GATEWAY, DNS $($DNS -join ', ')"

# Por que netsh e nao New-NetIPAddress: na primeira tentativa (31/08) o cmdlet
# devolveu "Inconsistent parameters PolicyStore PersistentStore and Dhcp
# Enabled" (erro 87). Set-NetIPInterface -Dhcp Disabled nao pega enquanto a
# placa ainda segura o endereco que o DHCP deu, e o cmdlet entrou pela METADE:
# gravou o endereco e nao gravou o gateway. O servidor ficou com IP e sem rota,
# que e pior do que nao ter mexido. O netsh faz endereco, mascara e gateway
# numa operacao so, e desliga o DHCP junto.
$saida = (& netsh @('interface','ipv4','set','address',"name=$PLACA",'static',$IP,$MASCARA,$GATEWAY) | Out-String).Trim()
if ($saida) {
  Erro "FALHOU ao fixar o endereco: $saida"
  VoltarParaDhcp
  exit 1
}
& netsh @('interface','ipv4','set','dnsservers',"name=$PLACA",'static',$DNS[0],'primary','validate=no') | Out-Null
& netsh @('interface','ipv4','add','dnsservers',"name=$PLACA",$DNS[1],'index=2')                        | Out-Null
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

# ---- 4. conferir que o servidor continua de pe ----------------------------
# Fixar IP errado tira o servidor do ar. Nao adianta dizer "pronto" sem provar
# que o gateway responde, que o DNS resolve (backup e espelho dependem de
# internet) e que o app abre COM O CERTIFICADO ACEITO no endereco novo -- que
# e exatamente o que a maquina da fabrica vai fazer.
Passo "Conferindo"
$falhas = @()

$e = Get-NetIPAddress -InterfaceAlias $PLACA -AddressFamily IPv4 -ErrorAction SilentlyContinue |
     Where-Object { $_.IPAddress -eq $IP }
if ($e -and $e.PrefixOrigin -eq 'Manual') { Ok "a placa esta com $IP, origem Manual (fixo)." }
else { $falhas += "o endereco nao ficou como Manual (origem: $($e.PrefixOrigin))" }

$rota = Get-NetRoute -InterfaceAlias $PLACA -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue
if ($rota) { Ok "rota padrao gravada (saida por $($rota[0].NextHop))." }
else { $falhas += "a rota padrao nao foi gravada - foi isto que falhou na 1a tentativa" }

if (Test-Connection -ComputerName $GATEWAY -Count 3 -Quiet -ErrorAction SilentlyContinue) {
  Ok "roteador ($GATEWAY) responde."
} else { $falhas += "o roteador $GATEWAY nao responde" }

try {
  Resolve-DnsName 'github.com' -ErrorAction Stop | Out-Null
  Ok "DNS resolvendo (internet de pe - backup e espelho seguem funcionando)."
} catch { $falhas += "o DNS nao resolveu - o servidor esta sem internet" }

$ProgressPreference = 'SilentlyContinue'
foreach ($u in @("https://localhost", "https://$IP")) {
  try {
    $r = Invoke-WebRequest -Uri $u -TimeoutSec 15 -UseBasicParsing
    Ok "$u responde com certificado ACEITO (HTTP $($r.StatusCode))."
  } catch { $falhas += "$u nao respondeu: $($_.Exception.Message)" }
}

# ---- 5. veredito ------------------------------------------------------------
Write-Host ""
if ($falhas.Count -eq 0) {
  Write-Host "PRONTO. O servidor volta a ser sempre $IP." -ForegroundColor Green
  Write-Host "Os atalhos da fabrica voltam a funcionar sem tocar em nenhuma maquina." -ForegroundColor Green
  Write-Host ""
  Write-Host "FALTA UM PASSO, no roteador: reservar $IP para o MAC B4-2E-99-F4-B6-8B," -ForegroundColor Yellow
  Write-Host "para ele nunca entregar este endereco a outro aparelho." -ForegroundColor Yellow
} else {
  Write-Host "DEU ERRADO:" -ForegroundColor Red
  $falhas | ForEach-Object { Erro "- $_" }
  VoltarParaDhcp
  exit 1
}
