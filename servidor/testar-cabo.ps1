<#
  TESTA A PLACA DO CABO COM O WI-FI DESLIGADO, SOZINHO, E RELIGA O WI-FI NO FIM.

  POR QUE EXISTE
  Em 31/08/2026 a placa do cabo parou de dar link: LED apagado, zero bytes, com
  dois cabos diferentes. O Windows nao acusa defeito nenhum nela (o controlador
  responde, o driver esta carregado) -- quem parece morto e o PHY, a metade que
  poe sinal eletrico no fio.

  Para testar sem o Wi-Fi no meio, a maquina precisa ficar sem internet -- e sem
  internet ninguem assiste ao teste de fora. Entao o teste ANOTA TUDO num
  arquivo e religa o Wi-Fi sozinho. Quem for ler, le depois.

  O QUE ELE FAZ, nesta ordem:

    1. registra uma tarefa de SEGURANCA que religa o Wi-Fi em 15 minutos,
       aconteca o que acontecer -- inclusive se este script morrer no meio ou a
       maquina travar. Sem isso, um erro aqui deixaria o servidor mudo, sem
       Wi-Fi e sem cabo, e alguem teria de ir ate ele;
    2. desliga o Wi-Fi;
    3. varre as velocidades: Auto, 1 Gbps, 100 Full, 100 Half, 10 Full, 10 Half.
       ESTE E O TESTE QUE FALTAVA: um cabo com so dois pares bons costuma
       FALHAR na autonegociacao e subir em 10 Mbps forcado. Se qualquer
       velocidade der link, a placa esta viva e o problema e o cabo;
    4. se algum link subir, tenta DHCP e fala com o roteador -- link sem
       trafego ainda pode ser cabo com um par ruim;
    5. religa o Wi-Fi, espera ele voltar de verdade e confirma a internet;
    6. apaga a tarefa de seguranca.

  COMO USAR
    Clique com o botao direito no PowerShell -> "Executar como administrador",
    e rode:

      cd C:\Users\Pichau\Desktop\Gerador-OS
      .\servidor\testar-cabo.ps1

    Leva uns 4 minutos. A internet cai durante o teste -- isso e esperado. No
    fim ele diz o veredito na tela E grava tudo em:

      servidor\tls\teste-cabo.log

  PARA DESFAZER, se algo ficar estranho:
    Enable-NetAdapter -Name 'Wi-Fi' -Confirm:$false
    Set-NetAdapterAdvancedProperty -Name 'Ethernet 3' -DisplayName 'Velocidade & Duplex' -DisplayValue 'Auto Negociação'
#>
[CmdletBinding()]
param(
  [string] $PlacaCabo = 'Ethernet 3',
  [string] $PlacaWifi = 'Wi-Fi',
  [int]    $SegundosPorVelocidade = 20,
  [int]    $MinutosDeSeguranca = 15,
  # Cria o atalho "Testar o cabo de rede" na Area de Trabalho e sai. Uma vez so.
  [switch] $CriarAtalho
)

$ErrorActionPreference = 'Continue'
$Log = Join-Path $PSScriptRoot 'tls\teste-cabo.log'
if (-not (Test-Path (Split-Path $Log))) { New-Item -ItemType Directory -Force -Path (Split-Path $Log) | Out-Null }

# ---- o atalho ---------------------------------------------------------------
# Digitar comando no PowerShell no meio de uma manha corrida e pedir erro -- e
# este teste tira a maquina da rede, entao nao e hora de errar a digitacao. O
# atalho ja pede a permissao de administrador sozinho, e deixa a janela ABERTA
# no fim (-NoExit): o veredito tem de poder ser lido antes de a janela sumir.
if ($CriarAtalho) {
  $ps1    = Join-Path $PSScriptRoot 'testar-cabo.ps1'
  $atalho = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Testar o cabo de rede.lnk'
  try {
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut($atalho)
    $s.TargetPath       = 'powershell.exe'
    $s.Arguments        = '-NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList ''-NoProfile'',''-ExecutionPolicy'',''Bypass'',''-NoExit'',''-File'',''' + $ps1 + '''"'
    $s.WorkingDirectory = (Split-Path -Parent $PSScriptRoot)
    $s.IconLocation     = 'shell32.dll,18'
    $s.Description      = 'Testa a placa do cabo com o Wi-Fi desligado e religa o Wi-Fi no fim.'
    $s.Save()
    Write-Host ""
    Write-Host "Atalho criado na Area de Trabalho: Testar o cabo de rede" -ForegroundColor Green
    Write-Host "Clique duas vezes nele e confirme a permissao do Windows." -ForegroundColor Green
    Write-Host ""
  } catch {
    Write-Host "nao consegui criar o atalho: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
  }
  exit 0
}

function Anotar($t) {
  $linha = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $t
  Add-Content -Path $Log -Value $linha -Encoding utf8
  Write-Host $linha
}
function Titulo($t) {
  Add-Content -Path $Log -Value "" -Encoding utf8
  Anotar "===== $t ====="
}

$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Host "Este teste precisa ser rodado como ADMINISTRADOR." -ForegroundColor Red
  Write-Host "Botao direito no PowerShell -> Executar como administrador." -ForegroundColor Red
  exit 1
}

Add-Content -Path $Log -Value "`r`n`r`n########## TESTE DE $(Get-Date -Format 'dd/MM/yyyy HH:mm') ##########" -Encoding utf8

# ---- 1. a rede de seguranca -----------------------------------------------
# Isto vem ANTES de qualquer coisa. Se o script morrer entre desligar e religar
# o Wi-Fi, esta tarefa religa sozinha e ninguem precisa ir ate a maquina.
$TAREFA = 'Gerador-OS Religar WiFi (seguranca)'
Titulo "Rede de seguranca"
try {
  Unregister-ScheduledTask -TaskName $TAREFA -Confirm:$false -ErrorAction SilentlyContinue
  $acao = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -Command `"Enable-NetAdapter -Name '$PlacaWifi' -Confirm:`$false`""
  $quando = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes($MinutosDeSeguranca)
  Register-ScheduledTask -TaskName $TAREFA -Action $acao -Trigger $quando -RunLevel Highest `
    -Description 'Religa o Wi-Fi se o teste de cabo nao chegar ao fim.' -ErrorAction Stop | Out-Null
  Anotar "ok - o Wi-Fi religa sozinho as $((Get-Date).AddMinutes($MinutosDeSeguranca).ToString('HH:mm')) se este teste falhar"
} catch {
  Anotar "NAO consegui criar a rede de seguranca: $($_.Exception.Message)"
  Anotar "ABORTANDO - sem ela, uma falha aqui deixaria o servidor sem rede nenhuma."
  exit 1
}

# ---- 2. como tudo esta antes ----------------------------------------------
Titulo "Antes do teste"
Get-NetAdapter -Physical | Select-Object Name, InterfaceDescription, Status, MediaConnectionState, LinkSpeed |
  Format-Table -AutoSize | Out-String -Width 140 | ForEach-Object { Add-Content -Path $Log -Value $_ -Encoding utf8; Write-Host $_ }

$velocidadeOriginal = (Get-NetAdapterAdvancedProperty -Name $PlacaCabo -DisplayName 'Velocidade & Duplex' -ErrorAction SilentlyContinue).DisplayValue
Anotar "velocidade configurada antes: $velocidadeOriginal"

# ---- 3. desliga o Wi-Fi ----------------------------------------------------
Titulo "Desligando o Wi-Fi"
try {
  Disable-NetAdapter -Name $PlacaWifi -Confirm:$false -ErrorAction Stop
  Start-Sleep -Seconds 3
  Anotar "ok - Wi-Fi desligado. A partir daqui a maquina esta sem internet."
} catch {
  Anotar "nao consegui desligar o Wi-Fi: $($_.Exception.Message)"
}

# ---- 4. a varredura de velocidades ----------------------------------------
# O TESTE QUE FALTAVA. Um cabo com dois pares bons de quatro nao fecha
# autonegociacao (que exige os quatro para Gigabit e negocia em cima disso),
# mas 10 Mbps usa so um par por sentido. Se subir em 10 e nao em Auto, o
# diagnostico esta fechado: a placa esta viva e o cabo e que esta partido.
$velocidades = @('Auto Negociação','1.0 Gbps Full Duplex','100 Mbps Full Duplex',
                 '100 Mbps Half Duplex','10 Mbps Full Duplex','10 Mbps Half Duplex')
$conseguiu = $null

foreach ($v in $velocidades) {
  Titulo "Tentando: $v"
  try {
    Set-NetAdapterAdvancedProperty -Name $PlacaCabo -DisplayName 'Velocidade & Duplex' -DisplayValue $v -ErrorAction Stop
  } catch {
    Anotar "esta placa nao aceita '$v' - pulando"
    continue
  }
  # Reiniciar forca o PHY a renegociar do zero, em vez de herdar o estado morto.
  Restart-NetAdapter -Name $PlacaCabo -Confirm:$false -ErrorAction SilentlyContinue

  $subiu = $false
  $passos = [Math]::Max(4, [int]($SegundosPorVelocidade / 5))
  foreach ($i in 1..$passos) {
    Start-Sleep -Seconds 5
    $n = Get-NetAdapter -Name $PlacaCabo -ErrorAction SilentlyContinue
    Anotar ("  {0,3}s  {1,-13} {2,-13} {3}" -f ($i*5), $n.Status, $n.MediaConnectionState, $n.LinkSpeed)
    if ($n.MediaConnectionState -eq 'Connected') { $subiu = $true; break }
  }
  if ($subiu) { $conseguiu = $v; Anotar "  *** LINK! em '$v' ***"; break }
}

# ---- 5. se subiu, o cabo carrega trafego de verdade? ----------------------
if ($conseguiu) {
  Titulo "Link de pe - testando se PASSA TRAFEGO"
  # Link e so o pulso eletrico. Um cabo com um par ruim pode dar link e perder
  # pacote - por isso a conta so fecha com trafego de verdade.
  & netsh @('interface','ipv4','set','address',"name=$PlacaCabo",'source=dhcp') | Out-Null
  ipconfig /renew "$PlacaCabo" | Out-Null
  Start-Sleep -Seconds 8
  $ip = Get-NetIPAddress -InterfaceAlias $PlacaCabo -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.PrefixOrigin -eq 'Dhcp' }
  if ($ip) {
    Anotar "endereco do DHCP: $($ip.IPAddress)"
    $gw = (Get-NetIPConfiguration -InterfaceAlias $PlacaCabo -ErrorAction SilentlyContinue).IPv4DefaultGateway.NextHop
    if ($gw) {
      $ping = Test-Connection -ComputerName $gw -Count 8 -ErrorAction SilentlyContinue
      $perdidos = 8 - (($ping | Measure-Object).Count)
      Anotar "roteador $gw : $(8 - $perdidos)/8 respostas ($perdidos perdidos)"
      if ($perdidos -eq 0) { Anotar "VEREDITO PARCIAL: o cabo passa trafego limpo nesta velocidade." }
      elseif ($perdidos -lt 8) { Anotar "VEREDITO PARCIAL: passa, mas PERDE PACOTE - cabo ruim, trocar." }
      else { Anotar "VEREDITO PARCIAL: link sem trafego - cabo ruim ou porta do switch." }
    } else { Anotar "sem gateway - o DHCP respondeu pela metade" }
  } else {
    Anotar "link de pe mas o DHCP NAO respondeu - cabo passa pulso e nao passa dado."
  }
  Get-NetAdapterStatistics -Name $PlacaCabo | Select-Object ReceivedBytes, SentBytes |
    Format-List | Out-String | ForEach-Object { Add-Content -Path $Log -Value $_ -Encoding utf8; Write-Host $_ }
} else {
  Titulo "Nenhuma velocidade deu link"
  Anotar "Nem 10 Mbps Half Duplex, que usa UM par de fios e e o piso do padrao."
  # Volta para Auto: e o certo para o dia a dia, e para a placa que vier depois.
  Set-NetAdapterAdvancedProperty -Name $PlacaCabo -DisplayName 'Velocidade & Duplex' `
    -DisplayValue 'Auto Negociação' -ErrorAction SilentlyContinue
}

# ---- 6. religa o Wi-Fi -----------------------------------------------------
Titulo "Religando o Wi-Fi"
Enable-NetAdapter -Name $PlacaWifi -Confirm:$false -ErrorAction SilentlyContinue
$voltou = $false
foreach ($i in 1..24) {
  Start-Sleep -Seconds 5
  $w = Get-NetAdapter -Name $PlacaWifi -ErrorAction SilentlyContinue
  if ($w -and $w.MediaConnectionState -eq 'Connected') {
    $end = (Get-NetIPAddress -InterfaceAlias $PlacaWifi -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.AddressState -eq 'Preferred' }).IPAddress
    if ($end) { Anotar "ok - Wi-Fi de volta em $end (levou $($i*5)s)"; $voltou = $true; break }
  }
}
if (-not $voltou) {
  Anotar "ATENCAO: o Wi-Fi nao voltou em 2 minutos. Rode a mao:"
  Anotar "   Enable-NetAdapter -Name '$PlacaWifi' -Confirm:`$false"
} else {
  try { Resolve-DnsName 'github.com' -ErrorAction Stop | Out-Null; Anotar "internet de pe (DNS resolvendo)" }
  catch { Anotar "Wi-Fi conectado mas SEM internet - o DNS nao resolveu" }
}

# ---- 7. veredito e limpeza -------------------------------------------------
Titulo "VEREDITO"
if ($conseguiu) {
  Anotar "A PLACA ESTA VIVA. Deu link em '$conseguiu'."
  if ($conseguiu -ne 'Auto Negociação') {
    Anotar "Como so subiu FORCANDO a velocidade, o cabo tem par danificado:"
    Anotar "a autonegociacao precisa dos quatro pares e nao fecha. TROCAR O CABO."
  }
} else {
  Anotar "NENHUMA velocidade deu link, com o Wi-Fi desligado e a placa reiniciada"
  Anotar "a cada tentativa. Somado ao LED apagado e aos dois cabos ja testados,"
  Anotar "a placa (o PHY) esta morta. Saida: adaptador USB-Ethernet no servidor."
  Anotar "Depois de plugar: .\servidor\fixar-ip-cabo.ps1 (acha a placa nova sozinho)."
}

if ($voltou) {
  Unregister-ScheduledTask -TaskName $TAREFA -Confirm:$false -ErrorAction SilentlyContinue
  Anotar "rede de seguranca removida (o Wi-Fi ja voltou)"
} else {
  Anotar "rede de seguranca MANTIDA - ela religa o Wi-Fi sozinha em instantes"
}

Write-Host ""
Write-Host "Terminou. O relatorio inteiro esta em:" -ForegroundColor Green
Write-Host "  $Log" -ForegroundColor Green
