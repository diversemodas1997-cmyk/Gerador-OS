<#
  Mantem a rede do servidor pronta para atender, pelo cabo ou pelo Wi-Fi.

  POR QUE ISTO EXISTE
  Em 31/08/2026 a fabrica abriu a segunda sem o app, com o servidor 100% de
  pe. Foram tres defeitos empilhados, e nenhum deles era do programa:

    1. o .200 vinha do DHCP e a concessao venceu no fim de semana;
    2. o cabo caiu e o servidor passou a atender so pelo Wi-Fi;
    3. o Windows classificou a rede nova como PUBLICA, e o perfil Publico
       ignora TODA regra de entrada -- entao nenhuma maquina chegava.

  O terceiro e o pior, porque nao ha o que ler: a regra do firewall aparece
  verde e correta, nao ha erro nem log, a conexao so nao chega. E conferir do
  proprio servidor SEMPRE da certo, porque o laco local nao passa pelo
  firewall. Sem este vigia, o defeito volta toda vez que a maquina trocar de
  rede -- e TODA rede nova entra como Publica.

  O QUE ELE GARANTE, a cada passagem:

    a) toda rede conectada esta como PARTICULAR (senao as regras nao valem);
    b) as portas 80 e 443 tem regra de entrada;
    c) o cabo, quando tem link, esta no 193.168.0.200 fixo.

  O (a) e o (b) nao tem risco: nao mexem em endereco nem derrubam conexao. O
  (c) mexe, e por isso e delegado ao `fixar-ip-cabo.ps1`, que confere se o
  endereco esta livre antes, prova rota/gateway/DNS/app depois, e devolve a
  placa para DHCP sozinho se qualquer conferencia falhar.

  O Wi-Fi NAO e mexido de proposito. E por ele que esta maquina fala com a
  internet -- backup, espelho e o proprio atendimento remoto -- e trocar o
  endereco dele as cegas, sem ninguem olhando, e arriscar o que nao precisa
  ser arriscado. Ele ja esta fixo; se sair do lugar, o log avisa.

  COMO USAR
    .\servidor\vigia-rede.ps1              conferir e consertar agora
    .\servidor\vigia-rede.ps1 -Somente     so conferir, sem mexer em nada
    .\servidor\vigia-rede.ps1 -Agendar     registrar no logon + a cada 15 min

  Escreve em servidor\tls\vigia-rede.log, e SO quando ha o que contar --
  um log que repete "tudo bem" a cada 15 minutos nao e lido por ninguem.
#>
[CmdletBinding()]
param(
  [switch] $Somente,
  [switch] $Agendar,
  [int]    $CadaMinutos = 15
)

$ErrorActionPreference = 'Continue'

$Raiz    = Split-Path -Parent $PSScriptRoot
$LogDir  = Join-Path $PSScriptRoot 'tls'
$Log     = Join-Path $LogDir 'vigia-rede.log'
$PLACA_CABO = 'Ethernet 3'
$IP_CABO    = '193.168.0.200'

if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }

$houve = $false
function Anotar($t) {
  $script:houve = $true
  $linha = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $t
  Add-Content -Path $Log -Value $linha -Encoding utf8
  Write-Host $linha
}
function Dizer($t) { Write-Host "   $t" }

# ---- agendar --------------------------------------------------------------
if ($Agendar) {
  $nome = 'Gerador-OS Vigia Rede'
  # Pelo wscript, como o vigia do Docker: sem janela preta piscando a cada
  # 15 minutos na cara de quem estiver usando a maquina.
  $vbs = Join-Path $LogDir 'vigia-rede-oculto.vbs'
  $ps  = Join-Path $PSScriptRoot 'vigia-rede.ps1'
  Set-Content -Path $vbs -Encoding ASCII -Value @"
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -File ""$ps""", 0, False
"@
  $acao = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $vbs + '"') -WorkingDirectory $Raiz
  $noLogon   = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
  $repetindo = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(2) `
                 -RepetitionInterval (New-TimeSpan -Minutes $CadaMinutos) `
                 -RepetitionDuration (New-TimeSpan -Days 3650)
  $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
  try {
    Unregister-ScheduledTask -TaskName $nome -Confirm:$false -ErrorAction SilentlyContinue
    # -RunLevel Highest: sem isto a tarefa roda sem administrador, e tudo o
    # que este script conserta exige administrador. Falharia em silencio.
    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger @($noLogon, $repetindo) `
      -Settings $conf -RunLevel Highest `
      -Description 'Mantem as redes do servidor como Particular, as portas abertas e o cabo no IP fixo.' `
      -ErrorAction Stop | Out-Null
    Anotar "tarefa '$nome' registrada: no logon e a cada $CadaMinutos min"
  } catch {
    Anotar "FALHA ao registrar a tarefa: $($_.Exception.Message)"
    exit 1
  }
  exit 0
}

# ---- preciso de administrador para consertar ------------------------------
$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$admin = $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin -and -not $Somente) {
  Write-Host "Sem administrador nao da para consertar nada. Rodando so a conferencia." -ForegroundColor Yellow
  $Somente = $true
}

Write-Host ""
Write-Host "Vigia de rede - $(Get-Date -Format 'dd/MM/yyyy HH:mm')" -ForegroundColor Cyan

# ---- a) toda rede conectada tem de ser Particular -------------------------
# Este e o passo que ninguem lembra. Uma rede Publica faz o Windows ignorar
# as regras de entrada, e o sintoma e "abre no servidor e nao abre em mais
# lugar nenhum" -- sem uma linha de erro em lugar algum.
foreach ($perfil in (Get-NetConnectionProfile -ErrorAction SilentlyContinue)) {
  # As placas virtuais do Docker/Hyper-V e o loopback nao servem ninguem de
  # fora; mexer nelas seria barulho sem beneficio.
  if ($perfil.InterfaceAlias -match '^(vEthernet|Loopback|Topaz)') { continue }

  if ($perfil.NetworkCategory -eq 'Public') {
    if ($Somente) {
      Anotar "a rede '$($perfil.Name)' ($($perfil.InterfaceAlias)) esta PUBLICA - o firewall ignora as regras de entrada"
    } else {
      try {
        Set-NetConnectionProfile -InterfaceAlias $perfil.InterfaceAlias -NetworkCategory Private -ErrorAction Stop
        Anotar "rede '$($perfil.Name)' ($($perfil.InterfaceAlias)) estava PUBLICA - passou para Particular"
      } catch {
        Anotar "FALHA ao reclassificar '$($perfil.Name)': $($_.Exception.Message)"
      }
    }
  } else {
    Dizer "rede '$($perfil.Name)' ($($perfil.InterfaceAlias)): $($perfil.NetworkCategory) - ok"
  }
}

# ---- b) as portas tem regra de entrada ------------------------------------
# `Get-NetFirewallPortFilter` exige ADMINISTRADOR: sem ele a consulta volta
# vazia, e "vazia" e indistinguivel de "nao ha regra". A primeira versao deste
# script caiu nessa e anunciou que as portas estavam fechadas quando estavam
# abertas ha tres semanas. Um vigia que mente e pior do que nenhum -- entao,
# sem poder olhar, ele diz que nao olhou.
if (-not $admin) {
  Dizer "sem administrador nao da para ler as regras do firewall - nao conferido"
} else {
foreach ($porta in @(443, 80)) {
  $aberta = Get-NetFirewallPortFilter -ErrorAction SilentlyContinue |
    Where-Object { $_.Protocol -eq 'TCP' -and ($_.LocalPort -eq "$porta" -or $_.LocalPort -contains "$porta") } |
    Get-NetFirewallRule -ErrorAction SilentlyContinue |
    Where-Object { $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' -and $_.Enabled -eq 'True' } |
    Select-Object -First 1

  if ($aberta) {
    Dizer "porta $porta aberta por '$($aberta.DisplayName)' - ok"
  } elseif ($Somente) {
    Anotar "a porta $porta NAO tem regra de entrada ligada"
  } else {
    try {
      New-NetFirewallRule -DisplayName "Gerador-OS $porta" -Direction Inbound -Protocol TCP `
        -LocalPort $porta -Action Allow -Profile Any -Enabled True `
        -Description 'App de ordens de servico servido pelo nginx em Docker.' -ErrorAction Stop | Out-Null
      Anotar "a porta $porta estava sem regra - regra 'Gerador-OS $porta' criada"
    } catch {
      Anotar "FALHA ao abrir a porta ${porta}: $($_.Exception.Message)"
    }
  }
}
}

# ---- c) o cabo, quando tem link, no IP fixo -------------------------------
# Se a placa de sempre nao existe (trocada por um adaptador USB, por exemplo),
# vale qualquer placa com fio que esteja com link -- e o fixar-ip-cabo.ps1
# descobre a dele do mesmo jeito.
$cabo = Get-NetAdapter -Name $PLACA_CABO -ErrorAction SilentlyContinue
if (-not $cabo) {
  $cabo = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
          # Mesma exclusao do fixar-ip-cabo.ps1, e pela mesma razao: o
          # loopback do Audaces se apresenta como placa fisica conectada.
          Where-Object { $_.MediaConnectionState -eq 'Connected' -and
                         $_.InterfaceDescription -notmatch 'Wireless|Wi-Fi|802\.11|Loopback|Virtual|TAP|VPN' -and
                         $_.Name -notmatch '^(vEthernet|Wi-Fi|Topaz|Loopback)' } |
          Select-Object -First 1
  if ($cabo) { Anotar "a placa '$PLACA_CABO' nao existe mais; usando '$($cabo.Name)'"; $PLACA_CABO = $cabo.Name }
}
if (-not $cabo) {
  Dizer "nenhuma placa com fio com link nesta maquina - nada a fazer"
} elseif ($cabo.MediaConnectionState -ne 'Connected') {
  Dizer "cabo sem link - nada a fazer (o Wi-Fi atende sozinho)"
} else {
  $enderecos = (Get-NetIPAddress -InterfaceAlias $PLACA_CABO -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
  if ($enderecos -contains $IP_CABO) {
    Dizer "cabo em $IP_CABO - ok"
    # Link a 100 Mbps numa placa Gigabit e par danificado. Foi o que derrubou
    # o cabo duas vezes em quatro dias; vale registrar quando reaparecer.
    if ($cabo.LinkSpeed -notmatch 'Gbps') {
      Anotar "cabo de pe, mas so a $($cabo.LinkSpeed) - sinal de cabo danificado, vale trocar"
    }
  } elseif ($Somente) {
    Anotar "o cabo tem link e NAO esta em $IP_CABO (esta em: $($enderecos -join ', '))"
  } else {
    Anotar "o cabo voltou e nao esta em $IP_CABO (esta em: $($enderecos -join ', ')) - fixando"
    $fix = Join-Path $PSScriptRoot 'fixar-ip-cabo.ps1'
    if (Test-Path $fix) {
      # Delegado de proposito: aquele script confere se o endereco esta livre
      # antes, prova rota/gateway/DNS/app depois, e devolve para DHCP sozinho
      # se qualquer conferencia falhar. Repetir a logica aqui seria repetir
      # tambem os erros que ela ja aprendeu a nao cometer.
      $saida = & $fix 2>&1 | Out-String
      if ($LASTEXITCODE -eq 0) { Anotar "cabo fixado em $IP_CABO" }
      else { Anotar "o conserto do cabo falhou e se desfez sozinho. Saida:`n$saida" }
    } else {
      Anotar "nao achei $fix - o cabo ficou como estava"
    }
  }
}

# ---- veredito ---------------------------------------------------------------
Write-Host ""
if ($houve) {
  Write-Host "Houve o que anotar - veja $Log" -ForegroundColor Yellow
} elseif (-not $admin) {
  # Dizer "esta tudo bem" sobre o que nao se olhou e o mesmo erro do falso
  # alarme, na direcao contraria -- e a direcao contraria e a que cala.
  Write-Host "Nada a corrigir no que deu para conferir (sem administrador, o firewall ficou de fora)." -ForegroundColor Yellow
} else {
  Write-Host "Nada a fazer: redes Particulares, portas abertas, enderecos no lugar." -ForegroundColor Green
}
