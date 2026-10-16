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
    c) o cabo, quando tem link, esta no 192.168.1.50 fixo (era 193.168.0.200
       ate a troca do roteador em 14/10/2026);
    d) o Wi-Fi esta ASSOCIADO ao SSID que ele ja conhece;
    e) o 192.168.1.50 nunca some: com o cabo sem link, ele vai para o Wi-Fi.

  O (a) e o (b) nao tem risco: nao mexem em endereco nem derrubam conexao. O
  (c) mexe, e por isso e delegado ao `fixar-ip-cabo.ps1`, que confere se o
  endereco esta livre antes, prova rota/gateway/DNS/app depois, e devolve a
  placa para DHCP sozinho se qualquer conferencia falhar.

  O ENDERECO do Wi-Fi continua sem ser mexido, de proposito. E por ele que
  esta maquina fala com a internet -- backup, espelho e o proprio atendimento
  remoto -- e troca-lo as cegas, sem ninguem olhando, e arriscar o que nao
  precisa ser arriscado. Ele ja esta fixo; se sair do lugar, o log avisa.

  O (d) e outra coisa, e nasceu do reboot de 31/08/2026: a maquina voltou
  com o cabo perfeito e o Wi-Fi DESCONECTADO -- placa habilitada, radio
  ligado, associada a SSID nenhuma, porque o perfil salvo estava em
  "conectar manualmente". O https://192.168.1.158 ficou mudo e nada
  denunciava: o .200 respondia, os 12 conteineres de pe, o resto verde.
  Reassociar a uma rede que a maquina ja conhece nao mexe em endereco nem
  derruba conexao -- e da mesma classe do (a) e do (b).

  O (e) nasceu em 16/10/2026 (Junior: "quando a rede por cabo cair, migrar
  para o endereco IP do Wi-Fi"). As maquinas da fabrica e os celulares abrem
  o https://192.168.1.50, e com o cabo fora esse endereco morria junto: o
  servidor continuava de pe no .158, mas ninguem ia ate la. Trocar o
  endereco nas maquinas nao serve -- para o navegador o .158 e outro site,
  com outro login e outro icone. Entao quem muda de lugar e o ENDERECO:

    - cabo sem link e Wi-Fi de pe -> o .50 entra no Wi-Fi como SEGUNDO
      endereco (o .158 continua la). Vai so na memoria (ActiveStore): um
      reboot nao o deixa preso no Wi-Fi. E SkipAsSource: a maquina continua
      SAINDO pelo .158 -- internet, backup e espelho nao mudam. Antes, um
      ping confere que nenhum outro aparelho responde no .50; depois, um
      ping a partir do .50 ensina o roteador onde ele mora agora.
    - o cabo voltou -> o .50 sai do Wi-Fi. Se o cabo acordou com o endereco
      marcado como DUPLICADO (ele testou a rede enquanto o Wi-Fi ainda
      segurava o .50), a placa do cabo e reiniciada para testar de novo.

  A porta 443 escuta em todos os enderecos (0.0.0.0) e o .50 esta no
  certificado: nada mais precisa mudar para ele ser atendido pelo Wi-Fi.
  A tarefa (-Agendar) dispara tambem quando uma rede CAI ou VOLTA (eventos
  10000/10001 do NetworkProfile), e nao so a cada 15 min: um minuto sem o
  .50 ja e a fabrica parada.

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
$IP_CABO    = '192.168.1.50'
$PLACA_WIFI = 'Wi-Fi'
$GATEWAY    = '192.168.1.1'

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
  # A REDE CAIU OU VOLTOU: o NetworkProfile grava 10001 quando uma rede some e
  # 10000 quando conecta. E o gatilho que faz o .50 mudar de placa em
  # segundos, e nao no proximo quarto de hora.
  $classe = Get-CimClass -ClassName MSFT_TaskEventTrigger -Namespace Root/Microsoft/Windows/TaskScheduler
  $naRede = New-CimInstance -CimClass $classe -ClientOnly
  $naRede.Subscription = '<QueryList><Query Id="0" Path="Microsoft-Windows-NetworkProfile/Operational">' +
    '<Select Path="Microsoft-Windows-NetworkProfile/Operational">*[System[(EventID=10000 or EventID=10001)]]</Select>' +
    '</Query></QueryList>'
  $naRede.Enabled = $true
  # Queue: um evento que chega com o vigia rodando espera a vez, em vez de ser
  # descartado -- o cabo que volta no meio da passagem do "caiu" nao se perde.
  $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
            -MultipleInstances Queue
  try {
    Unregister-ScheduledTask -TaskName $nome -Confirm:$false -ErrorAction SilentlyContinue
    # -RunLevel Highest: sem isto a tarefa roda sem administrador, e tudo o
    # que este script conserta exige administrador. Falharia em silencio.
    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger @($noLogon, $repetindo, $naRede) `
      -Settings $conf -RunLevel Highest `
      -Description 'Mantem as redes do servidor como Particular, as portas abertas, o cabo no IP fixo e o 192.168.1.50 no Wi-Fi quando o cabo cai.' `
      -ErrorAction Stop | Out-Null
    Anotar "tarefa '$nome' registrada: no logon, a cada $CadaMinutos min e quando uma rede cai ou volta"
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
          # loopback do Warsaw/Topaz (dito 'do Audaces' ate 17/09/2026) se
          # apresenta como placa fisica conectada.
          Where-Object { $_.MediaConnectionState -eq 'Connected' -and
                         $_.InterfaceDescription -notmatch 'Wireless|Wi-Fi|802\.11|Loopback|Virtual|TAP|VPN' -and
                         $_.Name -notmatch '^(vEthernet|Wi-Fi|Topaz|Loopback)' } |
          Select-Object -First 1
  if ($cabo) { Anotar "a placa '$PLACA_CABO' nao existe mais; usando '$($cabo.Name)'"; $PLACA_CABO = $cabo.Name }
}
# ---- e) o .50 acompanha a placa que estiver de pe -------------------------
# Ver o cabecalho. Vem ANTES do (c): quando o cabo volta, o .50 tem de sair do
# Wi-Fi primeiro, senao o cabo enxerga o proprio endereco "em uso".
$caboComLink = $cabo -and $cabo.MediaConnectionState -eq 'Connected'
$wifiAgora   = Get-NetAdapter -Name $PLACA_WIFI -ErrorAction SilentlyContinue
$wifiDePe    = $wifiAgora -and $wifiAgora.Status -eq 'Up'
$noWifi      = Get-NetIPAddress -InterfaceAlias $PLACA_WIFI -AddressFamily IPv4 -ErrorAction SilentlyContinue |
               Where-Object { $_.IPAddress -eq $IP_CABO }
if ($caboComLink -and $noWifi) {
  if ($Somente) {
    Anotar "o cabo tem link e o $IP_CABO ainda esta emprestado ao Wi-Fi"
  } else {
    try {
      Remove-NetIPAddress -InterfaceAlias $PLACA_WIFI -IPAddress $IP_CABO -Confirm:$false -ErrorAction Stop
      Anotar "o cabo voltou - $IP_CABO saiu do Wi-Fi e volta para o cabo"
    } catch {
      Anotar "FALHA ao tirar $IP_CABO do Wi-Fi: $($_.Exception.Message)"
    }
  }
} elseif (-not $caboComLink -and $wifiDePe -and -not $noWifi) {
  if ($Somente) {
    Anotar "cabo sem link: o $IP_CABO esta fora do ar (o Wi-Fi atende so no proprio endereco)"
  } else {
    # Nenhum OUTRO aparelho pode estar no .50. Se o endereco ainda consta ativo
    # nesta maquina (placa do cabo sem link mas com ele preso), o ping responde
    # daqui mesmo e nao prova nada -- e o New-NetIPAddress abaixo e quem diz.
    $ativosAqui = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                   Where-Object { $_.AddressState -eq 'Preferred' }).IPAddress
    $ocupado = ($ativosAqui -notcontains $IP_CABO) -and
               (Test-Connection -ComputerName $IP_CABO -Count 2 -Quiet -ErrorAction SilentlyContinue)
    if ($ocupado) {
      Anotar "cabo sem link, mas OUTRO aparelho responde em $IP_CABO - nao emprestado ao Wi-Fi (conflito de IP)"
    } else {
      $erro = $null
      try {
        New-NetIPAddress -InterfaceAlias $PLACA_WIFI -IPAddress $IP_CABO -PrefixLength 24 `
          -SkipAsSource $true -PolicyStore ActiveStore -ErrorAction Stop | Out-Null
      } catch {
        # O cmdlet ja entrou pela metade nesta maquina (ver fixar-ip-cabo.ps1);
        # o netsh e o caminho de reserva, com as mesmas duas garantias.
        $erro = $_.Exception.Message
        $out = (& netsh @('interface','ipv4','add','address',"name=$PLACA_WIFI",$IP_CABO,'255.255.255.0',
                          'store=active','skipassource=true') 2>&1 | Out-String).Trim()
        if (-not $out) { $erro = $null } else { $erro = "$erro / netsh: $out" }
      }
      if ($erro) {
        Anotar "cabo sem link - FALHA ao emprestar $IP_CABO ao Wi-Fi: $erro"
      } else {
        # O Windows testa o endereco na rede antes de usa-lo (Tentative).
        $estado = ''
        foreach ($i in 1..10) {
          Start-Sleep -Seconds 1
          $estado = (Get-NetIPAddress -InterfaceAlias $PLACA_WIFI -IPAddress $IP_CABO -ErrorAction SilentlyContinue).AddressState
          if ($estado -eq 'Preferred' -or $estado -eq 'Duplicate') { break }
        }
        if ($estado -eq 'Duplicate') {
          Remove-NetIPAddress -InterfaceAlias $PLACA_WIFI -IPAddress $IP_CABO -Confirm:$false -ErrorAction SilentlyContinue
          Anotar "cabo sem link - o Wi-Fi achou $IP_CABO DUPLICADO na rede; desfeito (conflito de IP)"
        } else {
          # Ensina o roteador (e quem perguntar) que o .50 agora mora no Wi-Fi.
          & ping.exe -n 2 -w 1000 -S $IP_CABO $GATEWAY 2>&1 | Out-Null
          $ProgressPreference = 'SilentlyContinue'
          $resp = ''
          try {
            $r = Invoke-WebRequest -Uri "https://$IP_CABO" -TimeoutSec 15 -UseBasicParsing
            $resp = "o app responde em https://$IP_CABO (HTTP $($r.StatusCode))"
          } catch { $resp = "o app NAO respondeu em https://${IP_CABO}: $($_.Exception.Message)" }
          Anotar "cabo sem link - $IP_CABO emprestado ao Wi-Fi ($estado); $resp"
        }
      }
    }
  }
} elseif (-not $caboComLink -and $noWifi) {
  Dizer "cabo sem link - $IP_CABO atendendo pelo Wi-Fi - ok"
}

if (-not $cabo) {
  Dizer "nenhuma placa com fio com link nesta maquina - nada a fazer"
} elseif ($cabo.MediaConnectionState -ne 'Connected') {
  Dizer "cabo sem link - o Wi-Fi atende (ver a secao do $IP_CABO acima)"
} else {
  $enderecos = (Get-NetIPAddress -InterfaceAlias $PLACA_CABO -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
  $estadoCabo = (Get-NetIPAddress -InterfaceAlias $PLACA_CABO -IPAddress $IP_CABO -ErrorAction SilentlyContinue).AddressState
  # O cabo que volta com o .50 ainda no Wi-Fi testa a rede, acha o endereco "em
  # uso" e o marca DUPLICADO -- fica com ele no papel e sem atender. Com o .50
  # ja fora do Wi-Fi (secao e, logo acima), reiniciar a placa refaz o teste.
  if ($enderecos -contains $IP_CABO -and $estadoCabo -eq 'Duplicate' -and -not $Somente) {
    Restart-NetAdapter -Name $PLACA_CABO -Confirm:$false -ErrorAction SilentlyContinue
    foreach ($i in 1..30) {
      Start-Sleep -Seconds 1
      $estadoCabo = (Get-NetIPAddress -InterfaceAlias $PLACA_CABO -IPAddress $IP_CABO -ErrorAction SilentlyContinue).AddressState
      if ($estadoCabo -eq 'Preferred') { break }
    }
    Anotar "o cabo voltou com $IP_CABO DUPLICADO - placa reiniciada, agora: $estadoCabo"
  }
  if ($enderecos -contains $IP_CABO -and $estadoCabo -eq 'Duplicate') {
    Anotar "o cabo esta com $IP_CABO DUPLICADO - outro aparelho ou o Wi-Fi ainda o segura"
  } elseif ($enderecos -contains $IP_CABO) {
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

# ---- d) o Wi-Fi associado ao SSID que ele conhece --------------------------
# So reassocia: nao escolhe rede nova, nao digita senha, nao toca no endereco.
# Se nao ha perfil salvo, nao ha o que fazer -- e dizer isso e melhor do que
# adivinhar a qual rede esta maquina deveria entrar.
$wifi = Get-NetAdapter -Name $PLACA_WIFI -ErrorAction SilentlyContinue
if (-not $wifi) {
  Dizer "nao ha placa '$PLACA_WIFI' nesta maquina - nada a fazer"
} elseif ($wifi.Status -eq 'Disabled') {
  Anotar "a placa '$PLACA_WIFI' esta DESABILITADA - nao ha caminho pelo Wi-Fi ate alguem religa-la"
} elseif ($wifi.Status -eq 'Up') {
  Dizer "Wi-Fi associado - ok"
} else {
  # O nome do perfil sai do proprio netsh, e nao chumbado aqui: o SSID desta
  # maquina ja mudou uma vez, e um nome escrito no script envelhece calado.
  $perfilWifi = $null
  foreach ($linha in (& netsh @('wlan','show','profiles') 2>&1)) {
    $m = [regex]::Match([string]$linha, '^\s*(?:Todos os Perfis de Usu.rios|All User Profile)\s*:\s*(.+?)\s*$')
    if ($m.Success) { $perfilWifi = $m.Groups[1].Value; break }
  }
  if (-not $perfilWifi) {
    Anotar "o Wi-Fi esta desconectado e nao ha perfil de rede salvo - so a mao"
  } elseif ($Somente) {
    Anotar "o Wi-Fi esta DESCONECTADO (perfil salvo: '$perfilWifi') - o https://192.168.1.158 nao responde"
  } else {
    # connectionmode=auto e o conserto de raiz: sem ele, todo reboot volta com
    # o Wi-Fi fora. O connect logo abaixo resolve o agora.
    & netsh @('wlan','set','profileparameter',"name=$perfilWifi",'connectionmode=auto') 2>&1 | Out-Null
    & netsh @('wlan','connect',"name=$perfilWifi") 2>&1 | Out-Null
    $voltou = $false
    foreach ($i in 1..10) {
      Start-Sleep -Seconds 3
      if ((Get-NetAdapter -Name $PLACA_WIFI -ErrorAction SilentlyContinue).Status -eq 'Up') { $voltou = $true; break }
    }
    if ($voltou) { Anotar "o Wi-Fi estava desconectado - reassociado a '$perfilWifi'" }
    else         { Anotar "o Wi-Fi estava desconectado e NAO voltou em 30 s (perfil '$perfilWifi')" }
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
