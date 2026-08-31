# Deixa o firewall do Windows aceitar as conexoes que chegam ao app.
#
# Sao DUAS coisas, e a segunda e a que engana:
#
#   1) existir regra de entrada para as portas 80 e 443;
#   2) o PERFIL da rede aceitar regras de entrada.
#
# O que aconteceu em 31/08/2026: com o cabo fora, o servidor passou a atender
# pelo Wi-Fi. As regras 'Gerador-OS 80' e 'Gerador-OS 443' existiam desde a
# instalacao de 10/08, ligadas, Allow, perfil Any -- e mesmo assim NENHUMA
# maquina da fabrica chegava ao nginx. O motivo estava numa linha so:
#
#     Get-NetFirewallProfile -Name Public  ->  AllowInboundRules : False
#
# Essa e a caixa "Bloquear todas as conexoes de entrada, inclusive as da lista
# de aplicativos permitidos" do Firewall do Windows. Com ela marcada, o perfil
# Publico IGNORA todas as regras de permissao -- a regra existe, esta ligada, e
# nao vale nada. E o Windows classificou a rede do Wi-Fi como Publica.
#
# O conserto e classificar a rede como PARTICULAR, e nao desmarcar aquela caixa:
# desmarcar valeria para toda rede publica que esta maquina encontrar um dia,
# incluindo as regras do AnyDesk e do Audaces. Marcar a rede da fabrica como
# particular vale so para ela. No perfil Particular, AllowInboundRules ja e True.
#
# Sintoma que leva ate aqui: o app abre na maquina do servidor (laco local, nao
# passa pelo firewall) e nao abre em nenhuma outra -- e o `netstat -ano` na
# porta 443 nao mostra UMA conexao vinda de fora.
#
# Uso:
#   .\liberar-portas-firewall.ps1              (placa Wi-Fi, o padrao de hoje)
#   .\liberar-portas-firewall.ps1 -Placa 'Ethernet 3'
#
# Para desfazer:
#   Set-NetConnectionProfile -InterfaceAlias '<placa>' -NetworkCategory Public
#   Remove-NetFirewallRule -DisplayName 'Gerador-OS *'

param(
  [string] $Placa = 'Wi-Fi'
)

$ErrorActionPreference = 'Continue'

function Passo($t) { Write-Host "`n>> $t" -ForegroundColor Cyan }
function Ok($t)    { Write-Host "   OK - $t" -ForegroundColor Green }
function Aviso($t) { Write-Host "   [AVISO] $t" -ForegroundColor Yellow }
function Erro($t)  { Write-Host "   $t" -ForegroundColor Red }

$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Erro "Este script precisa ser rodado como ADMINISTRADOR."
  exit 1
}

$falhas = @()

# ---- 1. as portas tem regra? ----------------------------------------------
# Procurar pelo NOME nao basta, e a primeira versao deste script errou nisso: a
# instalacao de 10/08 ja tinha criado 'Gerador-OS 80' e 'Gerador-OS 443', e como
# os nomes nao batiam, o script criou duas regras a mais para as mesmas portas.
# Nao quebra nada, mas polui o firewall justo onde alguem vai procurar resposta
# um dia. Quem responde se a porta esta aberta e a PORTA.
function PortaJaAberta($porta) {
  Get-NetFirewallPortFilter -ErrorAction SilentlyContinue |
    Where-Object { $_.Protocol -eq 'TCP' -and ($_.LocalPort -eq "$porta" -or $_.LocalPort -contains "$porta") } |
    Get-NetFirewallRule -ErrorAction SilentlyContinue |
    Where-Object { $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' -and $_.Enabled -eq 'True' } |
    Select-Object -First 1
}

foreach ($porta in @(443, 80)) {
  Passo "Porta TCP $porta"
  $aberta = PortaJaAberta $porta
  if ($aberta) {
    Ok "ja esta aberta pela regra '$($aberta.DisplayName)' - nao vou criar outra."
  } else {
    try {
      New-NetFirewallRule -DisplayName "Gerador-OS $porta" -Direction Inbound -Protocol TCP `
        -LocalPort $porta -Action Allow -Profile Any -Enabled True `
        -Description 'App de ordens de servico servido pelo nginx em Docker.' -ErrorAction Stop | Out-Null
      Ok "regra 'Gerador-OS $porta' criada."
    } catch {
      $falhas += "nao consegui criar a regra da porta ${porta}: $($_.Exception.Message)"
    }
  }
}

# ---- 2. o perfil da rede aceita regras de entrada? ------------------------
# Este e o passo que faltava. Uma regra perfeita num perfil que ignora regras
# nao serve para nada, e nao ha erro nenhum para ler: a conexao simplesmente
# nunca chega.
Passo "Perfil da rede na placa '$Placa'"

$perfil = Get-NetConnectionProfile -InterfaceAlias $Placa -ErrorAction SilentlyContinue
if (-not $perfil) {
  $falhas += "a placa '$Placa' nao tem rede ativa - nao da para conferir o perfil"
} else {
  Write-Host "   rede '$($perfil.Name)', categoria atual: $($perfil.NetworkCategory)"

  if ($perfil.NetworkCategory -eq 'Public') {
    $publico = Get-NetFirewallProfile -Name Public
    if ($publico.AllowInboundRules -eq $false) {
      Aviso "o perfil Publico esta com AllowInboundRules=False: ele IGNORA toda regra de entrada."
      Aviso "e por isto que o app abre aqui e nao abre em outra maquina."
    }
    try {
      Set-NetConnectionProfile -InterfaceAlias $Placa -NetworkCategory Private -ErrorAction Stop
      Start-Sleep -Seconds 2
      $agora = (Get-NetConnectionProfile -InterfaceAlias $Placa -ErrorAction SilentlyContinue).NetworkCategory
      if ($agora -eq 'Private') { Ok "rede reclassificada como Particular." }
      else { $falhas += "a rede nao aceitou virar Particular (esta como $agora)" }
    } catch {
      $falhas += "nao consegui reclassificar a rede: $($_.Exception.Message)"
    }
  } else {
    Ok "ja e '$($perfil.NetworkCategory)' - as regras de entrada valem nesta rede."
  }
}

# ---- 3. conferir ------------------------------------------------------------
Passo "Como ficou"
Get-NetConnectionProfile -InterfaceAlias $Placa -ErrorAction SilentlyContinue |
  Select-Object InterfaceAlias, Name, NetworkCategory | Format-Table -AutoSize | Out-String | Write-Host
Get-NetFirewallProfile | Select-Object Name, Enabled, AllowInboundRules |
  Format-Table -AutoSize | Out-String | Write-Host
Get-NetFirewallRule -DisplayName 'Gerador-OS *' -ErrorAction SilentlyContinue |
  Select-Object DisplayName, Enabled, Action, Profile | Format-Table -AutoSize | Out-String | Write-Host

Write-Host ""
if ($falhas.Count -eq 0) {
  Write-Host "PRONTO. As portas tem regra E a rede aceita regras de entrada." -ForegroundColor Green
  Write-Host "Agora o teste que vale e de OUTRA maquina - daqui, o laco local nao" -ForegroundColor Green
  Write-Host "passa pelo firewall e por isso nunca acusou o problema." -ForegroundColor Green
} else {
  Write-Host "DEU ERRADO:" -ForegroundColor Red
  $falhas | ForEach-Object { Erro "- $_" }
  exit 1
}
