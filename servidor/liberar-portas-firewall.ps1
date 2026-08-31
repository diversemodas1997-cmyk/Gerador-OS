# Abre as portas 80 e 443 no firewall do Windows, em TODOS os perfis.
#
# Por que: em 31/08, com o cabo fora, o servidor passou a atender pelo Wi-Fi.
# O Windows classificou essa rede como PUBLICA (Get-NetConnectionProfile ->
# NetworkCategory: Public), e no perfil publico o firewall e mais fechado. Uma
# maquina da fabrica podia chegar ate aqui pela rede e ainda assim levar
# "tempo esgotado", sem nenhuma pista do motivo -- o pacote morre no firewall,
# nao no nginx.
#
# A alternativa seria marcar a rede como Particular. Nao foi o caminho: mudar
# a categoria liga descoberta e compartilhamento de arquivos junto, que e mais
# do que se pediu. Duas regras nomeadas abrem exatamente o que o app precisa e
# nada alem.
#
# E idempotente: rodar de novo nao duplica regra, so confere.
#
# Para desfazer:
#   Remove-NetFirewallRule -DisplayName 'Gerador-OS *'

$ErrorActionPreference = 'Continue'

function Passo($t) { Write-Host "`n>> $t" -ForegroundColor Cyan }
function Ok($t)    { Write-Host "   OK - $t" -ForegroundColor Green }
function Erro($t)  { Write-Host "   $t" -ForegroundColor Red }

$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Erro "Este script precisa ser rodado como ADMINISTRADOR."
  exit 1
}

$REGRAS = @(
  @{ Nome = 'Gerador-OS HTTPS (443)'; Porta = 443 },
  @{ Nome = 'Gerador-OS HTTP (80)';   Porta = 80  }
)

$falhas = @()

# Procurar pelo NOME nao basta, e a primeira versao deste script errou nisso:
# a instalacao de 10/08 ja tinha criado 'Gerador-OS 80' e 'Gerador-OS 443', e
# como os nomes nao batiam, o script criou duas regras a mais para as mesmas
# portas. Nao quebra nada, mas polui o firewall justamente onde alguem vai
# procurar respostas um dia. Quem responde se a porta esta aberta e a PORTA.
function PortaJaAberta($porta) {
  Get-NetFirewallPortFilter -ErrorAction SilentlyContinue |
    Where-Object { $_.Protocol -eq 'TCP' -and ($_.LocalPort -eq "$porta" -or $_.LocalPort -contains "$porta") } |
    Get-NetFirewallRule -ErrorAction SilentlyContinue |
    Where-Object { $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' -and $_.Enabled -eq 'True' } |
    Select-Object -First 1
}

foreach ($r in $REGRAS) {
  Passo "Porta TCP $($r.Porta)"

  $aberta = PortaJaAberta $r.Porta
  if ($aberta) {
    Ok "ja esta aberta pela regra '$($aberta.DisplayName)' - nao vou criar outra."
    continue
  }

  $ja = Get-NetFirewallRule -DisplayName $r.Nome -ErrorAction SilentlyContinue
  if ($ja) {
    # Existir nao basta: uma regra desligada, ou virada para Block, engana a
    # conferencia. Melhor reafirmar o estado do que confiar no nome.
    Set-NetFirewallRule -DisplayName $r.Nome -Enabled True -Action Allow -Profile Any -ErrorAction SilentlyContinue
    Ok "ja existia - reafirmada como ligada, Allow, todos os perfis."
  } else {
    try {
      New-NetFirewallRule -DisplayName $r.Nome -Direction Inbound -Protocol TCP `
        -LocalPort $r.Porta -Action Allow -Profile Any -Enabled True `
        -Description 'App de ordens de servico servido pelo nginx em Docker.' -ErrorAction Stop | Out-Null
      Ok "criada."
    } catch {
      $falhas += "nao consegui criar a regra da porta $($r.Porta): $($_.Exception.Message)"
    }
  }
}

Passo "Como ficaram"
Get-NetFirewallRule -DisplayName 'Gerador-OS *' -ErrorAction SilentlyContinue |
  Select-Object DisplayName, Enabled, Action, Direction, Profile |
  Format-Table -AutoSize | Out-String | Write-Host

Write-Host ""
if ($falhas.Count -eq 0) {
  Write-Host "PRONTO. As portas 80 e 443 aceitam conexao de qualquer rede desta maquina." -ForegroundColor Green
  Write-Host "Isto abre o caminho ate o nginx; quem responde e o que decide o resto e ele." -ForegroundColor Green
} else {
  Write-Host "DEU ERRADO:" -ForegroundColor Red
  $falhas | ForEach-Object { Erro "- $_" }
  exit 1
}
