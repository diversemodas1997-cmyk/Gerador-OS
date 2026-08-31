<#
  Monta a pasta que vai para as maquinas da fabrica instalarem o certificado.

  POR QUE ISTO EXISTE
  O instalador precisa de DOIS arquivos lado a lado: o `instalar-certificado.cmd`
  e o `ca.crt`. Um sem o outro nao serve, e o ca.crt mora em `servidor\tls\`,
  que e uma pasta que nao vai para o git. Copiar a mao os dois, toda vez, e
  como se esquece um deles.

  Este script junta os dois numa pasta pronta para levar no pendrive ou por
  rede — e CONFERE, antes de copiar, que o ca.crt e mesmo o que assinou o
  certificado do servidor. Levar um CA que nao assina nada e instalar confianca
  em coisa nenhuma.

  COMO USAR
    .\servidor\preparar-instalador.ps1
    .\servidor\preparar-instalador.ps1 -Destino 'D:\Instalar Gerador-OS'

  Depois: copie a PASTA inteira para a maquina e clique duas vezes no
  `instalar-certificado.cmd`. Ele pede a permissao de administrador sozinho.
#>
[CmdletBinding()]
param(
  [string] $Destino = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Instalar Gerador-OS')
)

$ErrorActionPreference = 'Stop'
$Raiz = Split-Path -Parent $PSScriptRoot
$ca   = Join-Path $Raiz 'servidor\tls\ca.crt'
$srv  = Join-Path $Raiz 'servidor\tls\servidor.crt'
$cmd  = Join-Path $Raiz 'servidor\instalar-certificado.cmd'
$abre = Join-Path $Raiz 'servidor\abrir-gerador-os.cmd'

foreach ($f in @($ca, $cmd, $abre)) {
  if (-not (Test-Path $f)) { throw "nao achei $f" }
}

# A impressao digital que o instalador confere tem de ser a DESTE ca.crt. Se
# alguem gerar um certificado novo e esquecer de atualizar o .cmd, o instalador
# diria "instalado" e depois "nao aparece na lista" na maquina da fabrica — longe
# de quem poderia consertar. Melhor parar aqui.
$digital = (Get-FileHash -Path $ca -Algorithm SHA1).Hash  # do ARQUIVO, so para log
$cert    = New-Object Security.Cryptography.X509Certificates.X509Certificate2 $ca
$thumb   = $cert.Thumbprint
$texto   = Get-Content $cmd -Raw
if ($texto -notmatch [regex]::Escape($thumb)) {
  throw ("o instalador confere uma impressao digital diferente da do ca.crt.`n" +
         "  ca.crt : $thumb`n" +
         "  conserte a linha `"set `"DIGITAL=...`"`" em servidor\instalar-certificado.cmd")
}

# E o CA tem de assinar mesmo o certificado do servidor.
if (Test-Path $srv) {
  $s = New-Object Security.Cryptography.X509Certificates.X509Certificate2 $srv
  if ($s.Issuer -ne $cert.Subject) {
    throw ("o ca.crt nao e quem assinou o servidor.crt.`n" +
           "  assinou o servidor : $($s.Issuer)`n" +
           "  este ca.crt e      : $($cert.Subject)")
  }
}

New-Item -ItemType Directory -Force -Path $Destino | Out-Null
Copy-Item $ca  (Join-Path $Destino 'ca.crt') -Force
Copy-Item $cmd (Join-Path $Destino 'instalar-certificado.cmd') -Force
# O lancador vai junto: e ele que o instalador copia para a maquina e vira o
# atalho da area de trabalho. Sem ele ao lado, o instalador cai no atalho
# antigo, de endereco fixo -- que e a armadilha que derrubou a fabrica em 31/08.
Copy-Item $abre (Join-Path $Destino 'abrir-gerador-os.cmd') -Force

# Um LEIA-ME curto: quem leva o pendrive nao e quem escreveu isto.
$leia = @"
GERADOR-OS - instalar o certificado nesta maquina

1. Clique duas vezes em  instalar-certificado.cmd
2. Quando o Windows perguntar, clique em SIM (permissao de administrador).
3. Espere a mensagem "Pronto" e feche.

Depois disso, abra o atalho "Gerador-OS" que aparece na area de trabalho.
Ele PROCURA o servidor sozinho - pelo nome, pelo cabo ou pelo Wi-Fi - e
por isso continua funcionando quando o servidor troca de rede.

Se precisar digitar:  https://GERADOR-OS

Nao separe os arquivos desta pasta: o instalador procura o ca.crt e o
abrir-gerador-os.cmd ao lado dele.

Certificado: $($cert.Subject)
Vale ate:    $($cert.NotAfter.ToString('dd/MM/yyyy'))
"@
Set-Content -Path (Join-Path $Destino 'LEIA-ME.txt') -Value $leia -Encoding utf8

Write-Host ''
Write-Host "  Pasta pronta: $Destino"
Write-Host ''
Get-ChildItem $Destino | ForEach-Object { Write-Host ("    {0,-28} {1,7} bytes" -f $_.Name, $_.Length) }
Write-Host ''
Write-Host "  Certificado : $($cert.Subject)"
Write-Host "  Vale ate    : $($cert.NotAfter.ToString('dd/MM/yyyy'))"
Write-Host "  Impressao   : $thumb"
Write-Host ''
Write-Host '  Copie a PASTA inteira para a maquina e clique duas vezes no'
Write-Host '  instalar-certificado.cmd.'
Write-Host ''
