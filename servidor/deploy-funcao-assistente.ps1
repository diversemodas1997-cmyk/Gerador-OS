<#
  Instala (ou atualiza) a funcao "assistente" - o chat do Claude dentro do
  Gerador-OS - no Supabase LOCAL da fabrica.

  O que faz:
    1. Copia servidor/funcoes/assistente/index.ts para o volume do edge-runtime.
    2. Na primeira vez (ou com -TrocarChave), pede a chave da API da Anthropic
       e grava em chave.txt AO LADO da funcao, dentro de C:\supabase. A chave
       NUNCA vai para o git nem para a pasta do app (que o nginx publica).
    3. Opcional: -Modelo grava modelo.txt (padrao: claude-sonnet-5-5).
    4. Reinicia so o container das funcoes. O app (OS, folha, expedicao) nao para.

  Roda NO PROPRIO SERVIDOR, de dentro da pasta do Gerador-OS:
    powershell -NoProfile -ExecutionPolicy Bypass -File servidor\deploy-funcao-assistente.ps1

  Trocar a chave depois:
    powershell -NoProfile -ExecutionPolicy Bypass -File servidor\deploy-funcao-assistente.ps1 -TrocarChave
#>
param(
  [string] $PastaSupabase = "C:\supabase",
  [string] $Container     = "supabase-edge-functions",
  [switch] $TrocarChave,
  [string] $Modelo        = ""
)
$ErrorActionPreference = "Stop"

$origem  = Join-Path $PSScriptRoot "funcoes\assistente"
$destino = Join-Path $PastaSupabase "docker\volumes\functions\assistente"

if (-not (Test-Path (Join-Path $origem "index.ts"))) {
  throw "nao achei $origem\index.ts - rode este script de dentro da pasta do app (e faca o git pull antes)"
}

# 1) Codigo novo no volume que o edge-runtime le.
New-Item -ItemType Directory -Force -Path $destino | Out-Null
Copy-Item (Join-Path $origem "index.ts") $destino -Force
Write-Host ("copiado -> " + $destino) -ForegroundColor Green

# 2) Chave da Anthropic.
$arqChave = Join-Path $destino "chave.txt"
$semBom = New-Object System.Text.UTF8Encoding $false
if ($TrocarChave -or -not (Test-Path $arqChave)) {
  Write-Host ""
  Write-Host "Cole a chave da API da Anthropic (comeca com sk-ant-) e aperte Enter." -ForegroundColor Cyan
  Write-Host "(Ela nao aparece na tela enquanto voce cola - e normal.)" -ForegroundColor DarkGray
  $seguro = Read-Host -AsSecureString "Chave"
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($seguro)
  try { $chave = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr).Trim() }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
  if (-not $chave.StartsWith("sk-ant-")) {
    throw "isso nao parece uma chave da Anthropic (deveria comecar com sk-ant-). Nada foi gravado."
  }
  [System.IO.File]::WriteAllText($arqChave, $chave, $semBom)
  $chave = $null
  Write-Host "chave gravada em $arqChave" -ForegroundColor Green
} else {
  Write-Host "chave ja existe (use -TrocarChave para trocar)" -ForegroundColor DarkGray
}

# 3) Modelo (opcional).
if ($Modelo) {
  [System.IO.File]::WriteAllText((Join-Path $destino "modelo.txt"), $Modelo.Trim(), $semBom)
  Write-Host ("modelo -> " + $Modelo) -ForegroundColor Green
}

# 4) Reinicia so o edge-runtime.
$achado = Get-Command docker.exe -ErrorAction SilentlyContinue
if ($achado) { $dockerExe = $achado.Source }
else { $dockerExe = 'C:\Program Files\Docker\Docker\resources\bin\docker.exe' }
if (-not (Test-Path $dockerExe)) { throw "nao achei o docker.exe" }

& $dockerExe restart $Container | Out-Null
Write-Host ("reiniciado: " + $Container) -ForegroundColor Green
Start-Sleep -Seconds 3
$status = & $dockerExe ps --filter ("name=" + $Container) --format "{{.Names}}: {{.Status}}"
if ($status) { Write-Host $status -ForegroundColor Cyan }
else { Write-Host "ATENCAO: o container nao aparece rodando - confira com 'docker ps'" -ForegroundColor Yellow }

Write-Host ""
Write-Host "Pronto. Abra o Gerador-OS, aperte F5 e clique no botao 'Assistente' no canto de baixo." -ForegroundColor Cyan
