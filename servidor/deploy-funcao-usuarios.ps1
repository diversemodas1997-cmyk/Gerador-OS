<#
  Deploy da funcao "usuarios" (contas de acesso) para o Supabase LOCAL da fabrica.

  Copia servidor/funcoes/usuarios/* para o volume do edge-runtime e reinicia so
  esse container, para carregar o codigo novo. Roda NO PROPRIO SERVIDOR
  (193.168.0.200). E o passo que falta depois de mexer em funcoes/usuarios/index.ts
  — o resto do instalar.ps1 nao precisa rodar de novo.

  A reinicializacao dura segundos e so afeta o PAINEL DE CONTAS (criar/editar
  usuario), que e o unico lugar que chama esta funcao. O app em si (OS, folha,
  expedicao) usa REST/Realtime e nao para.

  Uso:
    powershell -NoProfile -ExecutionPolicy Bypass -File servidor\deploy-funcao-usuarios.ps1
#>
param(
  [string] $PastaSupabase = "C:\supabase",
  [string] $Container     = "supabase-edge-functions"
)
$ErrorActionPreference = "Stop"

$origem  = Join-Path $PSScriptRoot "funcoes\usuarios"
$destino = Join-Path $PastaSupabase "docker\volumes\functions\usuarios"

if (-not (Test-Path (Join-Path $origem "index.ts"))) {
  throw "nao achei $origem\index.ts - rode este script de dentro da pasta do app"
}

# 1) Copia o codigo novo para o volume que o edge-runtime le.
New-Item -ItemType Directory -Force -Path $destino | Out-Null
Copy-Item (Join-Path $origem "*") $destino -Recurse -Force
$tam = (Get-Item (Join-Path $destino "index.ts")).Length
Write-Host ("copiado -> " + $destino + "  (" + $tam + " bytes)") -ForegroundColor Green

# 2) Acha o docker.exe (mesma logica do desligar-servidor.ps1).
$achado = Get-Command docker.exe -ErrorAction SilentlyContinue
if ($achado) { $dockerExe = $achado.Source }
else { $dockerExe = 'C:\Program Files\Docker\Docker\resources\bin\docker.exe' }
if (-not (Test-Path $dockerExe)) { throw "nao achei o docker.exe" }

# 3) Reinicia so o edge-runtime para ele reler as funcoes.
& $dockerExe restart $Container | Out-Null
Write-Host ("reiniciado: " + $Container) -ForegroundColor Green

# 4) Confere que o container voltou.
Start-Sleep -Seconds 3
$status = & $dockerExe ps --filter ("name=" + $Container) --format "{{.Names}}: {{.Status}}"
if ($status) { Write-Host $status -ForegroundColor Cyan }
else { Write-Host "ATENCAO: o container nao aparece rodando - confira com 'docker ps'" -ForegroundColor Yellow }

Write-Host ""
Write-Host "Pronto. Entre como admin, abra Configuracoes -> Contas de acesso e teste o botao Editar." -ForegroundColor Cyan
