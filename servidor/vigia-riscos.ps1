<#
  Mantem dados/riscos-pdf.json em dia com o que esta na pasta de riscos.

  POR QUE ISTO EXISTE
  Em 01/09/2026 a OS 0508 avisou que a medida da grade nao veio de um risco -- e
  o aviso estava certo, mas por um motivo que ninguem tinha como adivinhar:

    08:58  os tres PDFs de encaixe da CM.TRI P-G-GG-G2 entram na pasta
    09:02  a grade e criada, com os numeros certos e sem apontar arquivo nenhum
    12:04  a OS e criada -> "a medida desta grade nao veio de um risco"

  O navegador nao lista pasta: a lista de PDFs que o app le e um ARQUIVO
  ESTATICO, dados/riscos-pdf.json, gerado pelo indexar-riscos.js. A ultima
  geracao era de 31/08 as 14:36 -- antes de os PDFs existirem. Enquanto a lista
  nao sabe do arquivo, a coluna Riscos da grade fica vazia, nao ha de onde
  importar, e a medida acaba digitada a mao. O passo esquecido nao da erro
  nenhum: ele so faz o programa parecer errado dias depois.

  O QUE ELE FAZ, a cada passagem: roda o indexador. So isso. O indexador varre a
  pasta e SO REESCREVE o JSON quando a lista mudou de verdade -- entao rodar de
  cinco em cinco minutos nao suja o git nem gasta disco a toa.

  POR QUE UMA TAREFA E NAO UM VIGIA DE ARQUIVO (FileSystemWatcher)
  Porque o vigia de arquivo e um processo que precisa estar VIVO, e um processo
  que morre em silencio devolve exatamente o problema que este script existe
  para resolver -- com o agravante de todo mundo achar que esta resolvido. A
  tarefa agendada e reiniciada pelo proprio Windows, sobrevive a reboot (o
  servidor entra com autologon) e nao tem estado nenhum para se perder. A varredura
  custa milissegundos: sao ~270 arquivos.

  O QUE ELE NAO FAZ: commit e push. A copia da nuvem (GitHub Pages) so muda com
  um commit, e commit automatico sem ninguem olhando nao e coisa que se faca com
  a pasta de producao. Na fabrica -- que e onde as pessoas trabalham -- o efeito
  e imediato: o nginx serve esta pasta direto, entao o JSON regravado ja esta no
  ar. O log abaixo e o que avisa que ha algo para commitar.

  RODAR A MAO:   powershell -File servidor\vigia-riscos.ps1
  AGENDAR:       powershell -File servidor\vigia-riscos.ps1 -Agendar  (como admin)
#>

[CmdletBinding()]
param(
  [string]$Repo = 'C:\Users\Pichau\Desktop\Gerador-OS',
  [string]$Log  = 'C:\Users\Pichau\Desktop\Gerador-OS\servidor\tls\vigia-riscos.log',
  # Registra a tarefa agendada e sai. Precisa de ADMINISTRADOR -- registrar
  # tarefa e a unica parte disto que o Windows nao deixa fazer de usuario comum.
  [switch]$Agendar
)

$ErrorActionPreference = 'Continue'

function Anotar($texto) {
  $linha = ('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $texto)
  try { Add-Content -Path $Log -Value $linha -Encoding utf8 } catch { }
  Write-Host $linha
}

if ($Agendar) {
  $ehAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
      [Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $ehAdmin) {
    Write-Host 'Registrar tarefa exige ADMINISTRADOR.' -ForegroundColor Red
    Write-Host 'Abra o PowerShell como administrador e rode de novo:' -ForegroundColor Yellow
    Write-Host ('  powershell -ExecutionPolicy Bypass -File "' `
      + (Join-Path $Repo 'servidor\vigia-riscos.ps1') + '" -Agendar')
    exit 1
  }
  # O .vbs e ESCRITO AQUI, e nao versionado: servidor/tls esta no .gitignore
  # (guarda chaves), entao um clone novo do repositorio nao o traria e a tarefa
  # apontaria para um arquivo que nao existe -- falhando em silencio, que e o
  # feitio de defeito que este script inteiro existe para evitar.
  $vbs = Join-Path $Repo 'servidor\tls\vigia-riscos-oculto.vbs'
  $ps1 = Join-Path $Repo 'servidor\vigia-riscos.ps1'
  New-Item -ItemType Directory -Force -Path (Split-Path $vbs) | Out-Null
  $abre = 'CreateObject("Wscript.Shell").Run "powershell -NoProfile -ExecutionPolicy Bypass -File "'
  $fecha = '", 0, False'
  Set-Content -Path $vbs -Value ($abre + '"' + $ps1 + '""' + $fecha) -Encoding ascii
  $acao = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $vbs + '"')
  # No logon (o servidor entra sozinho) E de cinco em cinco minutos: o logon
  # cobre o reboot, a repeticao cobre o dia. A varredura custa milissegundos.
  $t1 = New-ScheduledTaskTrigger -AtLogOn
  $t2 = New-ScheduledTaskTrigger -Once -At 00:03 `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
  Register-ScheduledTask -TaskName 'Gerador-OS Vigia Riscos' `
    -Action $acao -Trigger $t1,$t2 -Force | Out-Null
  $t = Get-ScheduledTask -TaskName 'Gerador-OS Vigia Riscos'
  Write-Host ("Tarefa 'Gerador-OS Vigia Riscos' registrada - estado " + $t.State) -ForegroundColor Green
  Write-Host 'Roda no logon e a cada 5 minutos. Para rodar agora:' -ForegroundColor Gray
  Write-Host '  Start-ScheduledTask "Gerador-OS Vigia Riscos"' -ForegroundColor Gray
  exit 0
}

$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) {
  # Sem node nao ha o que fazer, e ficar calado seria o mesmo defeito de novo.
  Anotar 'node nao encontrado no PATH - o indice de riscos NAO esta sendo mantido'
  exit 1
}

$script = Join-Path $Repo 'servidor\indexar-riscos.js'
if (-not (Test-Path $script)) {
  Anotar "nao achei $script - o indice de riscos NAO esta sendo mantido"
  exit 1
}

# O indexador devolve 10 quando reescreveu a lista e 0 quando ja estava em dia.
# E a saida dele que diz QUAL arquivo entrou ou saiu, e e isso que vale guardar:
# "272 PDFs" nao ajuda ninguem tres dias depois.
Push-Location $Repo
$saida = & $node $script '--silencioso' 2>&1
$codigo = $LASTEXITCODE
Pop-Location

if ($codigo -eq 10) {
  Anotar 'a lista de riscos mudou:'
  foreach ($l in $saida) { Anotar ("   " + [string]$l) }
  Anotar 'ja vale na fabrica; para a copia da nuvem, falta commit de dados/riscos-pdf.json'
  Write-Host ''
  Write-Host 'Indice de riscos atualizado.' -ForegroundColor Yellow
} elseif ($codigo -eq 0) {
  Write-Host 'Nada a fazer: a lista de riscos ja estava em dia.' -ForegroundColor Green
} else {
  Anotar ("o indexador falhou (codigo $codigo): " + ($saida -join ' | '))
  Write-Host 'O indexador falhou - veja o log.' -ForegroundColor Red
  exit 1
}

<#
  AGENDAR (uma vez), no PowerShell aberto como ADMINISTRADOR:

    powershell -ExecutionPolicy Bypass -File "C:\Users\Pichau\Desktop\Gerador-OS\servidor\vigia-riscos.ps1" -Agendar

  Fica no logon (cobre o reboot) e de cinco em cinco minutos (cobre o dia). O
  .vbs existe so para a janela do PowerShell nao piscar na tela do servidor a
  cada passagem -- e o mesmo arranjo do Vigia Docker e do Vigia Rede.
#>
