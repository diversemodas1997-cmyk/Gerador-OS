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

  COMMIT E PUSH DOS RISCOS (04/11/2026, Junior: "automatizar o processo de
  envio de commit e push no github quando usuario criar nova grade a partir do
  pdf, sem precisar acessar e enviar comando no claude code"). Na fabrica o
  efeito ja era imediato (o nginx serve esta pasta direto); o que dependia de
  alguem era a copia do GitHub. Agora, depois do indexador, o vigia commita e
  envia -- e SO ISSO, com estas travas:
    - o commit leva apenas a pasta de riscos e dados/riscos-pdf.json
      (`git commit -- <caminhos>`): codigo em edicao no projeto, staged ou nao,
      fica de fora;
    - arquivo mexido ha menos de $QuietoMin minutos = copia em andamento: espera
      a proxima passagem;
    - git ocupado (index.lock, merge, rebase) ou fora do branch main: pula;
    - PDF acima de $MaxMB MB (o GitHub recusa a partir de 100): nao commita
      nada e avisa no log;
    - push que falha fica PENDENTE (tls/vigia-riscos-push-pendente) e e
      tentado de novo a cada passagem. O push leva o branch inteiro: commits
      locais que estejam esperando vao junto.
  Apagar ou mover PDF tambem entra no commit; o historico do git guarda o
  arquivo, e e de la que se volta.

  RODAR A MAO:   powershell -File servidor\vigia-riscos.ps1
  AGENDAR:       powershell -File servidor\vigia-riscos.ps1 -Agendar  (como admin)
#>

[CmdletBinding()]
param(
  [string]$Repo = 'C:\Users\Pichau\Desktop\Gerador-OS',
  [string]$Log  = 'C:\Users\Pichau\Desktop\Gerador-OS\servidor\tls\vigia-riscos.log',
  # Registra a tarefa agendada e sai. Precisa de ADMINISTRADOR -- registrar
  # tarefa e a unica parte disto que o Windows nao deixa fazer de usuario comum.
  [switch]$Agendar,
  # Sem commit/push nesta passagem (so o indice, como era ate 04/11/2026).
  [switch]$SemPublicar,
  [int]$QuietoMin = 2,
  [int]$MaxMB = 95
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
  Write-Host ''
  Write-Host 'Indice de riscos atualizado.' -ForegroundColor Yellow
} elseif ($codigo -eq 0) {
  Write-Host 'A lista de riscos ja estava em dia.' -ForegroundColor Green
} else {
  Anotar ("o indexador falhou (codigo $codigo): " + ($saida -join ' | '))
  Write-Host 'O indexador falhou - veja o log.' -ForegroundColor Red
  exit 1
}

if ($SemPublicar) { exit 0 }

# ---------------------------------------------------------------------------
# COMMIT E PUSH (ver o cabecalho). Tudo em ASCII: o PowerShell 5.1 le este
# arquivo sem BOM como ANSI, entao o nome da pasta ("tecnicos" com acento) e
# achado no disco, nunca escrito aqui.
# ---------------------------------------------------------------------------
$git = (Get-Command git -ErrorAction SilentlyContinue).Source
if (-not $git) { Anotar 'git nao encontrado no PATH - riscos NAO estao indo para o GitHub'; exit 1 }
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
# Janela oculta: nada de pedir senha numa tela que ninguem ve (travaria a tarefa).
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'never'
$pend = Join-Path $Repo 'servidor\tls\vigia-riscos-push-pendente'

function G { & $git -C $Repo -c core.quotepath=false @args 2>&1; }

$dirGit = Join-Path $Repo '.git'
foreach ($ocupado in 'index.lock', 'MERGE_HEAD', 'rebase-merge', 'rebase-apply', 'CHERRY_PICK_HEAD') {
  if (Test-Path (Join-Path $dirGit $ocupado)) { Write-Host "git ocupado ($ocupado) - fica para a proxima passagem"; exit 0 }
}
$branch = ([string](G rev-parse --abbrev-ref HEAD)).Trim()
if ($branch -ne 'main') { Write-Host "fora do main ($branch) - nao publica"; exit 0 }

$pasta = (Get-ChildItem -LiteralPath $Repo -Directory | Where-Object { $_.Name -like 'Desenhos t*cnicos -grades de corte' } | Select-Object -First 1).Name
if (-not $pasta) { Anotar 'nao achei a pasta de riscos - nada publicado'; exit 1 }
$caminhos = @($pasta, 'dados/riscos-pdf.json')

# So as linhas de status de verdade ("XY caminho"): os avisos do git (LF/CRLF...)
# chegam pela mesma saida e nao sao arquivo.
$status = @(G status --porcelain -uall -- @caminhos | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^[ MADRCU?]{2} ' })
if ($status.Count) {
  # "XY caminho" (ou "R  velho -> novo"): o que vale e o caminho de agora.
  $itens = $status | ForEach-Object {
    $c = $_.Substring(3); if ($c -match ' -> ') { $c = ($c -split ' -> ')[-1] }
    [pscustomobject]@{ St = $_.Substring(0, 2); Caminho = $c.Trim('"') }
  }
  $agora = Get-Date
  foreach ($i in $itens) {
    $f = Join-Path $Repo $i.Caminho
    if (-not (Test-Path -LiteralPath $f)) { continue }   # apagado: nada a medir
    $a = Get-Item -LiteralPath $f
    # O indice e o proprio indexador que acabou de escrever: so os arquivos da
    # pasta de riscos podem estar no meio de uma copia. Vale a MAIS RECENTE entre
    # modificacao e criacao: copiar pelo Explorer guarda a data de modificacao do
    # original (04/11/2026: PDF de 12:01 colado depois das 12:03), e quem diz a
    # hora da copia e a criacao.
    $mexido = @($a.LastWriteTime, $a.CreationTime) | Sort-Object | Select-Object -Last 1
    if ($i.Caminho -ne 'dados/riscos-pdf.json' -and ($agora - $mexido).TotalMinutes -lt $QuietoMin) {
      Write-Host ("copia em andamento (" + $i.Caminho + ") - fica para a proxima passagem"); exit 0
    }
    if ($a.Length -gt $MaxMB * 1MB) {
      Anotar ("NAO publicado: " + $i.Caminho + " tem " + [math]::Round($a.Length / 1MB) + " MB (limite $MaxMB MB do GitHub) - reduza o PDF")
      exit 1
    }
  }

  G add -A -- @caminhos | Out-Null
  $pdfs = @($itens | Where-Object { $_.Caminho -like '*.pdf' })
  $linhas = @($pdfs | ForEach-Object {
    $sinal = if ($_.St -match 'D') { '-' } else { '+' }
    '  ' + $sinal + ' ' + ($_.Caminho -replace ('^' + [regex]::Escape($pasta) + '/'), '')
  })
  $titulo = if ($pdfs.Count -eq 1) { 'Riscos: ' + ($linhas[0].Trim()) } else { 'Riscos: ' + $pdfs.Count + ' PDFs' }
  if (-not $pdfs.Count) { $titulo = 'Riscos: indice dados/riscos-pdf.json' }
  $msg = $titulo + "`n`nEnviado sozinho pelo Vigia Riscos (servidor/vigia-riscos.ps1).`n" + ($linhas -join "`n") + "`n"
  $arqMsg = Join-Path $Repo 'servidor\tls\vigia-riscos-commit.txt'
  [IO.File]::WriteAllText($arqMsg, $msg, $utf8)
  $r = G commit -q -F $arqMsg -- @caminhos
  if ($LASTEXITCODE -ne 0) { Anotar ('commit falhou: ' + ($r -join ' | ')); exit 1 }
  $hash = ([string](G rev-parse --short HEAD)).Trim()
  Anotar ("commit $hash - " + $titulo)
  foreach ($l in $linhas) { Anotar $l }
  Set-Content -LiteralPath $pend -Value $hash -Encoding ascii
}

if (Test-Path -LiteralPath $pend) {
  $r = G -c http.lowSpeedLimit=1000 -c http.lowSpeedTime=60 push -q origin main
  if ($LASTEXITCODE -eq 0) {
    Remove-Item -LiteralPath $pend -Force
    Anotar 'push ok - riscos no GitHub'
  } else {
    Anotar ('push falhou, tento de novo na proxima passagem: ' + (($r | ForEach-Object { [string]$_ }) -join ' | '))
  }
}

<#
  AGENDAR (uma vez), no PowerShell aberto como ADMINISTRADOR:

    powershell -ExecutionPolicy Bypass -File "C:\Users\Pichau\Desktop\Gerador-OS\servidor\vigia-riscos.ps1" -Agendar

  Fica no logon (cobre o reboot) e de cinco em cinco minutos (cobre o dia). O
  .vbs existe so para a janela do PowerShell nao piscar na tela do servidor a
  cada passagem -- e o mesmo arranjo do Vigia Docker e do Vigia Rede.
#>
