<#
  Backup diario do servidor da fabrica — o que o Agendador de Tarefas chama.

  POR QUE EXISTE, em vez de o agendador chamar o node direto:
  um backup agendado que falha em silencio e pior do que nao ter backup, porque
  parece que voce esta coberto. Este involucro grava SEMPRE uma linha de log,
  deu certo ou nao, com o tamanho do arquivo que saiu. Assim da para responder
  "quando foi o ultimo backup bom?" sem abrir o Google Drive.

  Registrar a tarefa (uma vez, no servidor):
    .\servidor\backup-diario.ps1 -Senha 'a-senha' -Agendar

  Rodar na mao, para testar:
    .\servidor\backup-diario.ps1 -Senha 'a-senha'
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string] $Senha,
  [string] $Docker  = 'C:\supabase\docker',
  [string] $Destino = 'J:\Meu Drive\Backup Gerador-OS',
  [int]    $Manter  = 14,
  [string] $Hora    = '12:30',
  # Pastas de PDF que o pacote cifrado NAO leva e que so existiam neste disco.
  # Caminhos relativos a raiz do projeto.
  [string[]] $Desenhos = @('Desenhos técnicos', 'Desenhos técnicos -grades de corte'),
  [switch] $Agendar
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\backup-diario.log'   # tls/ e ignorado pelo git

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  try { Add-Content -Path $Log -Value $linha -Encoding utf8 } catch { }
  Write-Host $linha
}

# ------------------------------------------------------------------- agendar
if ($Agendar) {
  $nome = 'Backup Gerador-OS'
  $args = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', ('"' + (Join-Path $PSScriptRoot 'backup-diario.ps1') + '"'),
    # ASPAS DUPLAS, nao simples. O agendador chama "powershell.exe -File", e nesse
    # modo o PowerShell tira as aspas DUPLAS e deixa as SIMPLES dentro do valor.
    # Com aspas simples a senha do pacote virava 'a-senha' com as aspas coladas:
    # quem restaurasse digitaria a senha anotada e ouviria "senha errada", no pior
    # dia possivel. Descoberto em 14/08/2026 conferindo um pacote recem-gerado —
    # os pacotes de 10/08 a 14/08 so abrem com as aspas.
    '-Senha', ('"' + $Senha + '"'),
    '-Docker', ('"' + $Docker + '"'),
    '-Destino', ('"' + $Destino + '"'),
    '-Manter', $Manter
  ) -join ' '

  $acao = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $args -WorkingDirectory $Raiz
  $quando = New-ScheduledTaskTrigger -Daily -At $Hora

  # NAO roda "esteja o usuario logado ou nao", de proposito. O pg_dump precisa do
  # Docker Desktop, que no Windows so existe DENTRO da sessao do usuario; e o
  # destino e o Google Drive, que so aparece como J: na sessao. Fora dela as duas
  # coisas somem e o backup falharia todo dia, com log dizendo o porque.
  # TENTAR DE NOVO. A falha tipica nao e do backup: e da HORA. Numa manha de
  # recuperacao o Docker leva minutos para voltar (ja levou 5m18s), e o
  # StartWhenAvailable dispara a execucao perdida justamente ai — o pg_dump nao
  # acha o banco e a tarefa morre. Sem retry, o dia inteiro fica sem pacote e
  # ninguem percebe ate precisar restaurar. Tres tentativas a cada 30 min cobrem
  # a subida mais lenta ja medida, e o intervalo e largo porque cada execucao
  # leva minutos e move ~64 MB para o Google Drive.
  $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries `
            -AllowStartIfOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
            -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)

  try {
    Unregister-ScheduledTask -TaskName $nome -Confirm:$false -ErrorAction SilentlyContinue
    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $quando -Settings $conf `
      -Description 'Pacote de recuperacao do servidor da fabrica, cifrado, no Google Drive.' `
      -ErrorAction Stop | Out-Null
    Anotar "tarefa '$nome' agendada para todo dia as $Hora"
  } catch {
    Anotar "FALHA ao agendar: $($_.Exception.Message)"
    exit 1
  }
  exit 0
}

# -------------------------------------------------------------------- rodar
$script = Join-Path $PSScriptRoot 'backup-servidor.js'
if (-not (Test-Path $script)) { Anotar "FALHA: nao achei $script"; exit 1 }

# O destino e o Google Drive: se a unidade nao estiver montada, gravar criaria
# uma pasta local com o mesmo nome e o backup ficaria DENTRO do PC que ele
# deveria proteger — parecendo que deu certo.
$raizDestino = [System.IO.Path]::GetPathRoot($Destino)
if ($raizDestino -and -not (Test-Path $raizDestino)) {
  Anotar "FALHA: a unidade $raizDestino nao esta disponivel (Google Drive fora do ar?). Nada foi gravado."
  exit 1
}

$falhou = $false

# ---------------------------------------- 1) pacote cifrado do servidor inteiro
$saida = & node $script --docker $Docker --destino $Destino --senha $Senha --manter $Manter 2>&1 | Out-String
$codigo = $LASTEXITCODE

if ($codigo -eq 0) {
  $ultimo = Get-ChildItem $Destino -Filter '*.bkp' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($ultimo) {
    $mb = [math]::Round($ultimo.Length / 1MB, 1)
    Anotar "ok  $($ultimo.Name)  $mb MB"
  } else {
    Anotar 'ATENCAO: o script disse ok, mas nao achei nenhum .bkp no destino'
    $falhou = $true
  }
} else {
  # Parenteses obrigatorios: "Anotar 'x' + $y" o PowerShell le como Anotar 'x' e
  # depois soma o resto no vazio - o log ficaria dizendo que falhou sem dizer
  # por que, que e justamente a informacao que faz o log existir.
  $motivo = ($saida.Trim() -replace '\s+', ' ')
  if ($motivo.Length -gt 400) { $motivo = $motivo.Substring(0, 400) + '…' }
  Anotar ("FALHA (codigo $codigo): " + $motivo)
  $falhou = $true
}

# ------------------------------------------- 2) PDFs dos desenhos e mapas de corte
#
# POR QUE ISTO EXISTE:
# o pacote cifrado leva banco, imagens do Storage e .env. NAO leva estas pastas —
# 160 PDFs, ~100 MB de desenhos tecnicos e mapas de corte que so existiam neste
# disco (e nem todos no git). Se o disco morresse, iam junto.
#
# NAO usa /MIR de proposito. /MIR apaga no destino o que sumiu na origem, e isso
# e veneno num backup: uma pasta renomeada ou um arquivo apagado por engano seria
# apagado da copia de resgate tambem, que e exatamente de onde se ia buscar. O
# preco e acumular duplicata quando uma pasta e renomeada — barato perto de
# perder o desenho. Limpar sobra e trabalho manual, e tem de ser.
#
# Roda MESMO SE o pacote acima falhar: sao duas protecoes independentes, e uma
# cair nao e motivo para a outra nem ser tentada.
$destinoDesenhos = Join-Path $Destino 'Desenhos'
foreach ($pasta in $Desenhos) {
  $origem = Join-Path $Raiz $pasta
  if (-not (Test-Path $origem)) { Anotar "desenhos: pulei '$pasta' (nao existe aqui)"; continue }
  $alvo = Join-Path $destinoDesenhos $pasta

  # /NFL /NDL /NJH /NJS /NP: sem listar arquivo por arquivo. O log e para
  # responder "o backup rodou?", nao para despejar 160 linhas por dia.
  & robocopy $origem $alvo /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
  $rc = $LASTEXITCODE

  # Robocopy nao usa 0 para sucesso: 0 = nada a copiar, 1 = copiou, 2 = extras no
  # destino, 3 = os dois. De 8 para cima e que e erro de verdade. Tratar como
  # comando comum daria "falhou" todo dia em que copiasse alguma coisa.
  if ($rc -ge 8) {
    Anotar "FALHA ao copiar desenhos de '$pasta' (robocopy $rc)"
    $falhou = $true
  } else {
    # Conta TUDO, nao so PDF. A primeira versao filtrava '*.pdf' e escreveu
    # "0 PDFs" para a pasta de desenhos, que tem 30 PNG e 46 MB: a copia estava
    # certa e o log dizia zero. Num log de backup, zero se le como falha — e um
    # log que assusta a toa e tao ruim quanto um que esconde problema.
    $arquivos = @(Get-ChildItem $alvo -Recurse -File -ErrorAction SilentlyContinue)
    $mbD = [math]::Round((($arquivos | Measure-Object Length -Sum).Sum) / 1MB, 1)
    Anotar "ok  desenhos '$pasta' -> $($arquivos.Count) arquivos, $mbD MB no Drive"
  }
}

if ($falhou) { exit 1 }
exit 0
