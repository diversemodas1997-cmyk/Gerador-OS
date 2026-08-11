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
    '-Senha', ("'" + $Senha + "'"),
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
  $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries `
            -AllowStartIfOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)

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
    exit 1
  }
} else {
  # Parenteses obrigatorios: "Anotar 'x' + $y" o PowerShell le como Anotar 'x' e
  # depois soma o resto no vazio - o log ficaria dizendo que falhou sem dizer
  # por que, que e justamente a informacao que faz o log existir.
  $motivo = ($saida.Trim() -replace '\s+', ' ')
  if ($motivo.Length -gt 400) { $motivo = $motivo.Substring(0, 400) + '…' }
  Anotar ("FALHA (codigo $codigo): " + $motivo)
  exit 1
}
