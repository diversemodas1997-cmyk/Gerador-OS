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
  # Sem -Senha, vale a senha guardada cifrada nesta maquina (ver "a senha" abaixo).
  [string] $Senha,
  [string] $Docker  = 'C:\supabase\docker',
  [string] $Destino = 'J:\Meu Drive\Backup Gerador-OS',
  [int]    $Manter  = 14,
  [string] $Hora    = '12:30',
  # Pastas de PDF que o pacote cifrado NAO leva e que so existiam neste disco.
  # Caminhos relativos a raiz do projeto.
  [string[]] $Desenhos = @('Desenhos técnicos', 'Desenhos técnicos -grades de corte'),
  [switch] $Agendar,
  [switch] $GravarSenha
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\backup-diario.log'   # tls/ e ignorado pelo git

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  try { Add-Content -Path $Log -Value $linha -Encoding utf8 } catch { }
  Write-Host $linha
}

# --------------------------------------------------------------------- senha
#
# A SENHA DO PACOTE SAIU DO XML DA TAREFA (10/09/2026, pedido do Junior).
#
# Ela protege o pacote NO GOOGLE DRIVE — quem senta neste PC ja tem tudo, porque
# o servidor faz logon automatico. Ainda assim, escrita em texto puro nos
# argumentos da tarefa ela tinha dois precos: qualquer coisa que exporte a
# tarefa (um print, um suporte remoto, o proprio backup do Windows) levava a
# senha junto; e lembrar dela dependia de UMA anotacao a mao — perdida a
# anotacao, perdido o pacote.
#
# Agora ela mora cifrada em %LOCALAPPDATA%\Gerador-OS\backup-senha.txt, pelo
# DPAPI do Windows: so a MESMA conta de usuario desta maquina consegue decifrar,
# e o arquivo copiado para outro PC nao serve para nada. E de proposito que ele
# NAO vai para o Google Drive: guardar a chave ao lado do cofre e nao ter cofre.
#
#   gravar (uma vez):  .\servidor\backup-diario.ps1 -Senha 'a-senha' -GravarSenha
#   rodar a mao:       .\servidor\backup-diario.ps1              (le a guardada)
#                      .\servidor\backup-diario.ps1 -Senha 'x'   (manda esta)
$ArquivoSenha = Join-Path (Join-Path $env:LOCALAPPDATA 'Gerador-OS') 'backup-senha.txt'

function GravarSenhaCifrada($texto) {
  $pasta = Split-Path -Parent $ArquivoSenha
  if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
  $cifrada = ConvertTo-SecureString $texto -AsPlainText -Force | ConvertFrom-SecureString
  Set-Content -Path $ArquivoSenha -Value $cifrada -Encoding ascii
}

function LerSenhaCifrada {
  if (-not (Test-Path $ArquivoSenha)) { return '' }
  try {
    # .Trim() obrigatorio: Set-Content grava uma quebra de linha no fim, e
    # ConvertTo-SecureString recusa a cadeia com ela ("nao estava em um formato
    # correto"). Sem o Trim o backup falhava dizendo que nao havia senha — com a
    # senha guardada ali do lado, que e o jeito mais rapido de perder uma tarde.
    $segura = (Get-Content $ArquivoSenha -Raw).Trim() | ConvertTo-SecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($segura)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
  } catch { return '' }
}

if ($GravarSenha) {
  if (-not $Senha) { Anotar 'FALHA: -GravarSenha precisa de -Senha'; exit 1 }
  GravarSenhaCifrada $Senha
  Anotar "senha do pacote guardada cifrada em $ArquivoSenha"
  exit 0
}

# ------------------------------------------------------------------- agendar
if ($Agendar) {
  $nome = 'Backup Gerador-OS'
  $args = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', ('"' + (Join-Path $PSScriptRoot 'backup-diario.ps1') + '"'),
    # A SENHA NAO ENTRA AQUI. Ate 10/09/2026 ela vinha como '-Senha "<senha>"' e
    # ficava legivel para qualquer um que exportasse a tarefa. Agora o script a
    # le do arquivo cifrado desta maquina (ver "senha", acima) — e some junto o
    # velho problema das aspas: em 14/08/2026 os pacotes de 10 a 14/08 so abriam
    # com a senha entre aspas simples, porque "powershell.exe -File" tira as
    # duplas e deixa as simples dentro do valor. Sem argumento, sem aspas.
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
# Sem -Senha na linha de comando, vale a guardada. Falhar AQUI, alto e no log, e
# melhor que cifrar o pacote com senha vazia e ninguem notar ate precisar dele.
if (-not $Senha) { $Senha = LerSenhaCifrada }
if (-not $Senha) {
  Anotar "FALHA: nao ha senha do pacote. Grave uma com: backup-diario.ps1 -Senha '<senha>' -GravarSenha"
  exit 1
}

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

# ------------------------------------------------- esperar o banco acordar
#
# POR QUE ESPERAR, EM VEZ DE FALHAR (15/09/2026, pedido do Junior).
#
# Em 14/09 as 07:15 o backup morreu com "container ... is not running": era a
# execucao de recuperacao do StartWhenAvailable, disparada quando a maquina
# ligou, e o Docker ainda estava subindo o Supabase. Nao havia nada de errado —
# so chegou cedo demais. A tentativa das 12:30 no mesmo dia deu certo, entao
# aquele dia nao ficou sem copia; mas depender do RestartCount para isso e
# apostar que a segunda tentativa vai cair numa hora melhor.
#
# A janela e conhecida e curta: o nginx sobe sozinho com o Docker e o Supabase
# so entra quando o vigia manda "up -d" — em 13/08/2026 foram quatro minutos de
# diferenca. Vinte minutos de paciencia cobrem isso com folga e cabem dentro do
# ExecutionTimeLimit de uma hora da tarefa.
#
# A pergunta e a CERTA: nao "o container existe?", e sim "o Postgres atende?".
# Um container de pe com o banco ainda recuperando faria o pg_dump falhar do
# mesmo jeito, e o log diria a mesma coisa.
function EsperarBanco([int] $LimiteSegundos = 1200, [int] $IntervaloSegundos = 15) {
  $inicio = Get-Date
  $avisou = $false
  while ($true) {
    # pg_isready devolve 0 quando o banco aceita conexao. A saida vai para o
    # vazio de proposito: o que interessa e o codigo, e o texto so sujaria o log.
    $null = & docker exec supabase-db pg_isready -U postgres -d postgres 2>$null
    if ($LASTEXITCODE -eq 0) {
      if ($avisou) {
        $s = [int]((Get-Date) - $inicio).TotalSeconds
        Anotar "banco respondeu depois de $s s de espera"
      }
      return $true
    }
    $decorrido = ((Get-Date) - $inicio).TotalSeconds
    if ($decorrido -ge $LimiteSegundos) {
      # Em minutos so quando ha minutos: um limite curto (num teste) dizendo
      # "0 min de espera" faria o log mentir sobre o que foi tentado.
      $quanto = if ($LimiteSegundos -ge 60) { "$([int]($LimiteSegundos / 60)) min" } else { "$LimiteSegundos s" }
      Anotar "FALHA: o banco nao respondeu em $quanto de espera. Nada foi exportado."
      return $false
    }
    if (-not $avisou) {
      # Sem acento e sem travessao: esta linha passa pelo console do Agendador
      # antes do log, e la a acentuacao vira caca (ver as linhas de 14/09).
      Anotar 'banco ainda nao responde (maquina recem-ligada?) - esperando'
      $avisou = $true
    }
    Start-Sleep -Seconds $IntervaloSegundos
  }
}

# ---------------------------------------- 1) pacote cifrado do servidor inteiro
if (-not (EsperarBanco)) {
  # Sai do passo 1 sem tentar o dump — mas os PDFs abaixo continuam, porque sao
  # protecao independente e nao dependem de banco nenhum.
  $falhou = $true
  $codigo = 1
  $saida = ''
} else {
$saida = & node $script --docker $Docker --destino $Destino --senha $Senha --manter $Manter 2>&1 | Out-String
$codigo = $LASTEXITCODE
}

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
