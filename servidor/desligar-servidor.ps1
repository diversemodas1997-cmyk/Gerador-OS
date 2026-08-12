<#
  Desliga o PC do servidor sem arrancar o banco no meio de uma escrita.

  POR QUE EXISTE:
  em 10/08/2026, as 17:37, o PC foi desligado pelo menu do Windows. O logoff
  matou o Docker Desktop, que arrancou a maquina virtual do WSL com o Postgres
  escrevendo — e o log do banco registrou, 4 segundos depois:

      PANIC: could not write to log file ... Input/output error

  Daquela vez o Postgres se recuperou sozinho relendo o WAL e nada se perdeu.
  E sorte, nao garantia: e a mesma sorte que se pede toda noite se o
  desligamento continuar sendo pelo menu.

  Este script pede aos conteineres que parem com calma (o Postgres fecha os
  arquivos e escreve o checkpoint), fecha o Docker Desktop com jeito e SO DEPOIS
  desliga o Windows. O passo do Docker Desktop entrou em 12/08/2026, quando a
  fabrica perdeu uma manha: os conteineres tinham parado direitinho, mas o
  -Force do desligamento matou o Docker Desktop escrevendo a configuracao dele.

  Usar (o jeito de todo dia): o atalho "Desligar o servidor" na Area de Trabalho.
    .\servidor\desligar-servidor.ps1 -CriarAtalho     # cria o atalho, uma vez
    .\servidor\desligar-servidor.ps1                  # para tudo e desliga
    .\servidor\desligar-servidor.ps1 -NaoDesligar     # so para, para testar
#>
[CmdletBinding()]
param(
  [string] $Docker      = 'C:\supabase\docker',
  [int]    $Paciencia   = 60,
  [switch] $NaoDesligar,
  [switch] $CriarAtalho
)

$Raiz = Split-Path -Parent $PSScriptRoot
$Log  = Join-Path $Raiz 'servidor\tls\desligar-servidor.log'   # tls/ e ignorado pelo git

function Anotar($texto) {
  $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $texto
  try { Add-Content -Path $Log -Value $linha -Encoding utf8 } catch { }
  Write-Host $linha
}

# -------------------------------------------------------------- criar atalho
if ($CriarAtalho) {
  $ps1    = Join-Path $PSScriptRoot 'desligar-servidor.ps1'
  $atalho = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Desligar o servidor.lnk'
  try {
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut($atalho)
    $s.TargetPath       = 'powershell.exe'
    $s.Arguments        = '-NoProfile -ExecutionPolicy Bypass -File "' + $ps1 + '" -Docker "' + $Docker + '"'
    $s.WorkingDirectory = $Raiz
    $s.IconLocation     = 'shell32.dll,27'
    $s.Description      = 'Para o banco com seguranca e SO ENTAO desliga o servidor.'
    $s.Save()
    Anotar "atalho criado: $atalho"
  } catch {
    Anotar "FALHA ao criar o atalho: $($_.Exception.Message)"
    exit 1
  }
  exit 0
}

# --------------------------------------------------------------------- parar
# CUIDADO ao renomear: esta variavel NAO pode se chamar $docker. O parametro
# $Docker acima guarda a PASTA do Supabase, e no PowerShell $docker e $Docker sao
# a MESMA variavel. Pior: como o parametro e [string], atribuir o objeto do
# Get-Command nele nao da erro — o PowerShell converte para o texto "docker.exe"
# calado, .Source vira vazio e a pasta do Supabase se perde.
$achado = Get-Command docker.exe -ErrorAction SilentlyContinue
if ($achado) { $dockerExe = $achado.Source } else { $dockerExe = 'C:\Program Files\Docker\Docker\resources\bin\docker.exe' }

if (-not (Test-Path $dockerExe)) {
  # Sem Docker nao ha o que parar com cuidado, e travar o desligamento por causa
  # disso seria pior do que deixar desligar.
  Anotar 'nao achei o docker.exe — nada a parar'
} else {
  & $dockerExe info 2>$null | Out-Null
  if ($LASTEXITCODE -ne 0) {
    Anotar 'o motor do Docker ja estava fora do ar — nada a parar'
  } else {
    # O app primeiro: com o nginx fora, ninguem manda pedido novo para uma API
    # que esta fechando. Depois o Supabase, que o proprio compose desmonta na
    # ordem inversa das dependencias — o banco por ultimo, que e o que importa.
    $pilhas = @(
      @{ nome = 'app';      arquivo = (Join-Path $PSScriptRoot 'docker-compose.app.yml') },
      @{ nome = 'supabase'; arquivo = (Join-Path $Docker 'docker-compose.yml') }
    )
    foreach ($p in $pilhas) {
      if (-not (Test-Path $p.arquivo)) { Anotar ("pulei a pilha '" + $p.nome + "': nao achei " + $p.arquivo); continue }
      $saida = & $dockerExe compose -f $p.arquivo stop -t $Paciencia 2>&1 | Out-String
      if ($LASTEXITCODE -eq 0) {
        Anotar ("pilha '" + $p.nome + "' parada com calma")
      } else {
        $motivo = ($saida.Trim() -replace '\s+', ' ')
        if ($motivo.Length -gt 300) { $motivo = $motivo.Substring(0, 300) + '…' }
        Anotar ("ATENCAO: a pilha '" + $p.nome + "' nao parou direito: " + $motivo)
      }
    }

    # Conferencia que vale a pena: o Postgres fechou em paz ou foi morto no soco?
    #
    # A primeira versao procurava "database system is shut down" no docker logs.
    # NAO FUNCIONA: nesta imagem as mensagens de arranque e parada do Postgres nao
    # chegam ao docker logs de forma confiavel (so os FATAL e o PANIC chegam), e a
    # conferencia acusava fechamento sujo TODA VEZ. Um alarme que dispara todo dia
    # sem motivo e pior do que nenhum alarme: em uma semana ninguem le mais o log.
    #
    # O codigo de saida do conteiner e objetivo: 0 = saiu por conta propria depois
    # do SIGINT (fast shutdown, que e o STOPSIGNAL desta imagem); 137 = levou
    # SIGKILL por estourar a paciencia, que e exatamente o que se quer evitar.
    $codigo = (& $dockerExe inspect supabase-db --format '{{.State.ExitCode}}' 2>$null | Select-Object -First 1)
    if ($codigo -eq '0') {
      Anotar 'Postgres fechou em paz (saida 0)'
    } else {
      Anotar ("ATENCAO: o Postgres saiu com codigo $codigo — nao fechou em paz. " +
              "O proximo arranque vai reler o WAL. Aumentar -Paciencia se repetir.")
    }
  }
}

# ------------------------------------------------------------------ desligar
if ($NaoDesligar) { Anotar 'parei por aqui (-NaoDesligar): o Windows continua ligado'; exit 0 }

# Fechar o Docker Desktop com jeito ANTES de desligar.
#
# POR QUE: em 12/08/2026 a fabrica ficou sem o programa das 7h as 8h. Os
# conteineres tinham parado com calma na noite anterior — o Postgres saiu com
# codigo 0, o log esta ali — mas o Stop-Computer -Force logo em seguida matou o
# Docker Desktop no meio de uma escrita. Tres arquivos de configuracao dele
# ficaram cheios de bytes zero (o NTFS salvou o tamanho, nao o conteudo) e no dia
# seguinte o motor simplesmente nao subia. Parar os conteineres nao basta: o
# proprio Docker Desktop precisa fechar antes de a energia cair.
if (Test-Path $dockerExe) {
  Anotar 'pedindo ao Docker Desktop que feche'
  try { & $dockerExe desktop stop 2>&1 | Out-Null } catch { }

  $limite = (Get-Date).AddSeconds(90)
  $saiu   = $false
  while ((Get-Date) -lt $limite) {
    if (-not (Get-Process -Name 'Docker Desktop', 'com.docker.backend' -ErrorAction SilentlyContinue)) { $saiu = $true; break }
    Start-Sleep -Seconds 3
  }

  if ($saiu) {
    Anotar 'Docker Desktop fechado'
  } else {
    # Nao vale travar o desligamento para sempre por causa disso: matar aqui e
    # pior do que fechar com jeito, mas melhor do que o PC passar a noite ligado.
    Anotar 'ATENCAO: o Docker Desktop nao fechou em 90s — encerrando na marra'
    Get-Process -Name 'Docker Desktop', 'com.docker.backend' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  }

  # Um respiro para o Windows terminar de gravar o que ficou em cache. Sao 3
  # segundos que compram a manha seguinte.
  Start-Sleep -Seconds 3
}

Anotar 'desligando o Windows'
Stop-Computer -Force
