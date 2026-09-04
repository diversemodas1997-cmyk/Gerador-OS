# Os icones do atalho na tela de inicio do celular.
#
# Por que um script e nao quatro PNGs soltos no repositorio: icone e a unica
# imagem do programa que nao vem de fora -- se algum dia a marca mudar, o
# certo e mudar aqui e rodar de novo, nao abrir um editor e adivinhar a cor.
#
# Rodar, da pasta do Gerador-OS:
#   powershell -ExecutionPolicy Bypass -File servidor\gerar-icones.ps1
#
# Sai em icones\: 192 e 512 (Android/manifest), 180 (iPhone) e 32 (aba do
# navegador). O 512 e tambem o "maskable" do Android: por isso o "OS" ocupa so
# o miolo -- o Android recorta a borda em circulo conforme o aparelho, e o que
# encostar na beirada some.

Add-Type -AssemblyName System.Drawing

$raiz = Split-Path -Parent $PSScriptRoot
$saida = Join-Path $raiz 'icones'
if (-not (Test-Path $saida)) { New-Item -ItemType Directory -Path $saida | Out-Null }

# As mesmas cores do programa (ver :root no styles.css).
$tinta = [System.Drawing.ColorTranslator]::FromHtml('#1a1a1a')
$papel = [System.Drawing.ColorTranslator]::FromHtml('#f5f2ea')
$ouro  = [System.Drawing.ColorTranslator]::FromHtml('#c9a961')

function Novo-Icone([int]$S, [string]$arquivo) {
  $bmp = New-Object System.Drawing.Bitmap($S, $S)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

  $g.Clear($tinta)

  # "OS" no miolo, monoespacado como todo numero do programa.
  $tamFonte = [float]($S * 0.30)
  $fonte = New-Object System.Drawing.Font('Consolas', $tamFonte, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $pincel = New-Object System.Drawing.SolidBrush($papel)
  $fmt = New-Object System.Drawing.StringFormat
  $fmt.Alignment = [System.Drawing.StringAlignment]::Center
  $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
  $caixa = New-Object System.Drawing.RectangleF(0, [float](-$S * 0.045), [float]$S, [float]$S)
  $g.DrawString('OS', $fonte, $pincel, $caixa, $fmt)

  # A barra dourada embaixo -- o mesmo acento do cabecalho das telas.
  $barraL = [float]($S * 0.30)
  $barraA = [float][Math]::Max(2, [Math]::Round($S * 0.045))
  $barra = New-Object System.Drawing.SolidBrush($ouro)
  $g.FillRectangle($barra, [float](($S - $barraL) / 2), [float]($S * 0.655), $barraL, $barraA)

  $caminho = Join-Path $saida $arquivo
  $bmp.Save($caminho, [System.Drawing.Imaging.ImageFormat]::Png)

  $g.Dispose(); $bmp.Dispose(); $fonte.Dispose(); $pincel.Dispose(); $barra.Dispose(); $fmt.Dispose()
  Write-Output ("  {0,-16} {1}x{1}" -f $arquivo, $S)
}

Write-Output 'Gerando os icones em icones\ :'
Novo-Icone 512 'icone-512.png'
Novo-Icone 192 'icone-192.png'
Novo-Icone 180 'icone-180.png'   # iPhone: o apple-touch-icon nao le o manifest
Novo-Icone 32  'favicon-32.png'
Write-Output 'Pronto.'
