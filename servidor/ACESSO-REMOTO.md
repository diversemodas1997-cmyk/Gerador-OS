# Acesso de fora da fábrica (unidade São Carlos)

São Carlos vai **cadastrar e editar**, não só consultar. Isso decide o caminho:

- **Pela nuvem, não.** A nuvem é cópia de mão única (ver "Espelho para a nuvem"
  no `README.md`): o espelho reescreve a nuvem com o que está na fábrica a cada
  30 minutos. Qualquer coisa editada lá some na passada seguinte, sem erro na
  tela. E o projeto antigo da nuvem nem existe mais (DNS não resolve desde, pelo
  menos, 13/10/2026 — última execução do espelho, com falha).
- **Direto no servidor da fábrica, sim.** As máquinas de São Carlos entram numa
  rede privada junto com o servidor e passam a ser, para o programa, mais um
  computador da fábrica: mesma fonte da verdade, gravação normal, nada a
  conciliar.

A rede privada escolhida é o **Tailscale** (gratuito até 3 usuários / 100
aparelhos). Nada fica aberto para a internet: só aparelhos que você aprovar na
sua conta enxergam o servidor. Não precisa mexer no roteador.

## O que o programa já resolve sozinho

- **Qual servidor usar.** O app pergunta a quem serviu a página
  (`/servidor-local.json`) e usa esse endereço (`configDoServidorQueServiu` em
  `app.js`). Abrindo pelo endereço do Tailscale, ele fala com a fábrica pelo
  Tailscale — ninguém configura nada máquina a máquina.
- **Certificado.** `gerar-certificado.js` reemite o certificado com endereços a
  mais **sem trocar a autoridade** da fábrica: quem já instalou o `ca.crt`
  continua valendo.

## Passo a passo

### 1. No servidor da fábrica (este PC, DESKTOP-SOV61AF)

1. Instalar o Tailscale: <https://tailscale.com/download/windows> e entrar com a
   conta da empresa (a mesma será usada em São Carlos).
2. Anotar o endereço que ele dá ao servidor:
   ```powershell
   & "C:\Program Files\Tailscale\tailscale.exe" ip -4
   ```
   É um número `100.x.y.z`. No painel do Tailscale (Machines), desligar a
   expiração da chave deste aparelho (**Disable key expiry**) — senão ele sai da
   rede sozinho a cada 180 dias e São Carlos fica sem acesso.
3. Reemitir o certificado **somando** o endereço novo aos de sempre (troque
   `100.x.y.z` pelo número do passo 2):
   ```powershell
   cd C:\Users\Pichau\Desktop\Gerador-OS
   node servidor\gerar-certificado.js --ip 193.168.0.200,192.168.1.50,192.168.1.158,100.x.y.z --nome DESKTOP-SOV61AF
   docker restart gerador-os-web
   ```
   Os IPs da fábrica precisam continuar na lista — tirar um deles quebra quem
   abre por ele.
4. Firewall: o Tailscale cria uma rede própria no Windows. Se a porta 443 só
   estiver liberada no perfil "Privado", libere também para a interface do
   Tailscale. Teste do próprio servidor: `https://100.x.y.z` tem de abrir.

### 2. Em cada computador de São Carlos

1. Instalar o Tailscale e entrar com a **mesma conta** (ou aprovar o aparelho no
   painel).
2. Instalar a CA da fábrica: abrir `https://100.x.y.z/instalar/ca.crt` (ou usar
   `instalar-certificado.cmd` da pasta `instalar`). É o mesmo procedimento dos
   PCs da fábrica.
3. Abrir `https://100.x.y.z` no Chrome e instalar como app (menu ⋮ → Transmitir,
   salvar e compartilhar → Instalar página como app). Usar **só esse atalho**.
4. Conferir no canto da tela a tarja **🏭 Servidor da fábrica**. Se aparecer
   **☁ Nuvem**, o Tailscale não está conectado nessa máquina.

## O que muda no dia a dia

- **O servidor da fábrica precisa estar ligado** para São Carlos trabalhar — é a
  única cópia que aceita gravação. Ele já é o servidor de todo mundo; agora a
  queda dele também para São Carlos.
- **Internet da fábrica caiu:** a fábrica continua trabalhando; São Carlos para
  até voltar. A tela de login diz "O servidor da fábrica não respondeu. Não é a
  sua senha."
- **Backup:** não muda. O pacote cifrado (12:30) e o JSON (16:00) já guardam no
  Google Drive tudo o que São Carlos gravar, porque grava no mesmo banco.
- **A nuvem e o espelho** continuam parados (`Espelho Gerador-OS` desativado).
  Só voltam a fazer sentido para consulta por quem não tiver Tailscale; para
  isso seria preciso criar um projeto novo no Supabase e trocar o endereço em
  `app.js` (`SUPA_URL`/`SUPA_KEY`), `servidor/espelhar-para-nuvem.js`,
  `servidor/espelhar-contas.js` e `servidor/copiar-desenhos.js`.
