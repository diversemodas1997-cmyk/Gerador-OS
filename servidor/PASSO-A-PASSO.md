# O que falta depois do instalador — passo a passo

O servidor está no ar em **`https://193.168.0.200`**. O instalador fez tudo o que
dava para fazer sozinho. Sobraram quatro tarefas, e todas as quatro sobraram
pelo mesmo motivo: **exigem uma decisão sua ou uma volta pelas máquinas**.

Faça na ordem. A 2 tem de vir antes da 3: sem o certificado, o navegador da
máquina barra o endereço do servidor e você não chega na tela de configuração.

| # | tarefa | onde | quanto tempo |
|---|---|---|---|
| 1 | Criar as contas | só no servidor | ~2 min por pessoa |
| 2 | Instalar o `ca.crt` | em cada máquina | ~2 min por máquina |
| 3 | Apontar para o servidor | em cada máquina | **não é mais preciso** |
| 4 | Logon automático + no-break | só no servidor | ~10 min |

> **A tarefa 3 deixou de existir.** O servidor publica o próprio endereço e a
> chave em `/servidor-local.json`, e o programa aberto por ele se conecta
> sozinho. Abrir `https://193.168.0.200` já basta. O texto do passo 3 continua
> abaixo como referência, para quando alguém precisar conferir ou desfazer.

---

## 1. Criar as contas

As contas não vieram no backup — elas vivem na autenticação, não nos dados.
**Sem conta, a pessoa não entra; e sem entrar, não vê cadastro, OS nem OE
nenhuma.** Este é o passo que, esquecido, faz parecer que "a intranet não
funciona" quando está tudo certo.

### Dentro do próprio programa

Entre como admin e vá em **Configurações → Contas de acesso**. Ali você vê
**todas as contas que têm acesso** e cria as que faltam:

- **E-mail:** serve como nome de usuário. Pode ser o e-mail real ou um interno,
  tipo `maria@diverse.local` — nunca chega mensagem nele.
- **Senha:** o botão **Sugerir senha** gera uma legível, sem letras que se
  confundem ao ditar por telefone.
- **Pode editar?** "Não" é o normal: a pessoa consulta e imprime tudo, mas não
  altera. "Sim" dá poder de administrador, inclusive de mudar o papel dos
  outros.

A conta já nasce válida — **não há e-mail de confirmação para esperar**, e a
pessoa entra na hora.

> **Anote a senha antes de sair da tela.** Ela não fica guardada em lugar nenhum
> que dê para consultar depois. Se perder, crie outra pela mesma tela.

> **Por que isso não é feito direto pelo navegador.** Criar conta pelo navegador
> exigiria ligar o auto-cadastro do Supabase — e a chave pública está dentro de
> toda página aberta na fábrica. Com o auto-cadastro ligado, qualquer pessoa na
> rede criaria a própria conta e passaria a ler tudo: OS, custos, produção. Por
> isso o auto-cadastro fica **desligado**, e a criação passa por uma função do
> servidor que confere se quem pediu é admin. Só funciona pelo endereço da
> fábrica; pela nuvem, o botão avisa isso.

### Pelo painel do Supabase (alternativa)

Se preferir, ou se precisar do primeiro admin antes de conseguir entrar: no
**servidor**, abra `http://localhost:8000` (usuário `admin`, senha da linha
`DASHBOARD_PASSWORD` em `C:\supabase\docker\.env`).

**Authentication → Users → Add user → Create new user**, com **Auto Confirm
User ligado**.

> **Não use "Send invitation" ali.** Convite manda e-mail, e não há servidor de
> e-mail na fábrica — o convite nunca chega e a conta fica pendurada sem poder
> entrar.

Para dar admin à primeira conta, no **SQL Editor**:

```sql
INSERT INTO user_roles (user_id, role)
SELECT id, 'admin' FROM auth.users WHERE email = 'SEU-EMAIL-AQUI'
ON CONFLICT (user_id) DO UPDATE SET role = 'admin';
```

Deve responder `INSERT 0 1`. Se responder `INSERT 0 0`, o e-mail não bateu com
nenhuma conta — confira a grafia em Authentication → Users.

### O que cada tipo de conta enxerga

Quem não é admin vê **tudo o que é leitura**: todos os cadastros, a lista e a
folha de OS pronta para imprimir, a Expedição, a folha do plano (OE), estoque,
ranking e operações. Some do menu só o **Nova OS**, e os botões de criar,
editar e excluir de cada tela.

Consultar e imprimir funciona igual para todo mundo. O que muda é quem altera.

---

## 2. Instalar o `ca.crt` em cada máquina

### O que é esse arquivo

`ca.crt` é um arquivinho de 1 KB. Ele é o **crachá da fábrica**: não abre nada
e não tem senha dentro. Serve só para o Windows de cada máquina saber que
`https://193.168.0.200` é de confiança. Sem ele, o navegador mostra "a conexão
não é particular" e o programa não funciona direito.

Ele nasceu junto com o servidor e mora em
`C:\Users\Pichau\Desktop\Gerador-OS\servidor\tls\ca.crt` — Área de Trabalho,
pasta `Gerador-OS`, pasta `servidor`, pasta `tls`.

### Jeito mais rápido: o atalho pronto (sem digitar comando)

Quando não se acha o PowerShell na máquina, use o atalho. **Em cada computador:**

1. No navegador, abra **`https://193.168.0.200/instalar-cracha-fabrica.bat`**
   (dê **Avançado → Prosseguir** no aviso, como no `ca.crt`). Ele baixa para
   **Downloads**.
2. Na pasta Downloads, **clique com o botão direito** no
   `instalar-cracha-fabrica.bat` → **Executar como administrador** → **Sim**.
3. Ele baixa e instala o crachá sozinho. Quando disser **PRONTO**, feche e
   reabra o navegador e abra `https://193.168.0.200`.

O atalho faz exatamente o que os passos abaixo fazem à mão — serve para quem não
quer mexer no PowerShell. Se preferir o manual, siga a seguir.

### À mão, sem pendrive

O próprio servidor entrega o arquivo. **Em cada computador da fábrica:**

1. No navegador, digite **`https://193.168.0.200/ca.crt`**.
2. Vai aparecer "a conexão não é particular" — é esperado, é exatamente o que
   estamos consertando. Clique em **Avançado** → **Prosseguir para
   193.168.0.200 (não seguro)**.
3. O arquivo baixa sozinho, para a pasta **Downloads**.
4. Abra o PowerShell **como administrador** (botão direito no menu Iniciar →
   "Windows PowerShell (Admin)") e rode:

```powershell
Import-Certificate -FilePath "$env:USERPROFILE\Downloads\ca.crt" -CertStoreLocation Cert:\LocalMachine\Root
```

O caminho da pasta Downloads já está embutido no comando — é colar e dar Enter,
sem navegar até pasta nenhuma.

> **Só o `ca.crt` é público.** Na mesma pasta do servidor mora o `ca.key`, que é
> segredo: quem tiver esse arquivo emite certificados em nome da sua autoridade.
> O servidor **nega** o acesso a ele pela rede, junto com os backups e o
> histórico do programa. Se algum dia você copiar essa pasta para outro lugar,
> lembre que o `ca.key` não pode ir junto.

Depois **feche e reabra o navegador**. Chrome e Edge usam a lista do Windows,
então já vale para os dois.

**Se alguém usa Firefox**, ele tem lista própria: Configurações → Privacidade e
Segurança → Certificados → Ver certificados → aba Autoridades → Importar →
escolher o `ca.crt` → marcar "Confiar nesta CA para identificar sites".

### Conferir

Abra `https://193.168.0.200` nessa máquina. O cadeado tem de aparecer **sem
aviso**. Se aparecer aviso de "conexão não é particular", o certificado não
entrou — refaça com administrador de verdade e reabra o navegador.

> Vale para o servidor também, se você for usar o programa nele.

> O certificado do servidor vale até **novembro de 2028**. Quando vencer, rode
> `node servidor\gerar-certificado.js --ip 193.168.0.200` de novo — a autoridade
> continua a mesma, então **não** é preciso passar nas máquinas outra vez.

---

## 3. Cada máquina, uma vez só — e depois ela se vira sozinha

O servidor tem **dois caminhos**: o cabo (`193.168.0.200`) e o Wi-Fi
(`192.168.1.158`). Ele troca de um para o outro quando o cabo cai. Em
**31/08/2026** isso parou a fábrica inteira numa segunda-feira: os atalhos
guardavam um endereço, o endereço morreu no fim de semana, e "tempo esgotado"
era toda a explicação.

Hoje **nenhuma máquina precisa saber em qual rede o servidor está.**

### 3.1 Instalar — uma vez por máquina

No servidor, monte o pacote (ele aparece na Área de Trabalho):

```powershell
.\servidor\preparar-instalador.ps1
```

Leve a **pasta inteira** (pen drive ou rede) e clique duas vezes em
`instalar-certificado.cmd` → **SIM** na permissão. Não separe os arquivos: o
instalador procura o `ca.crt` e o `abrir-gerador-os.cmd` ao lado dele.

Ele faz quatro coisas: instala o certificado da fábrica, copia o lançador para
`%ProgramData%\Gerador-OS`, cria o atalho **Gerador-OS** na Área de Trabalho de
todos os perfis, e **apaga o `.url` antigo** — que é a armadilha de endereço
fixo, e voltaria a ser.

### 3.2 Guardar as DUAS redes nessa máquina

No Windows, conecte-a uma vez ao **cabo** (`Diverse001`) e uma vez ao **Wi-Fi**
(`Desktop_F7027412`), deixando **"Conectar automaticamente"** marcado nos dois.

Este passo é o que ninguém lembra, e sem ele o resto não adianta: **o lançador
escolhe o caminho, não cria caminho.** Cliente e servidor precisam estar na
mesma rede. Se o servidor migrar para o Wi-Fi e a máquina só conhecer o cabo,
não há endereço no mundo que os junte.

### 3.3 Como a troca acontece sozinha — três camadas

Cada uma cobre a falha da anterior:

1. **O atalho procura.** `abrir-gerador-os.cmd` tenta, nesta ordem —
   `DESKTOP-SOV61AF` → `193.168.0.200` → `192.168.1.158` — e abre o primeiro que
   responde. Ele refaz a busca **a cada clique**, não só na instalação.
2. **O nome acompanha a rede.** Um número pertence a *uma* rede; um nome, não. O
   Windows resolve o nome da máquina sozinho (LLMNR/NetBIOS) na rede em que o
   **cliente** estiver. Os dois IPs ficam abaixo dele como rede de segurança,
   para a máquina onde a resolução de nome esteja desligada por política.
3. **O programa reaprende o endereço.** Se o endereço guardado nas Configurações
   daquela máquina não responder, o programa pergunta a quem serviu a página e
   passa a usar esse. É por isso que uma máquina que ficou presa num endereço
   velho se conserta sozinha, sem ninguém ir até ela.

O certificado cobre os três caminhos, com a **mesma autoridade** — por isso
reemiti-lo não obriga ninguém a reinstalar o `ca.crt`.

### 3.4 Configurações → "Servidor da fábrica": normalmente NÃO mexer

Vale para conferir ou corrigir à mão, numa máquina que ficou presa em algo
antigo. Vale **só naquela máquina**.

1. Abra o atalho **Gerador-OS**.
2. Entre com a conta da pessoa.
3. **Configurações** → cartão **Servidor da fábrica (rede local)**.
4. **Endereço:** `https://DESKTOP-SOV61AF` — sem porta, a API vem pelo mesmo
   endereço. Preferir o **nome** ao número: o número volta a envelhecer.
   **Chave pública (anon):** está em `servidor	lsesumo-instalacao.txt`.
5. **Testar conexão** → **Salvar e recarregar**.

A barra lateral deve mostrar **🏭 Servidor da fábrica**.

### 3.5 Quando não abre — o que a mensagem está dizendo

| O que aparece | O que é | O que fazer |
|---|---|---|
| O lançador lista os três e diz "não responde" | a máquina está em outra rede que o servidor | conferir o cabo / o Wi-Fi **desta** máquina (3.2) |
| "Sua conexão não é particular" (cadeado vermelho) | chegou ao servidor; falta o `ca.crt` **aqui** | rodar o instalador nesta máquina (3.1) |
| **☁ Nuvem** em vermelho na barra lateral | o servidor não respondeu; abriu a cópia da nuvem em **modo consulta** — vê e imprime, não edita | proposital, para os dois lados não ficarem diferentes; conferir servidor e rede |
| Abre no servidor e em **nenhuma** outra máquina | o Windows classificou a rede do servidor como **Pública**, e o perfil Público ignora TODA regra de entrada | no servidor: `.\servidor\liberar-portas-firewall.ps1` como administrador |

O último é o mais traiçoeiro: a regra do firewall aparece verde e correta, não
há erro nem log, a conexão simplesmente não chega — e testar do próprio servidor
**sempre dá certo**, porque o laço local não passa pelo firewall. A tarefa
**"Gerador-OS Vigia Rede"** (logon + 15 min) existe para que ele não volte, mas
**toda rede nova entra como Pública**: trocar de roteador o traz de volta.

---

## 4. Logon automático e no-break

**Esta é a tarefa mais importante das quatro**, e a mais fácil de adiar.

No Windows, o Docker Desktop roda dentro da sessão do usuário. Se a máquina
reiniciar de madrugada — atualização do Windows, queda de energia — e ninguém
fizer login, **o servidor não sobe e a fábrica abre sem sistema**. É a falha
mais provável de todo o conjunto.

### Configurar o logon automático

Nesta máquina a caixa comum do `netplwiz` está **escondida** (é o padrão do
Windows 10 recente). Há dois caminhos:

**Caminho recomendado — Autologon da Microsoft.** Guarda a senha cifrada, e não
em texto puro no registro:

1. Baixe <https://download.sysinternals.com/files/AutoLogon.zip> e descompacte.
2. Rode o `Autologon.exe`, aceite os termos.
3. Preencha **Username** (`Pichau`), **Domain** (`DESKTOP-SOV61AF`) e a senha.
4. Clique em **Enable**.

**Caminho alternativo — netplwiz.** Primeiro faça a caixa aparecer, no
PowerShell como administrador:

```powershell
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device" /v DevicePasswordLessBuildVersion /t REG_DWORD /d 0 /f
```

Depois rode `netplwiz`, desmarque **"Os usuários devem digitar um nome de
usuário e uma senha para usar este computador"**, clique em OK e informe a
senha duas vezes.

### O teste que vale

Configurar não prova nada. **Reinicie o servidor e não toque nele.** Depois de
uns três minutos, de outra máquina, abra `https://193.168.0.200`.

Se o app abrir, está resolvido. Se não abrir, o logon automático não pegou —
resolva agora, não depois.

### No-break

O servidor guarda os dados da fábrica. Queda de energia com o Postgres
escrevendo é a receita clássica de banco corrompido. Um no-break de linha já
cobre o caso comum, que é a oscilação curta.

---

## Se o servidor parar de responder na rede

**Sintoma:** o app abre no próprio servidor, mas nenhuma outra máquina alcança —
nem o navegador, nem o ping. O servidor continua navegando na internet
normalmente.

**Causa quase certa:** o Windows reclassificou a rede da fábrica como
**Pública**. O perfil Público está com *"Bloquear todas as conexões de entrada,
incluindo as da lista de aplicativos permitidos"*, e nesse modo ele **descarta
tudo o que chega e ignora todas as regras de permissão** — as do Gerador-OS
inclusive. As regras continuam lá, aparecem certas em qualquer conferência, e
não valem nada.

Isso aconteceu na instalação, em 10/08/2026, e custou uma hora de diagnóstico
em pistas erradas. O Windows reclassifica sozinho quando o roteador é trocado
ou reiniciado de um jeito que ele lê como rede nova.

**Conferir, em dez segundos:**

```powershell
Get-NetConnectionProfile | Select-Object InterfaceAlias, NetworkCategory
```

A placa do cabo (`Ethernet 3`) tem de estar **`Private`**. Se estiver `Public`,
é isso.

**Consertar** — PowerShell como administrador:

```powershell
Set-NetConnectionProfile -InterfaceAlias 'Ethernet 3' -NetworkCategory Private
```

Vale na hora, sem reiniciar nada.

> **Por que Privada e não destravar o perfil Público.** A rede da fábrica é
> interna, de equipamentos conhecidos — é o que "Privada" significa. Destravar
> o Público abriria a máquina em *qualquer* rede onde ela se conectasse, o que
> é bem diferente do que precisamos.

**Se não for isso**, o caminho que resolveu da última vez foi parar de deduzir e
ligar o registro do firewall, como administrador:

```powershell
netsh advfirewall set allprofiles logging droppedconnections enable
netsh advfirewall set allprofiles logging allowedconnections enable
```

Peça para alguém tentar acessar, e leia
`C:\Windows\system32\LogFiles\Firewall\pfirewall.log`. Uma linha `DROP TCP
<origem> <servidor> ... 443 ... RECEIVE` diz, sem margem para dúvida, que o
pedido chega e o Windows derruba. Nenhuma linha da máquina que tentou diz que o
pacote não chega, e aí o problema é de rede, não do servidor. Desligue depois
(`disable` no lugar de `enable`) — registrar tudo custa disco à toa.

---

## Conferência final

| o quê | como | esperado |
|---|---|---|
| Contas | Cada pessoa entra com a conta dela | entra e vê os cadastros |
| Seu admin | Menu mostra **Nova OS** | aparece só para você |
| Certificado | `https://193.168.0.200` em outra máquina | cadeado sem aviso |
| Arquivos servidos | `https://193.168.0.200/vendor-check.html` | 7 linhas "ok" |
| Pastas automáticas | Configurações → as três pastas | deixa conectar |
| Sem internet | Tire o cabo do roteador (não do switch) | app continua normal |
| Rede do cabo é Privada | `Get-NetConnectionProfile` | `Ethernet 3 → Private` |
| Sobe sozinho | Reinicie e espere | app volta sem ninguém tocar |

Os dois últimos são o motivo de tudo isto.

---

## Depois, com calma

- **Dados de 08 a 10/08.** O backup migrado é o de **07/08**. O que foi mexido
  depois disso não está no servidor.
- **O campo SKU perdeu o autocompletar.** A lista vem da tabela
  `skus_catalogo`, que o sistema de Estoque publica na nuvem e que não existe
  no servidor da fábrica. O campo aceita texto livre normalmente — só não
  sugere mais.
- **Espelho para a nuvem** — a receita está no fim do `README.md`.

---

## O backup diário (já configurado)

Roda **todo dia às 12:30**, pela tarefa `Backup Gerador-OS` do Windows. Gera um
pacote cifrado em `J:\Meu Drive\Backup Gerador-OS` com o banco inteiro (dados,
contas de login com as senhas, papéis), as imagens dos desenhos e o `.env` com
as chaves. Mantém 14 dias.

**Conferir se está rodando.** A pergunta certa não é "a tarefa existe?", é
"quando foi o último backup bom?":

```powershell
Get-Content servidor\tls\backup-diario.log -Tail 10
```

Cada execução escreve uma linha, deu certo ou não:

```
2026-08-10 17:31:57  ok  servidor-gerador-os-2026-08-10.bkp  50.7 MB
2026-08-10 17:30:53  FALHA: a unidade Q:\ nao esta disponivel (Google Drive fora do ar?)
```

**Por que 12:30 e não de madrugada.** O `pg_dump` precisa do Docker de pé, e o
Docker só existe dentro da sessão do Windows. De madrugada a máquina pode ter
reiniciado sem ninguém logar — e o backup falharia todo dia, em silêncio. Pelo
mesmo motivo a tarefa roda **só com a sessão aberta**: o `J:` do Google Drive
também só existe nela.

**Mudar o horário, ou refazer a tarefa:**

```powershell
.\servidor\backup-diario.ps1 -Senha 'a-senha' -Agendar -Hora 12:30
```

> **A senha fica legível no XML da tarefa**, para quem for administrador desta
> máquina. Como o servidor entra sozinho no Windows sem senha, quem chega nele
> fisicamente já tem tudo — então isto não abre uma porta que já não estivesse
> aberta. Mas vale saber o que a cifra protege: o pacote no Google Drive, não a
> máquina.

**Testar o backup sem precisar do desastre**, de vez em quando:

```powershell
node servidor\restaurar-servidor.js --arq "J:\Meu Drive\Backup Gerador-OS\<o mais recente>.bkp" --senha 'a-senha' --conferir
```

Abre o pacote e mostra o que tem dentro, **sem escrever nada**. É o único jeito
de saber que o backup presta antes de precisar dele.
