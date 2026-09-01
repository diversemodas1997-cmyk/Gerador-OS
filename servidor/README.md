# Servidor da fábrica — guia de instalação (Windows)

O que você vai montar: um PC na rede da fábrica rodando o **Supabase
auto-hospedado** (o mesmo software da nuvem, em Docker) e servindo os arquivos
do Gerador-OS. As máquinas passam a falar com ele; a nuvem vira espelho de
consulta.

> **Aviso honesto sobre este guia.** Escrevi os scripts e o schema e testei o que
> dava para testar aqui: o gerador de chaves foi validado contra uma
> implementação independente de JWT. O roteiro do Docker em si **não foi
> executado** — não tenho essa máquina. Vá conferindo passo a passo e me diga
> onde travar.

---

## Antes de começar: dois problemas conhecidos

Nenhum dos dois impede a instalação, mas os dois mudam o que você vai obter no
fim. Melhor saber agora.

### 1. Tudo passa por HTTPS — e isso não é opcional

O navegador só libera a escolha de pastas (gravação automática do PDF da OS, das
etiquetas, do backup e da OE) em **contexto seguro**: `https://` ou `localhost`.
Em `http://192.168.0.50` essas quatro funções simplesmente não existem.

Medido, não suposto — mesma máquina, mesmos arquivos, só mudando o endereço:

| endereço | contexto seguro | escolha de pastas |
|---|---|---|
| `http://192.168.0.8:8081` | não | indisponível |
| `https://192.168.0.8:8443` | **sim** | **disponível** |

Por isso o roteiro monta o servidor já com HTTPS (passos 5 e 6), usando um
certificado seu. Não envolve internet, domínio nem pagar nada.

### 2. As imagens dos desenhos vêm da sua pasta

As imagens não estão dentro dos dados: os dados guardam só o nome do arquivo, e
o app monta o endereço contra o servidor em uso. Os arquivos precisam existir na
fábrica, senão a folha de OS abre sem desenho.

Nenhum backup local tem essas imagens — todos guardam apenas o endereço. Mas a
pasta `Desenhos técnicos`, organizada por SKU, resolve: o passo 9 as importa
direto de lá, **sem depender da nuvem**. Foi o que tirou a restrição da conta do
caminho crítico da instalação.

---

## Caminho rápido: o instalador

Os passos 2 a 7 e 9 estão automatizados. No servidor, com o Docker Desktop já
instalado e aberto, dentro da pasta do Gerador-OS:

```powershell
.\servidor\instalar.ps1 -IP 192.168.0.50
```

Ele clona o Supabase, gera e grava as chaves, sobe os containers, cria as
tabelas, gera o certificado, migra os dados do backup mais recente, publica o
app, libera o firewall e importa as imagens dos desenhos. No fim imprime a chave
ANON e a lista do que sobrou para você.

**Pode rodar de novo.** Cada etapa confere se já foi feita e pula. As chaves não
são regeradas por acidente — trocar o `JWT_SECRET` invalidaria a chave já
configurada em todas as máquinas.

O que ele **não** faz, e por quê: instalar o Docker Desktop (exige tela,
administrador e reiniciar), criar as contas de login (cada pessoa tem a sua, e as
senhas são suas) e instalar o certificado nas máquinas (é um comando por máquina,
com administrador).

Os passos abaixo continuam valendo como referência — para entender o que o
instalador fez, ou para refazer alguma etapa isolada.

---

## Passo 1 — Docker Desktop

### Antes de baixar, confira três coisas no servidor

**1. A versão do Windows.** O Docker Desktop exige **Windows 10 Pro 22H2
(build 19045)** ou mais novo. Tecla Windows + R, digite `winver`, Enter.
Windows 10 **Home** não serve para o modo recomendado.

**2. A virtualização, ligada na BIOS.** É o tropeço mais comum. Abra o
Gerenciador de Tarefas (Ctrl+Shift+Esc) → aba **Desempenho** → **CPU** e procure
**Virtualização: Habilitada**.

Se estiver desabilitada, entra na BIOS (F2 ou Del ao ligar). O nome da opção
muda conforme o processador: em **AMD** chama-se **SVM Mode**; em **Intel**,
**Intel VT-x** ou **Intel Virtualization Technology**. Ligue, salve e saia.

**3. Memória.** O mínimo do Docker é 8 GB, mas aqui ele ainda vai carregar o
Supabase inteiro por cima. Com 8 GB funciona e fica apertado; **16 GB é o
confortável**.

### Licença

O Docker Desktop é gratuito para empresas com **menos de 250 funcionários e
menos de US$ 10 milhões de faturamento anual** — a Diverse/Dixie está
tranquilamente dentro disso. Não é preciso criar conta no Docker Hub: quando ele
pedir para entrar, pode pular.

### Instalar

1. Baixe em <https://www.docker.com/products/docker-desktop/> (botão
   *Download for Windows*).
2. Rode o `Docker Desktop Installer.exe`. Deixe marcado **Use WSL 2 instead of
   Hyper-V** — é a opção recomendada e a que o instalador do Gerador-OS espera.
3. Deixe terminar e **reinicie o computador** quando ele pedir.
4. Abra o Docker Desktop, aceite os termos e espere o canto inferior esquerdo
   mostrar **Engine running** (verde). A primeira vez demora alguns minutos.
5. Em **Settings → General**, marque **Start Docker Desktop when you sign in**.

### Conferir antes de seguir

No PowerShell:

```powershell
docker --version
docker run --rm hello-world
```

O segundo comando baixa uma imagem minúscula e imprime uma mensagem de boas
vindas. Se ele funcionar, o Docker está pronto de verdade — e não só instalado.

### Se der errado

| sintoma | causa provável |
|---|---|
| "WSL 2 installation is incomplete" | rode `wsl --update` no PowerShell e reabra |
| "Virtualization support not detected" | virtualização desligada na BIOS (ver acima) |
| Fica em "Docker Engine starting" para sempre | reinicie o computador; se persistir, Settings → Troubleshoot → Reset to factory defaults |
| `docker` não é reconhecido no PowerShell | feche e reabra o PowerShell depois de instalar |

> **Lembrete que vale mais que o resto.** O Docker Desktop só roda **dentro de
> uma sessão do Windows logada**. Sem o logon automático configurado, um
> reinício de madrugada deixa a fábrica sem sistema pela manhã. Isso é tratado
> na conferência final deste guia — não deixe para depois.

> **A fragilidade deste arranjo.** No Windows o Docker Desktop roda dentro da
> sessão do usuário: se a máquina reiniciar (atualização de madrugada, queda de
> energia) e ninguém fizer login, **o servidor não sobe** e a fábrica abre sem
> sistema. Trate isso como parte da instalação, não como detalhe: configure o
> logon automático do Windows numa conta local dedicada, e mantenha o no-break.
> É a falha mais provável de todo o conjunto.

## Passo 2 — Supabase local

Num terminal (PowerShell), numa pasta à sua escolha:

```powershell
git clone --depth 1 https://github.com/supabase/supabase
cd supabase\docker
copy .env.example .env
```

## Passo 3 — Chaves

Na pasta do Gerador-OS:

```powershell
node servidor\gerar-chaves.js
```

Ele imprime um bloco pronto. Abra `supabase\docker\.env` no Bloco de Notas e
**substitua** as linhas de mesmo nome pelas geradas. Guarde uma cópia do bloco
inteiro em lugar seguro — sem o `JWT_SECRET` não há como reemitir as chaves.

As chaves não são opcionais nem intercambiáveis: `ANON_KEY` e
`SERVICE_ROLE_KEY` são assinadas com o `JWT_SECRET`. Se as três não combinarem,
tudo responde 401 e a instalação parece quebrada sem dizer por quê.

Suba tudo:

```powershell
docker compose up -d
```

A primeira vez baixa vários gigabytes. Ao terminar, abra `http://localhost:8000`
— é o painel, com o usuário e a senha de `DASHBOARD_*`.

## Passo 4 — Criar as tabelas

No painel: **SQL Editor → New query**, cole o conteúdo de
`servidor/schema.sql` e rode. Confira com as consultas comentadas no fim do
arquivo.

## Passo 5 — Certificado HTTPS

Crie a autoridade certificadora da fábrica e o certificado do servidor:

```powershell
node servidor\gerar-certificado.js --ip 192.168.0.50
```

Ele confere o resultado antes de terminar: valida o certificado contra a própria
CA e verifica que o IP está na lista de nomes aceitos — um certificado sem isso
dá erro de segurança no navegador e parece problema de configuração.

Saem quatro arquivos em `servidor\tls`. O `ca.key` é **segredo**: quem o tiver
emite certificados em nome da sua CA. Ele fica no servidor e não vai para pasta
compartilhada nem para o repositório.

Depois, **em cada computador da fábrica**, com o `ca.crt` copiado para lá,
no PowerShell **como administrador**:

```powershell
Import-Certificate -FilePath .\ca.crt -CertStoreLocation Cert:\LocalMachine\Root
```

Feche e reabra o navegador. Chrome e Edge usam a lista do Windows; o Firefox tem
lista própria (Configurações → Certificados → Autoridades).

> O certificado do servidor vale até ~2 anos. Quando vencer, rode o gerador de
> novo — a CA continua a mesma, então **não** é preciso passar nas máquinas
> outra vez.

## Passo 6 — Servir o app e a API no mesmo endereço

O servidor precisa entregar o `index.html`, o `app.js` e a pasta `vendor/` —
senão as máquinas continuam dependendo do GitHub para abrir o programa e a
internet volta a ser obrigatória.

E precisa entregar a **API do Supabase no mesmo endereço**. Isto não é
preferência: uma página em `https://` não pode chamar uma API em `http://` — o
navegador bloqueia como conteúdo misto. Com os dois na mesma origem o problema
desaparece, e de quebra não há CORS para configurar.

Copie a pasta do Gerador-OS para o servidor (ex.: `C:\gerador-os`) e suba:

```powershell
docker compose -f servidor\docker-compose.app.yml up -d
```

O app fica em **`https://IP-DO-SERVIDOR`** (sem porta). Quem digitar `http://` é
redirecionado. Para atualizar o programa depois, basta um `git pull` na pasta.

## Passo 7 — Migrar os dados

Do backup local — **não precisa da nuvem**:

```powershell
node servidor\migrar-do-backup.js --url http://localhost:8000 --key <SERVICE_ROLE_KEY> --arq backups\BACKUP-COMPLETO-2026-08-07T19-57-40.json
```

O script recusa gravar por cima de um servidor que já tenha dados (use
`--sobrescrever` só se for mesmo a intenção). Confira no painel: **Table
Editor → shared_data**, uma linha `main`.

## Passo 8 — Contas de login

As contas vivem na autenticação, não nos dados: elas **não** vêm no backup.

**Crie uma conta para cada pessoa que vai usar o programa**, no painel
(**Authentication → Users → Add user**). Sem conta no servidor da fábrica, a
pessoa não entra — e sem entrar não vê cadastro, OS nem OE nenhuma. Este é o
passo que, esquecido, faz parecer que "a intranet não funciona" quando na
verdade está tudo certo e falta o login.

Depois, dê o papel de admin à sua conta. No **SQL Editor**:

```sql
INSERT INTO user_roles (user_id, role)
SELECT id, 'admin' FROM auth.users WHERE email = 'seu-email@exemplo.com'
ON CONFLICT (user_id) DO UPDATE SET role = 'admin';
```

As demais contas não precisam de linha nenhuma — sem registro, o app já as trata
como somente leitura.

### O que as máquinas de consulta enxergam

Tudo o que é de leitura, sem exceção: **todos os cadastros** (tecidos, cores,
modelos, grades, desenhos, etapas, equipe…), a **lista de OS e a folha de OS**
pronta para imprimir, a **Expedição e a folha do plano (OE)**, o estoque, o
ranking e as operações. No menu, só **Nova OS** é exclusiva do admin; os botões
de criar, editar e excluir de cada tela também somem.

Ou seja: consultar e imprimir funciona igual para todo mundo. O que muda é só
quem pode alterar.

## Passo 9 — Imagens dos desenhos técnicos

As imagens não estão nos dados: os dados guardam só o nome do arquivo, e o app
monta o endereço contra o servidor em uso. Os arquivos precisam existir na
fábrica, senão a folha de OS abre sem desenho.

**Caminho normal — a partir da sua pasta** (não precisa da nuvem):

```powershell
node servidor\importar-desenhos-da-pasta.js --pasta "Desenhos técnicos" --url http://localhost:8000 --key <SERVICE_ROLE_KEY> --so-listar
node servidor\importar-desenhos-da-pasta.js --pasta "Desenhos técnicos" --url http://localhost:8000 --key <SERVICE_ROLE_KEY>
```

O pareamento é pelo **SKU** do desenho, casado com a pasta e o nome do arquivo
(`CM.TRI.LISA-CAQUI` → `CM.TRI/TRICOLOR CAQUI.png`). Rode sempre o `--so-listar`
primeiro e confira a lista: é ali que se percebe um arquivo casado com o desenho
errado. Na dúvida ele não pareia — desenho trocado numa OS de corte custa tecido.

Desenho sem SKU preenchido vira pendência. O certo é preencher o SKU no cadastro
do desenho, no app; para resolver na hora, dá para usar um arquivo de mapa:

```json
{ "0013": "CM.REC/VERDE.png" }
```

e passar `--mapa mapa.json`.

**Caminho alternativo — copiar da nuvem** (exige o projeto da nuvem fora da
restrição):

```powershell
node servidor\copiar-desenhos.js --local http://localhost:8000 --key <SERVICE_ROLE_KEY>
```

## Passo 10 — Rede

1. **IP fixo** para o servidor (ou reserva por MAC no roteador). Se o IP mudar,
   todas as máquinas param de achá-lo.
2. **Firewall do Windows**: liberar entrada nas portas 80 e 443. A 8000 do
   Supabase **não** precisa ser aberta para a rede — as máquinas chegam nela
   pelo nginx, e mantê-la fechada evita um caminho sem HTTPS.

```powershell
New-NetFirewallRule -DisplayName "Gerador-OS (HTTP->HTTPS)" -Direction Inbound -LocalPort 80  -Protocol TCP -Action Allow
New-NetFirewallRule -DisplayName "Gerador-OS (HTTPS)"       -Direction Inbound -LocalPort 443 -Protocol TCP -Action Allow
```

## Passo 11 — Apontar as máquinas

Em cada computador, abra o app **pelo endereço do servidor**
(`https://IP-DO-SERVIDOR`), vá em **Configurações → Servidor da fábrica** e
preencha:

- **Endereço:** `https://IP-DO-SERVIDOR` (sem porta — a API vem pelo mesmo
  endereço, ver passo 6)
- **Chave:** a `ANON_KEY` gerada no passo 3

Clique em **Testar conexão** e depois em **Salvar e recarregar**. A barra
lateral deve passar a mostrar **🏭 Servidor da fábrica**.

Se aparecer **☁ Nuvem** em vermelho, o servidor não respondeu — o programa abre
a cópia da nuvem em modo consulta, de propósito, para os dois lados nunca
divergirem.

---

## Conferência final

| o quê | como |
|---|---|
| Arquivos sendo servidos | Abra `https://IP/vendor-check.html` — 7 linhas "ok" |
| Cadeado sem aviso | `https://IP` em outra máquina, depois de instalar o `ca.crt` |
| Pastas automáticas | Configurações → as três pastas devem deixar conectar |
| Banco e Realtime | As consultas comentadas no fim de `schema.sql` |
| Funciona sem internet | Tire o cabo do roteador (não do switch) e use o app normalmente |
| Sobe sozinho | Reinicie o servidor e veja se o app volta sem ninguém tocar nele |

O terceiro e o quarto são os que realmente importam — são o motivo de tudo isto.

---

# Espelho para a nuvem

Com a fábrica atendendo, a nuvem deixa de ser o dia a dia — mas continua valendo
por dois motivos: **consultar de fora** e ser uma **segunda cópia** dos dados,
fora do prédio. Para isso ela precisa receber o que acontece aqui.

O espelho é **sempre de mão única**: a fábrica manda, a nuvem recebe. Nada volta.
Isso é possível porque o app abre a nuvem em modo consulta — ninguém edita por
lá, então não existem dois lados para conciliar, que é onde mora o risco de
perder dados.

## As chaves

O script acha as duas sozinho, e por isso o comando agendado não carrega segredo
nenhum:

| chave | onde ele procura |
|---|---|
| servidor da fábrica | `SERVICE_ROLE_KEY` em `C:\supabase\docker\.env` (de onde os outros scripts já leem) |
| nuvem | `C:\supabase\nuvem-service-role.key` — uma linha só, a chave e nada mais |

O arquivo da nuvem já existe na máquina, com a explicação dentro; falta colar a
chave, que está no painel do Supabase em **Project Settings → API → service_role**.
Enquanto ela não estiver lá, o script para na hora e diz exatamente isso.

> Por que num arquivo, e não no comando: essa chave passa por cima de todas as
> permissões. Na linha de comando ela aparece no Agendador de Tarefas, no
> histórico do PowerShell e em qualquer print da tela. Num arquivo do servidor,
> não. Nunca no navegador, nunca dentro do repositório.

## Agendar

```powershell
$node = (Get-Command node).Source
$acao = New-ScheduledTaskAction -Execute $node -Argument 'servidor\espelhar-para-nuvem.js' `
  -WorkingDirectory 'C:\Users\Pichau\Desktop\Gerador-OS'
$quando = New-ScheduledTaskTrigger -Once -At 7am -RepetitionInterval (New-TimeSpan -Minutes 30) `
  -RepetitionDuration (New-TimeSpan -Days 3650)
Register-ScheduledTask -TaskName "Espelho Gerador-OS" -Action $acao -Trigger $quando -Force
```

**Já está agendada nesta máquina** (26/08/2026), a cada 30 minutos a partir das
7h. Se nada mudou desde a última vez, ele não escreve nada — dá para rodar com
frequência sem custo. Para rodar na hora: `Start-ScheduledTask "Espelho Gerador-OS"`,
ou `node servidor\espelhar-para-nuvem.js` dentro da pasta do programa.

## O que ele faz e o que se recusa a fazer

- Envia o blob e o mapa de versões, com o mesmo carimbo da fábrica — assim quem
  abrir pela nuvem também baixa só a chave que mudou.
- Envia as imagens de desenho que ainda não estiverem lá. **Nunca apaga** da
  nuvem: imagem sobrando não custa quase nada, imagem faltando quebra uma folha
  de OS.
- Envia as **mensagens** do canal de recados (tabela `mensagens`) e, ao
  contrário das imagens, **apaga da nuvem o que foi apagado na fábrica**: quem
  apagou o próprio recado fez isso de propósito, e deixá-lo legível de fora
  desfaria a decisão pelas costas. Este passo roda mesmo quando os dados não
  mudaram — recado novo não altera o carimbo do `shared_data`. Desliga com
  `--sem-mensagens`. Se um dos lados ainda não tiver a tabela
  (`sql/supabase-mensagens.sql`), ele avisa e segue: o espelho dos dados é o que
  não pode parar.
- **Recusa espelhar quando a fábrica está sem OS e sem desenhos e a nuvem tem
  dados.** É a mesma trava que já protege o app. Se a fábrica ficar vazia por um
  problema — restauração pela metade, migração que não rodou —, a nuvem é a
  única cópia boa que resta, e o espelho não pode ser justamente o que a destrói.
  Nesse caso ele sai com erro e não envia nada.
- Internet fora não é motivo de alarme: ele falha e a execução seguinte recupera
  o atraso sozinha.
- Envia as **contas de acesso** (`servidor/espelhar-contas.js`, chamado no fim de
  cada passada). Ver abaixo.

## As contas de acesso vão junto

A cópia da nuvem existe para o dia em que o servidor da fábrica cair. De nada
adianta ela ter todos os dados se ninguém consegue **entrar** nela — e era esse
o caso: as contas de nome (`admin`, `nathaly`, `enfesto.corte`…) nascem só na
fábrica, pela função `usuarios`, e a nuvem nunca soube delas. Em **01/09/2026** o
Admin digitava a senha certa e ouvia *"Nome ou senha incorretos"*: a senha estava
certa, o servidor é que era o outro.

O que o espelho de contas faz:

- Copia o **hash** da senha, nunca a senha. As senhas continuam sendo as mesmas
  nos dois lugares, e nenhuma senha em claro passa pelo script, pela rede ou pela
  tela.
- Mantém o **mesmo id** dos dois lados, e repõe o papel (`admin`) na nuvem.
- Só mexe no que mudou: guarda a impressão digital de cada hash em
  `C:\supabase\espelho-contas.json` e compara. Rodar de meia em meia hora não
  custa nada.
- Quando a senha muda na fábrica, a conta da nuvem é **apagada e recriada** com o
  hash novo. O caminho óbvio — `PUT /admin/users/<id>` com `password_hash` —
  responde 200 e **não troca a senha**: o campo só é lido na criação (testado em
  01/09/2026). Um espelho que diz ter copiado e não copiou é pior do que não
  existir.
- Só cuida do que termina em `@diverse.local`. As contas de e-mail de verdade
  (@gmail) não são tocadas.

Para rodar sozinho: `node servidor\espelhar-contas.js`, ou
`node servidor\espelhar-contas.js --so-listar` para ver o que ele faria.

## Custo

Enviar é ingresso de dados, que não entra na cota que restringiu o projeto. O
espelho lê da nuvem apenas um carimbo por execução. Ele não recria o problema de
tráfego.

---

# Recuperação: trocar o servidor por outro computador

O disco desse PC passa a ser onde seus dados moram. O pacote de recuperação
existe para que, se ele morrer, outro computador ocupe o lugar dele no mesmo
dia.

## O que vai no pacote

Um arquivo só, **cifrado**, com o banco inteiro (`pg_dump`: dados, contas de
login com as senhas, papéis, metadados do Storage), as imagens dos desenhos e o
**`.env` com as chaves originais**.

Esse último item é o que muda o tamanho do estrago. Restaurando as mesmas
chaves, o servidor novo nasce com a mesma `ANON_KEY` — e **as máquinas da
fábrica voltam a funcionar sem serem tocadas**. Sem isso, seria preciso passar
em cada computador reconfigurando a chave.

Vai cifrado porque carrega segredos e hashes de senha, e o destino é uma pasta
sincronizada. O formato detecta adulteração: um arquivo truncado pela
sincronização falha ao abrir, em vez de restaurar um servidor pela metade.

> **A senha do backup não fica em lugar nenhum do sistema.** Anote-a fora do
> computador — junto dos segredos do passo 3. Sem ela o pacote não abre, e não
> há como recuperá-la.

> **Pacotes de 10/08 a 14/08/2026 (até as 17h) pedem a senha entre aspas
> simples.** O `-Agendar` montava o argumento como `-Senha 'a-senha'`, e o
> `powershell.exe -File` tira as aspas duplas mas **deixa as simples dentro do
> valor** — a senha gravada nesses pacotes é `'a-senha'`, com as aspas. Corrigido
> em 14/08/2026; os pacotes gerados a partir dali abrem com a senha anotada, sem
> aspas. Para os antigos, digite a senha entre aspas simples. Achado conferindo um
> pacote recém-gerado: **o backup que nunca foi aberto não é backup.**

## Onde guardar: Google Drive, não a nuvem do Supabase

Guardar a recuperação de desastre dentro do serviço que já deixou a fábrica sem
sistema é frágil. Vocês já usam o `J:\Meu Drive` — uma pasta ali já é fora do
prédio, versionada e sem custo de tráfego.

## Backup diário automático

Instale o Google Drive no servidor e agende:

```powershell
$acao = New-ScheduledTaskAction -Execute "node" `
  -Argument 'C:\gerador-os\servidor\backup-servidor.js --docker C:\supabase\docker --destino "J:\Meu Drive\Backup Gerador-OS" --senha SUA-SENHA-AQUI' `
  -WorkingDirectory 'C:\gerador-os'
$quando = New-ScheduledTaskTrigger -Daily -At 12:30
Register-ScheduledTask -TaskName "Backup Gerador-OS" -Action $acao -Trigger $quando -RunLevel Highest
```

Horário de almoço, e não de madrugada, de propósito: o `pg_dump` roda com o
Docker de pé, e de madrugada a máquina pode ter reiniciado sem ninguém logar.

Cada execução **reabre e confere** o arquivo que acabou de gravar. Um backup
nunca aberto não é backup, é uma suposição — se a conferência falhar, o arquivo
é apagado em vez de ficar passando por bom. Mantém 14 dias por padrão.

## Testar o backup sem precisar do desastre

De vez em quando, em qualquer computador:

```powershell
node servidor\restaurar-servidor.js --arq "J:\Meu Drive\Backup Gerador-OS\servidor-gerador-os-2026-08-10.bkp" --senha SUA-SENHA --conferir
```

Ele abre o pacote e mostra o que tem dentro, **sem escrever nada**. É o único
jeito de saber que o backup presta antes de precisar dele.

## Trocar o servidor

Na máquina nova: passos 1 e 2 deste guia (Docker + clonar o Supabase). Não
precisa gerar chaves nem rodar o `schema.sql` — tudo isso vem no pacote.

```powershell
node servidor\restaurar-servidor.js --arq <o pacote mais recente> --docker C:\supabase\docker --senha SUA-SENHA
```

Ele restaura o `.env` e as imagens, grava o banco num `.sql` e imprime os três
comandos que faltam (subir os containers, aplicar o banco, reiniciar).

**Dê à máquina nova o mesmo IP fixo da antiga.** Aí a recuperação termina sem
tocar em nenhum computador da fábrica: eles procuram o mesmo endereço, com a
mesma chave, e voltam sozinhos.

## O que ainda falta depois

1. **O espelho local → nuvem**, para o backup fora do prédio e a consulta
   remota continuarem valendo.
3. **HTTPS**, se as pastas automáticas forem necessárias fora do servidor.
2. **Backup do servidor.** O disco desse PC passa a ser onde seus dados moram.
   Os snapshots diários do app continuam funcionando, mas ficam no mesmo disco —
   até o espelho existir, mantenha a pasta de backup automático ligada numa
   máquina que não seja o servidor.
