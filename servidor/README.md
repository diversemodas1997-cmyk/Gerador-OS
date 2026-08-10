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

### 1. As pastas automáticas param de funcionar em `http://`

O navegador só libera a escolha de pastas (`showDirectoryPicker`, usada em 14
pontos do app) em **contexto seguro**: `https://` ou `localhost`. Um endereço
como `http://192.168.0.50:8000` **não** é contexto seguro.

Na prática, nas máquinas que acessarem por IP deixam de funcionar:

- pasta de backup automático do JSON;
- gravação automática do PDF da OS e das etiquetas;
- pasta das Ordens de Expedição.

O resto do programa funciona normalmente. As saídas:

- **Servir por HTTPS** com um certificado próprio, instalando a autoridade dele
  nas ~4 máquinas. É a solução certa, e é trabalho à parte.
- **Deixar essas funções só na máquina do servidor**, que acessa por
  `http://localhost:8000` e portanto é contexto seguro.

Se as pastas automáticas importam para você (a de PDF é usada todo dia), me
avise que eu incluo o HTTPS no roteiro.

### 2. Os desenhos técnicos precisam ser copiados

As 25 imagens não estão dentro dos dados. Os dados guardam só o **nome do
arquivo**, e o app monta o endereço contra o servidor em uso — o mesmo desenho
é buscado na fábrica quando ela está de pé, e na nuvem quando não. Mas os
arquivos precisam existir dos dois lados.

O script `copiar-desenhos.js` copia da nuvem para a fábrica, e é **o único passo
que precisa da nuvem acessível**. Como o projeto está restrito agora, fica para
depois que o serviço voltar. Faça-o **antes** de virar a chave, senão a folha de
OS abre sem desenho — justamente o que se quer evitar.

---

## Passo 1 — Docker Desktop

1. Instale o **Docker Desktop para Windows** (ele instala o WSL2 se faltar).
2. Reinicie e abra o Docker Desktop uma vez, até aparecer "Engine running".
3. Em **Settings → General**, marque **Start Docker Desktop when you sign in**.

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

## Passo 5 — Servir os arquivos do app

O servidor precisa entregar também o `index.html`, o `app.js` e a pasta
`vendor/` — senão as máquinas continuam dependendo do GitHub para abrir o
programa, e a internet volta a ser obrigatória.

Copie a pasta do Gerador-OS para o servidor (ex.: `C:\gerador-os`) e suba um
servidor de arquivos, em um `docker-compose.app.yml` separado — separado de
propósito, para não ser sobrescrito quando você atualizar o Supabase:

```yaml
services:
  app:
    image: nginx:alpine
    container_name: gerador-os-web
    restart: always
    ports:
      - "8080:80"
    volumes:
      - C:\gerador-os:/usr/share/nginx/html:ro
```

```powershell
docker compose -f docker-compose.app.yml up -d
```

O app fica em `http://IP-DO-SERVIDOR:8080`. Para atualizar o programa depois,
basta um `git pull` nessa pasta.

## Passo 6 — Migrar os dados

Do backup local — **não precisa da nuvem**:

```powershell
node servidor\migrar-do-backup.js --url http://localhost:8000 --key <SERVICE_ROLE_KEY> --arq backups\BACKUP-COMPLETO-2026-08-07T19-57-40.json
```

O script recusa gravar por cima de um servidor que já tenha dados (use
`--sobrescrever` só se for mesmo a intenção). Confira no painel: **Table
Editor → shared_data**, uma linha `main`.

## Passo 7 — Contas de login

As contas vivem na autenticação, não nos dados: elas **não** vêm no backup.
Recrie-as no painel (**Authentication → Users → Add user**) — são poucas.

Depois, dê o papel de admin à sua conta. No **SQL Editor**:

```sql
INSERT INTO user_roles (user_id, role)
SELECT id, 'admin' FROM auth.users WHERE email = 'seu-email@exemplo.com'
ON CONFLICT (user_id) DO UPDATE SET role = 'admin';
```

As demais contas não precisam de linha nenhuma — sem registro, o app já as trata
como somente leitura.

## Passo 8 — Copiar os desenhos (exige a nuvem no ar)

```powershell
node servidor\copiar-desenhos.js --local http://localhost:8000 --key <SERVICE_ROLE_KEY> --so-listar
node servidor\copiar-desenhos.js --local http://localhost:8000 --key <SERVICE_ROLE_KEY>
```

Ele só copia arquivos; nada nos dados é alterado, porque o nome é o mesmo dos
dois lados. Por isso é seguro rodar de novo quantas vezes precisar — se alguma
falhar, repita.

## Passo 9 — Rede

1. **IP fixo** para o servidor (ou reserva por MAC no roteador). Se o IP mudar,
   todas as máquinas param de achá-lo.
2. **Firewall do Windows**: liberar entrada nas portas 8000 e 8080.

```powershell
New-NetFirewallRule -DisplayName "Supabase local"  -Direction Inbound -LocalPort 8000 -Protocol TCP -Action Allow
New-NetFirewallRule -DisplayName "Gerador-OS web"  -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow
```

## Passo 10 — Apontar as máquinas

Em cada computador, abra o app **pelo endereço do servidor**
(`http://IP:8080`), vá em **Configurações → Servidor da fábrica** e preencha:

- **Endereço:** `http://IP-DO-SERVIDOR:8000`
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
| Arquivos sendo servidos | Abra `http://IP:8080/vendor-check.html` — 7 linhas "ok" |
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

## Agendar

```powershell
$acao = New-ScheduledTaskAction -Execute "node" `
  -Argument 'C:\gerador-os\servidor\espelhar-para-nuvem.js --local http://localhost:8000 --local-key <SR-LOCAL> --nuvem https://ckkqrjkhorvaahyazqsr.supabase.co --nuvem-key <SR-NUVEM>' `
  -WorkingDirectory 'C:\gerador-os'
$quando = New-ScheduledTaskTrigger -Once -At 7am -RepetitionInterval (New-TimeSpan -Minutes 30)
Register-ScheduledTask -TaskName "Espelho Gerador-OS" -Action $acao -Trigger $quando -RunLevel Highest
```

A cada 30 minutos. Se nada mudou desde a última vez, ele não escreve nada — dá
para rodar com frequência sem custo.

> A `SERVICE_ROLE_KEY` da nuvem passa por cima de todas as permissões. Ela fica
> só nesse comando, na máquina do servidor. Nunca no navegador, nunca dentro do
> repositório.

## O que ele faz e o que se recusa a fazer

- Envia o blob e o mapa de versões, com o mesmo carimbo da fábrica — assim quem
  abrir pela nuvem também baixa só a chave que mudou.
- Envia as imagens de desenho que ainda não estiverem lá. **Nunca apaga** da
  nuvem: imagem sobrando não custa quase nada, imagem faltando quebra uma folha
  de OS.
- **Recusa espelhar quando a fábrica está sem OS e sem desenhos e a nuvem tem
  dados.** É a mesma trava que já protege o app. Se a fábrica ficar vazia por um
  problema — restauração pela metade, migração que não rodou —, a nuvem é a
  única cópia boa que resta, e o espelho não pode ser justamente o que a destrói.
  Nesse caso ele sai com erro e não envia nada.
- Internet fora não é motivo de alarme: ele falha e a execução seguinte recupera
  o atraso sozinha.

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
