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

### 2. Os desenhos técnicos moram na nuvem

As 25 imagens não estão dentro dos dados: os dados guardam o **endereço** delas
no Storage da nuvem. Sem copiá-las, a folha de OS abre sem desenho quando a
internet cair — o oposto do que se quer.

O script `copiar-desenhos.js` resolve, mas ele é **o único passo que precisa da
nuvem acessível**. Como o projeto está restrito agora, esse passo fica para
depois que o serviço voltar. Faça-o **antes** de virar a chave.

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

Ele só reescreve os endereços se **todas** as imagens vierem — reescrever pela
metade deixaria desenhos quebrados.

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

## O que ainda falta depois

1. **Endereço dos desenhos.** Depois do passo 8 eles apontam para a rede local,
   então **de fora da fábrica os desenhos não aparecem**. A solução é guardar só
   o nome do arquivo e montar o endereço conforme o servidor em uso. É mudança
   no app, ainda por fazer.
2. **O espelho local → nuvem**, para o backup fora do prédio e a consulta
   remota continuarem valendo.
3. **HTTPS**, se as pastas automáticas forem necessárias fora do servidor.
4. **Backup do servidor.** O disco desse PC passa a ser onde seus dados moram.
   Os snapshots diários do app continuam funcionando, mas ficam no mesmo disco —
   até o espelho existir, mantenha a pasta de backup automático ligada numa
   máquina que não seja o servidor.
