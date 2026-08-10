# O que falta depois do instalador — passo a passo

O servidor está no ar em **`https://193.168.0.8`**. O instalador fez tudo o que
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
> sozinho. Abrir `https://193.168.0.8` já basta. O texto do passo 3 continua
> abaixo como referência, para quando alguém precisar conferir ou desfazer.

---

## 1. Criar as contas

As contas não vieram no backup — elas vivem na autenticação, não nos dados.
**Sem conta, a pessoa não entra; e sem entrar, não vê cadastro, OS nem OE
nenhuma.** Este é o passo que, esquecido, faz parecer que "a intranet não
funciona" quando está tudo certo.

### Entrar no painel

No **servidor**, abra `http://localhost:8000`. Ele pede usuário e senha:

- **usuário:** `admin`
- **senha:** a linha `DASHBOARD_PASSWORD` do arquivo `C:\supabase\docker\.env`

O painel só responde aqui na máquina do servidor — de propósito.

### Criar cada pessoa

**Authentication** → **Users** → botão **Add user** → **Create new user**.

- **Email:** pode ser o e-mail real da pessoa ou um interno, tipo
  `maria@diverse.local`. Só precisa ser único e a pessoa precisar lembrar.
- **Password:** a senha dela.
- **Auto Confirm User:** deixe **ligado**.

> **Não use "Send invitation".** Convite manda e-mail, e não há servidor de
> e-mail configurado aqui — o convite nunca chega e a conta fica pendurada sem
> poder entrar. "Create new user" com Auto Confirm resolve na hora.

Repita para cada pessoa que vai usar o programa.

### Dar admin para você

Só a sua conta precisa disto. As outras não precisam de registro nenhum: sem
linha na tabela, o programa já trata como somente leitura.

**SQL Editor** → **New query**, troque o e-mail e rode:

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
`https://193.168.0.8` é de confiança. Sem ele, o navegador mostra "a conexão
não é particular" e o programa não funciona direito.

Ele nasceu junto com o servidor e mora em
`C:\Users\Pichau\Desktop\Gerador-OS\servidor\tls\ca.crt` — Área de Trabalho,
pasta `Gerador-OS`, pasta `servidor`, pasta `tls`.

### Não precisa de pendrive

O próprio servidor entrega o arquivo. **Em cada computador da fábrica:**

1. No navegador, digite **`https://193.168.0.8/ca.crt`**.
2. Vai aparecer "a conexão não é particular" — é esperado, é exatamente o que
   estamos consertando. Clique em **Avançado** → **Prosseguir para
   193.168.0.8 (não seguro)**.
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

Abra `https://193.168.0.8` nessa máquina. O cadeado tem de aparecer **sem
aviso**. Se aparecer aviso de "conexão não é particular", o certificado não
entrou — refaça com administrador de verdade e reabra o navegador.

> Vale para o servidor também, se você for usar o programa nele.

> O certificado do servidor vale até **novembro de 2028**. Quando vencer, rode
> `node servidor\gerar-certificado.js --ip 193.168.0.8` de novo — a autoridade
> continua a mesma, então **não** é preciso passar nas máquinas outra vez.

---

## 3. Apontar cada máquina para o servidor — automático

**Normalmente você não precisa fazer nada aqui.** Ao abrir
`https://193.168.0.8`, o programa pergunta ao próprio servidor quem ele é e se
conecta sozinho. A barra lateral já mostra **🏭 Servidor da fábrica**.

O passo abaixo continua valendo para conferir, corrigir ou desfazer — por
exemplo numa máquina que já ficou apontada para um endereço antigo.

Esta configuração vale **só na máquina onde foi feita** — é por isso que, quando
precisa ser feita à mão, tem de repetir em cada computador.

1. Abra **`https://193.168.0.8`** (não o endereço antigo da internet).
2. Entre com a conta daquela pessoa.
3. Vá em **Configurações** → cartão **Servidor da fábrica (rede local)**.
4. Preencha os dois campos:
   - **Endereço:** `https://193.168.0.8` — sem porta, a API vem pelo mesmo
     endereço.
   - **Chave pública (anon) do servidor local:** a chave ANON, que está em
     `servidor\tls\resumo-instalacao.txt` no servidor.
5. **Testar conexão**. Depois **Salvar e recarregar**.

A barra lateral deve passar a mostrar **🏭 Servidor da fábrica**.

**Se aparecer ☁ Nuvem em vermelho**, o servidor não respondeu, e o programa
abriu a cópia da nuvem em **modo consulta** — dá para ver e imprimir, mas não
editar. Isso é proposital: impede que os dois lados fiquem diferentes. Confira
o endereço, o certificado (tarefa 2) e se o servidor está ligado.

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
uns três minutos, de outra máquina, abra `https://193.168.0.8`.

Se o app abrir, está resolvido. Se não abrir, o logon automático não pegou —
resolva agora, não depois.

### No-break

O servidor guarda os dados da fábrica. Queda de energia com o Postgres
escrevendo é a receita clássica de banco corrompido. Um no-break de linha já
cobre o caso comum, que é a oscilação curta.

---

## Conferência final

| o quê | como | esperado |
|---|---|---|
| Contas | Cada pessoa entra com a conta dela | entra e vê os cadastros |
| Seu admin | Menu mostra **Nova OS** | aparece só para você |
| Certificado | `https://193.168.0.8` em outra máquina | cadeado sem aviso |
| Arquivos servidos | `https://193.168.0.8/vendor-check.html` | 7 linhas "ok" |
| Pastas automáticas | Configurações → as três pastas | deixa conectar |
| Sem internet | Tire o cabo do roteador (não do switch) | app continua normal |
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
- **Espelho para a nuvem** e **backup diário do servidor** — as duas receitas
  estão no fim do `README.md`. Até o backup existir, os dados moram num disco só.
