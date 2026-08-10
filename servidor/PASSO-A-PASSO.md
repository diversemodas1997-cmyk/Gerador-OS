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

### Não precisa de pendrive

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

## 3. Apontar cada máquina para o servidor — automático

**Normalmente você não precisa fazer nada aqui.** Ao abrir
`https://193.168.0.200`, o programa pergunta ao próprio servidor quem ele é e se
conecta sozinho. A barra lateral já mostra **🏭 Servidor da fábrica**.

O passo abaixo continua valendo para conferir, corrigir ou desfazer — por
exemplo numa máquina que já ficou apontada para um endereço antigo.

Esta configuração vale **só na máquina onde foi feita** — é por isso que, quando
precisa ser feita à mão, tem de repetir em cada computador.

1. Abra **`https://193.168.0.200`** (não o endereço antigo da internet).
2. Entre com a conta daquela pessoa.
3. Vá em **Configurações** → cartão **Servidor da fábrica (rede local)**.
4. Preencha os dois campos:
   - **Endereço:** `https://193.168.0.200` — sem porta, a API vem pelo mesmo
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
