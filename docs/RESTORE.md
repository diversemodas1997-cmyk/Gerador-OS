# Backup e Restauração — Gerador-OS

Guia para restaurar o sistema **sem perder dados**. O sistema tem 3 componentes
independentes: **CÓDIGO**, **DADOS** e **SUPABASE (infra)**. Cada um tem seu
backup próprio.

---

## Ponto de restauração deste backup

- **CÓDIGO:** tag `restore-2026-08-05-e` (cache-buster `app.js ?v=2026-08-05n`,
  `styles.css ?v=2026-08-05n`). O que mudou desde o `restore-2026-08-05-d` — 12
  commits, em três frentes: o **cadastro de operações por função**, o **tempo do
  enfesto** e a **importação de riscos**.

  **CADASTRO DE OPERAÇÕES POR FUNÇÃO**
  - **o PASSO da corrente passa a ser gravado, não adivinhado.** A que passo uma
    operação pertencia era decidido por regex sobre o NOME, a cada consulta:
    renomear "Corte de enfesto" para "Corte na esteira" quebrava o vínculo em
    silêncio, nenhuma linha respondia mais pelo passo 5, e a cascata caía no
    histórico do plano — voltando a criar a operação com o **nome e o tempo
    antigos**. Agora a linha grava `passoId`, escolhido num seletor
    pré-preenchido pelo mesmo regex de sempre;
  - **tempo por TIPO DE ENFESTO** (o SKU no meio do nome da grade: BM.TRI,
    CM.LISA, CM.REC…). Era um número só por posto — os mesmos 80 min de enfesto
    para a camiseta lisa de uma fase e para a blusa tricolor de cinco. O seletor
    se enche sozinho com os tipos que existem nas grades;
  - **"todos os tipos" saiu das opções**: cada linha é de um produto. Linha sem
    tipo é cadastro por terminar — ainda vale para todos, mas aparece com badge
    vermelho na tabela de Funções e num aviso ao salvar. Quando só existem linhas
    de OUTROS tipos, a função vale mas o **tempo não vem** (sai zero, e o aviso
    da alocação nomeia a operação): emprestar o número do vizinho seria errar
    parecendo certo;
  - **dois quadros na janela** — "Fila do posto" e "Hora marcada" —, cada um com
    título, contagem e o motivo de aquelas estarem juntas. Preencher "todo dia
    às" move a linha de um quadro para o outro na hora;
  - **"Mover enfesto" antes do "Enfesto"**: passo sem linha no cadastro daquele
    posto recebia índice 1e9 e afundava para o fim do grupo. Agora não reordena —
    quem não disse nada não inverte ordem física.

  **TEMPO DO ENFESTO**
  - **a medição passa a ser tempo de TRABALHO, não de relógio.** Era
    `fim − início` cru, com o café e o almoço dentro: das 29 medições, **15
    atravessam uma pausa**, e a pior (OS 0405 Barra/Punhos) gravava 165 min para
    75 min de trabalho. O número inflado alimentava toda média do programa — a
    taxa por camada/metro, a média da grade, a conferência da folha. No total:
    **2890 min → 2215 min**. O plano já fazia isso do outro lado: ele PARTE a
    operação longa em volta das pausas e reserva só o trabalho;
  - **o tempo cadastrado deixa de ser sobrescrito pela medição** no modal de
    operação. Ele era escrito na duração e apagado três linhas abaixo pelo
    "medido" — que, quando o nome não casa com nenhuma fase, é a SOMA de todas as
    fases da grade. Era o caminho de um enfesto de gola de 15 min nascer com
    horas.

  **IMPORTAÇÃO DE RISCOS (PDF)** — três defeitos que faziam a correção não colar:
  - **a gravação ia para uma grade ÓRFÃ e o toast dizia que deu certo.** A tela
    guardava a grade escolhida como referência ao objeto que estava em
    `STATE.grades` quando o PDF foi lido — mas o realtime/polling chama
    `loadState`, que substitui o array por objetos novos. A conta rodava, o save
    gravava o array vivo (sem a alteração) e o aviso anunciava "1 fase
    atualizada". Agora reacha pelo ID antes de desenhar e antes de gravar, e
    **confere no cadastro vivo** antes de dizer que deu certo;
  - **PDF sem medida deixa de morrer em silêncio.** Trocar a grade remarcava a
    linha sem conferir se havia comprimento/largura, e o clique estourava num
    `.toFixed` de `null` — nada gravado, nenhum aviso. Numa grade como a
    `P ao G3 | BM.LISA | 177cm`, onde **doze cadastros têm os mesmos tamanhos**,
    escolher a grade na lista é obrigatório, então o caminho era certeiro;
  - **a largura barra o palpite, e dá para CRIAR a fase que falta.** Um risco de
    RIBANA (1,68 × 0,542) numa grade que só tem Corpo (10,29 × 1,750) saía
    marcado "pela medida" como se fosse o corpo — 8,46 m de diferença contavam
    como acerto, porque o comprimento é justamente o número que veio ser
    corrigido. Agora, sem nenhuma fase com a largura do risco, a medida não
    decide, e a linha chega propondo criar a fase com o nome tirado do arquivo.

  **EXPEDIÇÃO** — ordem **crescente/decrescente** no planejamento semanal e
  mensal (a folha do plano continua em ordem de calendário).

  Cópia do código deste ponto em
  `backups-codigo/*.20260805e-tipo-enfesto.CÓPIA`.

- **DADOS:** exportação de **05/08/2026 15:05**, em
  `backups/BACKUP-COMPLETO-2026-08-05T15-05-08.json` (2,46 MB, 29 chaves, 1701
  registros): **182 OS**, 115 grades, 25 desenhos, **1157 operações**, 46 cargas
  de expedição, 49 mov. de estoque, 11 tecidos, 37 cores, 11 funções, 10 pessoas
  na equipe, 6 modelos, 15 etapas, 17 componentes, 4 materiais.
  A mesma cópia está em `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`.

  > Restaurar: app → Configurações → **Importar JSON** → escolher este arquivo.
  > Sobrescreve tudo, então é o caminho de "perdi geral". Para casos parciais,
  > ver a seção 2 abaixo.

  Contra o de 04/08 17:23:
  - **operações: 173 → 1157 (+984).** Não é planejamento novo de OS: é a rotina
    de **hora marcada** preenchendo o calendário. O plano agora cobre **48 dias,
    de 24/07 a 02/10** — o horizonte de 60 dias que
    `_opGarantirHorariosFixosNoPeriodo` alcança —, com **24 operações por dia
    útil** (café, almoço, preparação das máquinas, limpeza, cobrir máquinas, nas
    11 funções). Restaurar este ponto traz esse calendário junto.
  - **OS: 177 → 182** (+5); cargas de expedição 43 → 46; mov. de estoque 41 → 49.
  - **grades: 115, sem mudança.** Das 115, **104 têm a largura no nome**; as 11
    que faltam são as `CO.*` (Jaguar, Prime, Rugão), `PM.LISA`, `PM.TRI`,
    `SM.LISO` e `SM. ESPARTANA`, que não passaram pela importação.

  **A conferir neste retrato:** **73 das 115 grades têm ao menos uma fase sem
  comprimento**. É o que a importação de riscos existe para preencher — e é o
  que os três defeitos acima impediam de colar. Vale reimportar os riscos dessas
  grades com este ponto de código no ar.

### Ponto anterior a este

- **CÓDIGO:** tag `restore-2026-08-05-d` (cache-buster `app.js ?v=2026-08-05d`,
  `styles.css ?v=2026-08-04a`). O que mudou desde o `restore-2026-08-03-l` — 33
  commits, em três frentes: a **importação de riscos**, o **planejamento de
  operações** e a **gravação na nuvem**.

  **IMPORTAÇÃO DE RISCOS (PDF)** — a maior parte do trabalho. A janela pedia
  sete decisões de uma vez, para riscos que nem eram da mesma grade:
  - **um PDF por passo**, com uma pergunta só ("cria uma grade ou corrige uma
    existente?"). O agrupamento era só por distribuição de tamanhos, e juntava
    numa "grade" os quatro BM.LISA (174, 177, 180, 182 cm), os dois CM.LISA e o
    PM.LISA, todos "P ao G3";
  - **a fila anda por PASTA**: uma pasta de largura é uma grade, e o cadastro
    fecha inteiro — todas as fases — antes de o próximo começar. Entre pastas,
    primeiro as que têm grade a cadastrar, depois as que só corrigem;
  - **largura diferente é outra grade**, e ela entra no nome
    (`M-G-GG-G1 | BM.LISA 174cm`). A sugestão não atravessa larguras, e a escolha
    manual pergunta antes de trocar a medida de uma grade boa;
  - **grade que já existe só pode ser corrigida** — a opção de criar some, para
    não nascer duplicata;
  - **ribana, gola, viés e forro são FASES, não grades**: entram na grade do
    corpo, achada pela pasta (com a largura). O risco delas tem os tamanhos
    delas ("10xM-10xG…") e não casava com grade nenhuma;
  - **lê do nome do arquivo** o que ele já diz: SKU, tipo de peça, fase
    ("CORPO 2" → "Corpo Parte 2") e variação (LISA = básica, TRI = tricolor,
    REC = Recortada);
  - **preenche pelo SKU**: unidades da grade (BM e PM = 1, resto = 2) e peças
    por pacote (BM = 36, CM = 80);
  - **nome da grade editável**, e a sugestão nunca é a de uma grade existente;
  - o cabeçalho mostra a **soma do excedente por extenso** (2,79 + 15 cm =
    2,94), que antes aparecia só como total e passava por erro de conta.

  **PLANEJAMENTO DE OPERAÇÕES**
  - **enfesto longo repartido pelos vãos** ("Enfesto 1/3", "2/3", "3/3"). O maior
    vão do dia é 2h15 entre as pausas, e nenhum enfesto acima disso conseguia ser
    planejado — sumia a fase mais longa, e a corrente começava pelo Corpo 2;
  - **a corrente que não cabe hoje continua no próximo dia útil**, em vez de o
    resto da fase ser descartado;
  - **tempo de corte por fase**, com a OS de referência enquanto não há média (na
    OS 0405 o corte levou 60, 20 e 20 min; o cadastro dava 15 para as três);
  - **ordem das operações por função**, com setas ↑↓; hora marcada fica à parte,
    ordenada pelo relógio;
  - **botão "+ Expedição do dia"**, que monta a carga a partir das janelas de OE;
  - **hora marcada** repetida na mesma pessoa deixa de ser sobreposição, e o
    cadastro passa a alcançar a operação já planejada (inclusive em dia passado);
  - **gola e viés fora do plano automático**; **busca de OS por número** na
    janela de alocação; **Retirar OS** lista só o que o alcance retira.

  **GRAVAÇÃO NA NUVEM** — o clique feito DURANTE uma gravação se perdia e era
  revertido (`_dirtyKeys.clear()` limpava também o que sujou no meio do voo). É
  o que fazia o checklist e o número do tom voltarem atrás sozinhos.

  **ARQUIVO** — raiz do projeto organizada (`sql/`, `docs/`, `dados/`), desenhos
  técnicos com nome padronizado (`<LINHA> - <PEÇA> <GRADE>.pdf`) e sem cópia
  repetida — **menos as ribanas por largura, que são uma por grade e devem
  existir**.

  Cópia do código deste ponto em
  `backups-codigo/*.20260805d-importacao-riscos.CÓPIA`.

- **DADOS:** exportação de **04/08/2026 17:23**, em
  `backups/BACKUP-COMPLETO-2026-08-04T20-23-49.json` (1,77 MB, 29 chaves, 701
  registros): **177 OS**, **115 grades**, 25 desenhos, **173 operações**, 43
  cargas de expedição, 41 mov. de estoque, 11 tecidos, 37 cores, 11 funções, 10
  pessoas na equipe, 6 modelos, 15 etapas, 17 componentes, 4 materiais.
  A mesma cópia está em `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`.

  > Restaurar: app → Configurações → **Importar JSON** → escolher este arquivo.
  > Sobrescreve tudo, então é o caminho de "perdi geral". Para casos parciais,
  > ver a seção 2 abaixo.

  **É o retrato da importação de riscos rodando.** Contra o de 03/08:
  - **grades: 66 → 115.** 101 nomes novos, 52 que sumiram, 13 mantidos — o
    cadastro foi refeito com a **largura no nome** (`P ao G3 | CM.TRI | 116.5cm`,
    `2X P ao G3 | BM.TRI | 177cm`). Das 115, **104 já têm a largura**; as 11 que
    faltam são as `CO.*` (Jaguar, Prime, Rugão), que não passaram pela
    importação.
  - **operações: 288 → 173.** Os dias 22/07 e 27/07 foram esvaziados e 24/07,
    28/07, 03/08 e 04/08 encolheram — é a retirada e realocação das OS. **05/08
    (+19) e 06/08 (+23) apareceram**: é a corrente transbordando para o próximo
    dia útil, que este ponto de código passou a fazer em vez de descartar o resto
    da fase.
  - cargas de expedição: 40 → 43.

  **Conferido neste retrato:** **nenhuma grade tem fase repetida** — o defeito
  corrigido em `9db3a4a` (acrescentar uma fase nova em vez de corrigir a
  existente, quando a grade nascia durante a própria importação) não deixou
  rastro nas 115.

### Ponto anterior a este

- **CÓDIGO:** tag `restore-2026-08-03-l` (cache-buster `app.js ?v=2026-08-03l`,
  `styles.css ?v=2026-08-03d`). O que mudou desde o `restore-2026-07-30-k` — 13
  commits, quase todos no **planejamento de operações** e no **tempo de
  enfesto**:
  - **Horário fixo:** mudar o "todo dia às" no cadastro da função passa a valer
    nos dias JÁ planejados (antes só nascia com a operação); o passo da corrente
    de uma OS deixou de levar o 📌, que o travava no lugar.
  - **Agenda pela PESSOA, não só pelo posto:** a cascata perguntava só até quando
    o POSTO estava ocupado, e uma pessoa cobre vários postos aqui — o enfesto da
    fase 2 nascia às 08:20 com a mesma pessoa cortando até 08:35. A exceção da
    esteira também apagava a operação da checagem de conflito da PESSOA.
  - **Tempo de enfesto** ganhou uma cadeia de fontes: (1) média medida daquela
    fase NAQUELA grade — manda para mais e para menos; (2) tempo cadastrado **na
    fase da grade** (campo novo); (3) tempo cadastrado **na função**, agora
    editável e com o **comprimento de referência** ("1h20 para grade de 8 m"),
    valendo como piso só em grade de porte parecido; (4) média da grade, em grade
    curta; (5) estimativa por comprimento. O campo do Enfesto era travado na tela
    e zerado ao salvar.
  - **"Mover enfesto" inflado:** recebia a soma do tempo de estender TODAS as
    fases (7h33 na BM.TRICOLOR) porque a regra olhava o nome da FUNÇÃO
    ("Operador de enfestadeira" casa com /enfest/). Corrigido, e o "Organizar o
    dia" agora **devolve ao cadastro** a duração já inflada.
  - **Fase fora do plano:** cada fase da grade pode ser marcada como "produzida
    em outro momento" (gola e viés), e o dia deixa de montar a corrente dela.
  - **Operação partida:** trabalho que entra numa hora marcada é dividido em
    "parte 1" e "parte 2" em vez de ir inteiro para depois da pausa.
  - **Nome da operação** passa a trazer a fase e a OS: `Enfesto · Corpo 1 · OS 453`.
  - **Retirar OS:** a janela mostra as OS alocadas em lista (era um seletor que
    escondia tudo) e aceita **marcar várias** de uma vez.
  - **Importar risco:** o **nome do arquivo** passa a identificar a fase — os três
    PDFs de corpo de um tricolor cadastravam uma fase só, porque a memória
    "modelo|tecido" é a mesma nas três.
  - Cópia do código deste ponto em
    `backups-codigo/*.20260803l-operacoes-enfesto.CÓPIA`.
- **DADOS:** exportação de **03/08/2026 17:32**, em
  `backups/BACKUP-COMPLETO-2026-08-03T20-32-59.json` (1,75 MB, 29 chaves, 764
  registros): **177 OS**, 66 grades, 25 desenhos, **288 operações**, 40 cargas de
  expedição, 41 mov. de estoque, 11 tecidos, 37 cores, 11 funções, 10 pessoas na
  equipe, 6 modelos, 15 etapas, 17 componentes, 4 materiais.
  A mesma cópia está em `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`, ao lado das
  duas de 30/07.

  Cresceu **113 operações** e 3 OS desde 30/07 — é a semana de planejamento do
  setor de enfesto/corte.

  > Restaurar: app → Configurações → **Importar JSON** → escolher este arquivo.
  > Sobrescreve tudo, então é o caminho de "perdi geral". Para casos parciais,
  > ver a seção 2 abaixo.

  **O que este backup mostra que ainda NÃO foi cadastrado.** O código deste ponto
  já aceita os três, mas os dados ainda não os têm — vale conferir isto antes de
  culpar o programa por um plano de operações torto:
  - `Operador de enfestadeira · Enfesto = 80 min` ✔, porém **sem comprimento de
    referência** preenchido: vale o padrão de 8 m, que por acaso é o pretendido;
  - **nenhuma fase de grade tem tempo de enfesto** (`enfestoMin`) — as 66 grades
    estão em branco nesse campo, então fase sem medição própria cai no tempo do
    posto ou na estimativa;
  - **nenhuma fase está marcada como "produzida em outro momento"** — gola e viés
    continuam entrando na corrente do dia em todas as grades.
- **SUPABASE:** sem mudança de infraestrutura neste ponto (mesmas tabelas e
  políticas da seção 3).

### Pontos mais antigos

- **CÓDIGO:** tag `restore-2026-07-30-k` (cache-buster `app.js ?v=2026-07-30k`,
  `styles.css ?v=2026-07-30d`). O que mudou desde o `restore-2026-07-28-aq`:
  - **Impressão:** a folha de OE volta a ter margem no papel, e a de OS sai com
    as cores e ocupando a A4 inteira. As margens de folha A4 passam a morar num
    lugar só (`:root` do styles.css), com **recuo padrão de 15 mm na esquerda**.
    A etiqueta (100×50 mm) não entra nessa regra.
  - **Etiquetas:** o botão de etiquetas da OS estava quebrado em toda OS
    (`pdf.output is not a function`); e conferir na janela "etiquetas (tela)"
    deixou de gravar por cima do PDF bom que estava na pasta.
  - **Cadastros:** o usuário comum passa a **ver** todos os cadastros em modo
    leitura (só o admin escreve).
  - **Tecidos:** campo **Excedente de enfesto (cm)** por tecido; os 15 cm viram
    apenas o padrão de quem não cadastrou.
  - **Risco (PDF):** o leitor do relatório do CAD passou a entender o layout de
    verdade — antes vinha sem comprimento, sem largura e com a tabela de
    tamanhos errada em todos os 132 PDFs. A janela ganhou: escolher entre
    **corrigir** uma grade existente ou **criar nova**, a lista completa de
    pastas (era de 3 opções e mandava PM.LISA para a pasta das CM), e a
    **previsão de fases** por produto (camiseta, polo, moletom liso, moletom
    tricolor, camiseta recortada).
  - **Grades:** a sequência da lista e da fila do assistente passa a ser **por
    semelhança** de faixa de tamanhos.
  - Cópia do código deste ponto em
    `backups-codigo/*.20260730k-previsao-fases.CÓPIA`.
- **DADOS:** exportação de **30/07/2026 17:23**, em
  `backups/BACKUP-COMPLETO-2026-07-30T20-23-43.json` (1,65 MB, 29 chaves, 641
  registros): 174 OS, **66 grades**, 25 desenhos, **175 operações**, 37 cargas de
  expedição, 38 mov. de estoque, **11 tecidos**, 37 cores.
  A mesma cópia está em `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`, junto com
  as de 29/07 e das 10:51.

  Fecha o dia inteiro — inclui o que foi feito na tarde de 30/07:
  - as duas grades cadastradas pela importação de risco: `M-G-GG-G1-G3 |
    PM.LISA` (pasta Camiseta Polo / básica) e `P ao G3 | CM.LISA | 116.5cm`;
  - os cinco tecidos já com **excedente de enfesto** cadastrado (Ribana Moletom
    15, Ribana Malha Algodão 5, Texturizado Prime 15, Texturizado Rugão 15,
    Piquet Dry 15 cm);
  - o tecido novo e as 22 operações planejadas a mais.

  > Restaurar: app → Configurações → **Importar JSON** → escolher este arquivo.
  > Sobrescreve tudo, então é o caminho de "perdi geral". Para casos parciais,
  > ver a seção 2 abaixo.

---

## Ponto anterior

- **CÓDIGO:** tag `restore-2026-07-28-aq` (cache-buster `app.js ?v=2026-07-28aq`,
  `styles.css ?v=2026-07-28k`). O que mudou desde o `restore-2026-07-28-an`:
  - **Importar risco (PDF)** ganhou o outro lado: além de corrigir grade
    existente, agora **cria grade nova** a partir dos relatórios do CAD,
    perguntando só o que o PDF não traz (SKU, tipo de peça, variação, tecido e
    nome de cada fase). O nome da grade sai dos tamanhos, na convenção da casa.
    Ele **aprende** o produto e o código de tecido de cada fase — a segunda grade
    do mesmo produto abre preenchida.
  - **Excedente de enfesto:** o comprimento do relatório é a medida de CORTAR; o
    cadastro recebe +15 cm (a de ENFESTAR). A largura não muda. Conferido: o
    Corpo Parte 1 da 2X P ao G3 casa exato (4,5493 + 0,15 = 4,70).
  - **Pasta das exportações (JSON)** em Configurações: o "Exportar tudo" grava
    direto numa pasta conectada, com nome datado, sem passar pelos Downloads.
- **DADOS:** exportação de **28/07/2026 17:31**, em
  `backups/BACKUP-COMPLETO-2026-07-28T20-31-01.json` (1,73 MB, 28 chaves, 610
  registros): 174 OS, 64 grades, 25 desenhos, 153 operações planejadas, 10
  funções, 37 cores, 10 tecidos, 33 cargas de expedição, 36 mov. de estoque.
  A mesma cópia está em `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`.
  Cópia do código deste ponto em `backups-codigo/*.20260728aq-risco-grade-nova.CÓPIA`.

  > ATENÇÃO: esta cópia ainda tem as OS **0340** e **0398** com número repetido
  > (a de junho da 340 e uma das 398 são cascas vazias, sem expedição, estoque
  > nem operação). O app avisa no console. Se elas já tiverem sido apagadas
  > depois desta hora, a próxima exportação sai limpa.

---

## 1) CÓDIGO (app.js, index.html, styles.css)

- **Onde está:** repositório GitHub `diversemodas1997-cmyk/Gerador-OS`, branch `main`.
- **Ponto de restauração:** tag `restore-2026-07-28-t` (cache-buster `t`).
  Anterior: `restore-2026-07-24-n`.
- **Como restaurar / reimplantar:** basta hospedar os 3 arquivos (index.html +
  app.js + styles.css) em qualquer servidor de estático. A cópia VIVA é este
  repo (tem o banner de cor). Ao editar o app.js, sempre suba o `?v=` no
  index.html (cache-buster) — hoje em `app.js ?v=2026-07-28t` /
  `styles.css ?v=2026-07-28h`.
- **Config do Supabase fica no topo do app.js:** `SUPA_URL` e `SUPA_KEY` (chave
  `anon`). Se o projeto Supabase mudar, troque esses dois valores.

---

## 2) DADOS (tudo que o usuário cadastrou)

Todos os dados vivem numa ÚNICA linha no Supabase: tabela `shared_data`,
`id = 'main'`, coluna `data` (JSON com todas as chaves: ordens, desenhos,
tecidos, cores, grades, etapas, componentes, funções, expedição, operações,
estoque, meta, osCounter…).

### Camadas de backup dos dados (redundância)

1. **Backup manual completo (este):**
   - `J:\Meu Drive\Backup ERP Diverse\Gerador-OS\os-gen-backup-1785184261956.json`
     (exportado em 27/07/2026 17:31 — é o mesmo arquivo abaixo)
   - `C:\Users\Pichau\Desktop\Gerador-OS\backups\BACKUP-COMPLETO-2026-07-27T17-31-01.json`
   - Formato **pronto pra importar** (chaves = arrays reais).
   - Anteriores na mesma pasta: `BACKUP-COMPLETO-2026-07-23T20-25-51.json` e a
     cópia bruta do snapshot ao lado (`snapshot-bruto-...json`).
2. **Snapshots de contingência (automáticos, por alteração):** pasta
   `snapshots/` dentro de cada pasta de backup/PDF conectada no Drive
   (ex.: `J:\Meu Drive\Backup ERP Diverse\Gerador-OS\snapshots\snap-*.json`).
   Guarda os últimos ~30 estados. Também no navegador (IndexedDB).
3. **Snapshots DIÁRIOS no servidor:** tabela `shared_data_backups` no Supabase
   (1 por dia, retenção 30 dias). Acessível em Configurações → snapshots.
4. **Backup JSON automático:** o app grava um `os-gen-backup-*.json` na pasta de
   backup conectada a cada save.

### Como restaurar os dados (Supabase de pé)

- **Tudo de uma vez:** app → **Configurações → Importar JSON** → escolher o
  `BACKUP-COMPLETO-*.json`. (⚠️ sobrescreve tudo — use quando perdeu geral.)
- **Só as OEs (expedição):** Configurações → "Restaurar só as OEs de um
  snapshot" → escolher um snapshot de antes da perda (mescla, não apaga).
- **Restaurar um dia:** Configurações → snapshots diários → Restaurar.
- **Só desenhos ou parte:** baixar um snapshot e reimportar por chave.

> A gravação usa **merge por chave** (concorrência otimista): um dispositivo com
> cache velho não apaga o que outro gravou. A trava anti-apagamento bloqueia
> gravar vazio sobre servidor com dados (OS, desenhos e expedição).

---

## 3) SUPABASE (infraestrutura) — necessário para recriar do zero

Se o projeto Supabase for perdido, é preciso recriar. O código só precisa de
`SUPA_URL` + chave `anon` (no app.js). Tabelas usadas pelo app:

| Tabela | Colunas (uso no código) | Papel |
|---|---|---|
| `shared_data` | `id` (text PK, 'main'), `data` (jsonb), `updated_at` (timestamptz), `updated_by` | Estado inteiro do app |
| `shared_data_backups` | `id`, `snapshot_date` (date), `created_at` (timestamptz), `data` (jsonb) | Snapshots diários |
| `user_roles` | `user_id` (uuid), `role` (text: 'admin'/'usuario') | Papéis |
| `skus_catalogo` | `id` (text, 'main'), `data` (jsonb) | Catálogo de SKUs (só leitura) |
| `compras_materiais` | (livre) | Compras da Contabilidade (só leitura) |

- **Realtime:** habilitar Realtime (postgres_changes) na tabela `shared_data`.
- **RLS (inferido do comportamento):** usuário **autenticado** pode `select`/
  `insert`/`update` em `shared_data` e `shared_data_backups`; **anon** é
  bloqueado (por isso ferramentas externas sem login não leem). `user_roles`
  legível pelo próprio usuário.
- **Auth:** contas de e-mail/senha do Supabase Auth. O admin principal é
  `diversemodas1997@gmail.com`. Ao recriar, cadastre os usuários e ponha o papel
  `admin` na `user_roles`.

### Passos de recriação total (pior caso)

1. Criar projeto Supabase novo; anotar URL + `anon key`.
2. Criar as tabelas acima; habilitar RLS com as políticas para `authenticated`.
3. Habilitar Realtime em `shared_data`.
4. Recriar os usuários no Auth e a linha de papel admin em `user_roles`.
5. Trocar `SUPA_URL`/`SUPA_KEY` no app.js (bumpar o `?v=`) e reimplantar.
6. Logar como admin → **Importar JSON** com o `BACKUP-COMPLETO-*.json`.

---

## Rotina recomendada de backup contínuo

- Manter uma **pasta de backup conectada** no app (Configurações) apontando pra
  dentro do `J:\Meu Drive` → gera snapshots automáticos por alteração.
- De tempos em tempos, **Exportar tudo (JSON)** e guardar com data.
- O código já fica versionado no GitHub a cada mudança.
