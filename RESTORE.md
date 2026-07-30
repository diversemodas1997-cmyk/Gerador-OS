# Backup e Restauração — Gerador-OS

Guia para restaurar o sistema **sem perder dados**. O sistema tem 3 componentes
independentes: **CÓDIGO**, **DADOS** e **SUPABASE (infra)**. Cada um tem seu
backup próprio.

---

## Ponto de restauração deste backup

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
