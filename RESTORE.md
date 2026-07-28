# Backup e Restauração — Gerador-OS

Guia para restaurar o sistema **sem perder dados**. O sistema tem 3 componentes
independentes: **CÓDIGO**, **DADOS** e **SUPABASE (infra)**. Cada um tem seu
backup próprio.

---

## Ponto de restauração deste backup

- **CÓDIGO:** tag `restore-2026-07-28-an` (cache-buster `app.js ?v=2026-07-28an`,
  `styles.css ?v=2026-07-28k`). O que mudou desde o `restore-2026-07-28-t`, por
  assunto:
  - **Operações, a cascata por FASE DO ENFESTO:** alocar uma OS monta a corrente
    uma vez por fase da grade (blusa moletom tricolor = 5 voltas de 9 operações,
    não 9 no lote inteiro). O tempo do enfesto passou a ser **apurado do
    histórico** — não se cadastra mais em Funções, porque depende do modelo, da
    grade e da fase. Operação criada pelo programa nasce com o responsável do
    posto, o que fez as linhas do tempo de quem trabalha em duas funções se
    mesclarem. Operação de **horário fixo** ("todo dia às") no cadastro de
    Funções, independente das OS. Botão de **retirar OS** em massa.
  - **Correções de cálculo:** camadas de ribana pelas "unidades da grade" (36 de
    moletom = 18 de ribana); divisão por tonalidade proporcional à grade
    (2M-4G-2GG dá G=48, não 24); quantidade dos componentes acompanhando as
    camadas lançadas na folha.
  - **Número da OS e a pasta de PDFs:** número canônico de 4 dígitos, arquivo
    `OS-<numero>.pdf` sem data (regravar substitui em vez de criar irmão) e
    recusa de número já usado. Resolveu 17 números com arquivo duplicado.
  - **Etiquetas:** imprimem pelo PDF, que carrega a página de 100×50 mm.
  - **Permissões:** só admin escreve, em todo o programa. A escalada a admin foi
    fechada NO SERVIDOR (`supabase-admin-roles.sql`, precisa ser rodado uma vez).
    Saiu o botão de "limpar todos os cadastros".
  - **Novidades de cadastro:** planilha `.xlsx` das grades agrupada por SKU, e
    **importação do risco do CAD em PDF**, que lança comprimento e largura nas
    fases da grade sozinha.
- **DADOS:** exportação de **28/07/2026 17:00**, em
  `backups/BACKUP-COMPLETO-2026-07-28T20-00-34.json` (1,73 MB, 28 chaves, 610
  registros): 174 OS, 64 grades, 25 desenhos, 153 operações planejadas, 10
  funções, 37 cores, 10 tecidos, 33 cargas de expedição. A mesma cópia está na
  pasta conectada `J:\Meu Drive\Backup ERP Diverse\Gerador-OS`.
  O anterior é `backups/BACKUP-COMPLETO-2026-07-27T17-31-01.json` (165 OS).
  Cópia do código deste ponto em `backups-codigo/*.20260728an-risco-pdf.CÓPIA`.

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
