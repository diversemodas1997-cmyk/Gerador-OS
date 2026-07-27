# Backup e Restauração — Gerador-OS

Guia para restaurar o sistema **sem perder dados**. O sistema tem 3 componentes
independentes: **CÓDIGO**, **DADOS** e **SUPABASE (infra)**. Cada um tem seu
backup próprio.

---

## Ponto de restauração deste backup

- **CÓDIGO:** tag `restore-2026-07-28-t` (cache-buster `app.js ?v=2026-07-28t`,
  `styles.css ?v=2026-07-28h`), 48 commits desde o `restore-2026-07-24-n`. O que
  mudou, por assunto:
  - **Sincronização (a correção mais importante):** o flush passou a fazer merge
    por REGISTRO nas listas, com base de três vias. Antes, cada chave subia
    inteira e uma aba desatualizada apagava a OS que outro aparelho tinha
    acabado de criar — foi a causa das perdas da OS 188.
  - **Planejamento de operações, reescrito em volta da OS:** o dia começa com
    "+ Alocar OS" e o programa monta a cascata dos 11 passos, cada um na função
    que o executa, com o tempo cadastrado nela. Corrente obrigatória do lote
    (com duas correntes paralelas: carga e matéria-prima), conflito de ordem,
    jornada 07:15–17:30 com transbordo para o próximo dia útil, pausas
    sincronizadas, tempo parado por função, botão "Organizar o dia" e a análise
    "Analisar o dia" com o que falta em cada lote.
  - **Funções:** cada função cadastra as suas operações com o tempo de cada uma
    (o campo de texto virou lista); as etapas marcadas deram lugar a essa lista.
  - **Grades:** tempo medido por fase de enfesto no cadastro, chave de grade
    unificada por identidade e conferência do horário lançado contra a média.
  - **Folhas:** OE com tabela da carga parcial e cor completa do tricolor;
    folha de operações sem o vão de 100mm, com somatório por função, o que falta
    em cada lote e quadrinho de checklist preenchível.
  - **Expedição:** marcar Ensaque não põe mais a OS em OE nenhuma — só o
    planejamento aloca; e a carga registra que foi feita.
- **DADOS:** o último snapshot exportado no app está em
  `backups/BACKUP-COMPLETO-2026-07-23T20-25-51.json` (161 OS, 25 desenhos, 64
  grades, 37 cores, 10 tecidos, 15 etapas, 17 componentes, 8 funções, 26 cargas
  de expedição, 6 operações, 18 mov. de estoque · osCounter 444). ⚠️ ESTE ARQUIVO
  ESTÁ DESATUALIZADO em relação ao uso de 24 a 28/07 (o planejamento de operações
  do dia 24 em diante não está nele). Para um backup de dados ATUAL, exporte pelo
  app (Configurações → Exportar tudo (JSON)): o código só carrega a chave `anon`
  do Supabase, que a RLS bloqueia para leitura fora do app logado — por isso a
  cópia de dados não é gerada automaticamente aqui.

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
   - `J:\Meu Drive\Backup ERP Diverse\Gerador-OS\BACKUP-COMPLETO-2026-07-23T20-25-51.json`
   - `C:\Users\Pichau\Desktop\Gerador-OS\backups\BACKUP-COMPLETO-2026-07-23T20-25-51.json`
   - Formato **pronto pra importar** (chaves = arrays reais).
   - Cópia bruta do snapshot ao lado (`snapshot-bruto-...json`).
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
