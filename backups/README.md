# Backups de dados — Gerador-OS

Cópias de segurança dos **dados** (cadastros/OS). O app guarda os dados no
Supabase (`shared_data`, id `main`); estes arquivos são cópias exportadas.

## Arquivos

- **desenhos-restaurados-2026-07-06.json** — os **24 desenhos técnicos**
  restaurados em 06/07/2026 (após a perda de dados). Contém só a chave
  `desenhos` (array). Para restaurar: app → Configurações → **⬆ Importar JSON**.
  Restaura apenas os desenhos, sem tocar no resto.

- **gerador-os-COMPLETO-2026-05-27.json** — cópia **completa** de todos os
  cadastros (tecidos, cores, modelos, grades, desenhos, ordens, etc.), tirada
  do ERP em 27/05/2026. Fonte usada para recuperar os desenhos. Serve para
  restaurar também os demais cadastros que sumiram, se necessário.
  Obs.: neste arquivo cada chave é uma **string JSON** (formato do blob).

## Como fazer um backup novo da versão atual (recomendado periodicamente)

No app, logado como admin: Configurações → **⬇ Exportar tudo (JSON)**. Guarde
o arquivo aqui com a data no nome. Além disso, o app já grava snapshots
diários automáticos no Supabase (retidos por 30 dias) e um backup local a cada
alteração (se a pasta de backup estiver conectada).
