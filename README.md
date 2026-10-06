# Gerador-OS

Gerador de Ordem de Serviço para confecção (Diverse/Dixie). É um app de página
única, servido como arquivo estático: os três arquivos da raiz são o programa
inteiro.

## O que fica na raiz, e por quê

| Arquivo | Papel |
|---|---|
| `index.html` | A página. É a porta de entrada do site — o servidor a procura na raiz. |
| `app.js` | Todo o programa. Carregado por `index.html` com um `?v=` que precisa ser trocado a cada mudança (cache-buster). |
| `styles.css` | Todo o visual, inclusive as folhas de impressão (OS, OE, etiqueta). |

Esses três não podem ir para dentro de uma pasta: o endereço publicado aponta
para a raiz, e `index.html` chama os outros dois ao lado dele.

## As pastas

| Pasta | O que guarda |
|---|---|
| `docs/` | [Backup e restauração](docs/RESTORE.md) — como voltar atrás quando algo se perde, e o histórico dos pontos de restauração. |
| `sql/` | Scripts do Supabase, para rodar no SQL Editor: papéis de admin, políticas de RLS e a tabela de compras da Contabilidade. |
| `dados/` | JSONs avulsos de reparo e restauração pontual, e a leitura de um relatório de risco guardada como referência. Não são backup — são remendos datados. |
| `backups/` | Cópias `shared_data-antes-<correção>-<data>.json`, gravadas pelos scripts de `servidor/` antes de mexer nos dados. Fora do git, por tamanho. O backup diário apaga as com mais de 14 dias. As exportações completas antigas (`BACKUP-COMPLETO-*.json`) ficam no Drive, em `Backup ERP Diverse\Gerador-OS`. |
| `servidor/` | Instalação e manutenção do servidor da fábrica, os backups diários e os scripts de correção de dados. |
| `Desenhos técnicos -grades de corte/` | Os riscos em PDF, por linha (BM.LISA, BM.TRI, CM.LISA, CM.TRI, PM.LISA), depois por grade e por largura do tecido. Nome do arquivo: `<LINHA> - <PEÇA> <GRADE>.pdf`. |

## Onde estão os dados

Não estão aqui. Todo o estado do programa vive numa única linha do Supabase do
servidor da fábrica (`shared_data`, `id = 'main'`), compartilhada por todos os
usuários. Os backups diários vão para o Google Drive — o que existe, onde e
como restaurar está em [docs/RESTORE.md](docs/RESTORE.md).
