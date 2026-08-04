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
| `backups/` | Exportações completas dos dados (`BACKUP-COMPLETO-<data>.json`) e as cópias automáticas do app. Fora do git, por tamanho. |
| `backups-codigo/` | Cópia do `app.js`/`index.html`/`styles.css` de cada ponto de restauração. |
| `Desenhos técnicos -grades de corte/` | Os riscos em PDF, por linha (BM.LISA, BM.TRI, CM.LISA, CM.TRI, PM.LISA), depois por grade e por largura do tecido. Nome do arquivo: `<LINHA> - <PEÇA> <GRADE>.pdf`. |

## Onde estão os dados

Não estão aqui. Todo o estado do programa vive numa única linha do Supabase
(`shared_data`, `id = 'main'`), compartilhada por todos os usuários. O que há
neste repositório é o código e as cópias de segurança — ver
[docs/RESTORE.md](docs/RESTORE.md).
