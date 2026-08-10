# vendor/ — bibliotecas e fontes servidas pelo próprio site

Antes, `index.html` buscava estas bibliotecas em CDN (jsdelivr) e as fontes no
Google Fonts. Isso tinha uma consequência que só apareceria no pior momento: com
a **internet fora, o programa não abria** — mesmo com o servidor de dados da
fábrica funcionando perfeitamente, porque um `<script src>` que não baixa deixa
a página em branco.

Como o plano é rodar contra um servidor na rede local, nada pode depender de um
endereço externo. Tudo o que a página carrega mora aqui.

## O que tem

| arquivo | versão | origem |
|---|---|---|
| `supabase-js-2.112.2.min.js` | 2.112.2 | `@supabase/supabase-js@2` |
| `html2canvas-1.4.1.min.js` | 1.4.1 | `html2canvas@1.4.1` |
| `jspdf-2.5.2.umd.min.js` | 2.5.2 | `jspdf@2.5.2` |
| `pdf-3.11.174.min.js` | 3.11.174 | `pdfjs-dist@3.11.174` |
| `pdf.worker-3.11.174.min.js` | 3.11.174 | `pdfjs-dist@3.11.174` |
| `fonts/ibm-plex.css` + 16 `.woff2` | — | Google Fonts, subconjuntos `latin` e `latin-ext` |

A versão vai no nome do arquivo de propósito: assim dá para ver o que está
rodando sem abrir nada, e trocar de versão não deixa o navegador servindo a
antiga do cache.

O **worker do pdf.js** é caso à parte: ele não é carregado pelo `index.html`, e
sim buscado por rede no instante em que alguém lê um PDF de risco. O caminho
está em `app.js`, na função `_riscoLerPdf`. Era o último endereço externo do
programa, e o mais traiçoeiro — com a internet fora o app abria normal e só
quebrava naquele ponto específico.

## Conferir se está tudo servido

Abra `vendor-check.html` no navegador (pelo endereço do servidor, não pelo
arquivo). Ele carrega as quatro bibliotecas, busca o worker e testa as duas
fontes, escrevendo o resultado na tela. Serve também para validar um servidor
novo da fábrica: se algo não estiver sendo servido, aparece ali.

## Atualizar alguma delas

1. Baixe a versão nova com o número no nome do arquivo.
2. Aponte o `index.html` (e, no caso do pdf.js, também o `workerSrc` em `app.js`).
3. Apague a antiga e atualize a tabela acima.
4. Abra `vendor-check.html` e confirme.

Não sirva nada de CDN de novo, nem "só esta". Basta uma para o programa parar
de abrir quando a internet cair.
