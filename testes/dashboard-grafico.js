/* Rode com:  node testes/dashboard-grafico.js

   O GRÁFICO DE COLUNAS de cada passo do dashboard: UMA COLUNA POR OS.

   Ele nasceu agregado — uma barra por quadro — e agregado repetia o que os
   cartões logo acima já diziam, só que em desenho. Com uma coluna por OS ele
   passa a dizer o que nenhum número do painel dizia: QUAIS lotes estão ali e de
   que tamanho é cada um. Um passo com 11.000 peças pode ser uma OS gigante ou
   trinta pequenas, e a diferença entre as duas coisas é a diferença entre um dia
   de trabalho e um mês.

   É desenho, e desenho errado não dá erro: dá uma barra que mente sobre o
   tamanho da outra, e ninguém confere altura de retângulo com régua. Daí o teste.

   O que ele guarda:

     · uma coluna por OS, na ordem dos quadros e, dentro de cada um, da maior
       para a menor — é a leitura que se procura primeiro ("qual é o lote grande
       que está segurando este campo?");
     · a altura é PROPORCIONAL à maior coluna DO PRÓPRIO passo. A escala é do
       grupo, e não do painel: numa escala única, um passo de 200 peças ao lado
       de outro de 11.000 viraria uma linha rente ao chão;
     · um TETO de colunas, com as demais somadas numa só. O cartão Estoque tem
       264 OS: 264 riscos de um pixel não são um gráfico, são uma textura;
     · a dica do mouse diz de qual quadro é cada OS — sem isso, num passo de duas
       unidades, não haveria como saber de que lado está cada barra. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const i = src.indexOf('function _dashGraficoColunas');
const j = src.indexOf('\n}', i);
const iTeto = src.indexOf('const DASH_GR_MAX');
if (i < 0 || j < 0 || iTeto < 0) { console.error('nao achei o grafico no app.js'); process.exit(1); }
const desenhar = new Function('esc', 'cards',
  src.slice(iTeto, i) + src.slice(i, j + 2) + '\nreturn _dashGraficoColunas(cards);');
const TETO = Number((src.match(/const DASH_GR_MAX = (\d+)/) || [])[1]);

const esc = (x) => String(x == null ? '' : x)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const g = (cards) => desenhar(esc, cards);
// Um quadro do painel: nome e as OS que estao nele agora.
const quadro = (nome, lista) => ({
  nome, v: { pecas: lista.reduce((s, x) => s + x.pecas, 0), os: lista.length, lista }
});
const os = (n, pecas) => ({ os: n, pecas });
// Altura de cada barra, em % da régua do eixo (desde 16/09/2026 o gráfico é
// HTML no formato do Power BI, e a altura mora no style da barra).
const alturas = (h) => [...h.matchAll(/class="dash-barra[^"]*" style="height:([\d.]+)%/g)].map(m => Number(m[1]));
const rotulos = (h) => [...h.matchAll(/dash-barra-rot[^>]*>([^<]*)/g)].map(m => m[1]);
const dicas = (h) => [...h.matchAll(/class="dash-gr-tip"[^>]*>([^<]*)</g)].map(m => m[1]);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------- 1. quando NAO ha grafico ---------- */

ok('passo sem OS nenhuma não vira gráfico',
   g([quadro('A', []), quadro('B', [])]) === '');
ok('lista vazia não quebra', g([]) === '' && g(null) === '');
ok('OS de zero peça não inventa gráfico', g([quadro('A', [os('0500', 0)])]) === '');

/* ---------- 2. uma coluna por OS, na ordem dos quadros ---------- */

let h = g([
  quadro('Unidade Descalvado', [os('0525', 1200), os('0530', 744)]),
  quadro('Unidade São Carlos', [os('0516', 5280)])
]);
ok('três OS em dois quadros dão três colunas',
   rotulos(h).length === 3, rotulos(h));
ok('e saem na ordem dos quadros, a maior de cada um primeiro',
   rotulos(h).join(' ') === '0525 0530 0516', rotulos(h));
ok('a dica diz de qual quadro é cada OS',
   dicas(h)[0] === 'OS 0525 · Unidade Descalvado: 1.200 produtos'
   && dicas(h)[2] === 'OS 0516 · Unidade São Carlos: 5.280 produtos', dicas(h));

/* ---------- 3. a altura e proporcional a MAIOR do passo ---------- */

let a = alturas(h);
// A régua é "redonda" (5.280 → topo 6.000), como no Power BI: a maior coluna
// fica perto do topo, e não colada nele.
ok('a maior OS do passo mede contra a régua redonda (5.280 de 6.000)', Math.abs(a[2] - 88) < 0.01, a);
ok('e as outras ficam na proporção exata (1.200/5.280)',
   Math.abs(a[0] / a[2] - 1200 / 5280) < 0.005, { proporcao: a[0] / a[2] });
// A escala e do PASSO: os mesmos numeros, noutro grupo, dao a mesma figura.
const h2 = g([quadro('X', [os('1', 600), os('2', 372)]), quadro('Y', [os('3', 2640)])]);
// Com a régua redonda a figura não é idêntica (2.640 → topo 3.200; 5.280 → 6.000),
// mas o passo pequeno continua ocupando a altura dele, e não rente ao chão.
ok('a escala é do passo, não do painel: o passo pequeno também enche o gráfico',
   alturas(h2)[2] >= 75 && a[2] >= 75
   && Math.abs(alturas(h2)[0] / alturas(h2)[2] - a[0] / a[2]) < 0.005, { pequeno: alturas(h2), grande: a });

/* ---------- 4. o teto de colunas ---------- */

const muitas = Array.from({ length: 264 }, (_, k) => os(String(300 + k), 100 + k));
h = g([quadro('Produto acabado', muitas)]);
ok('264 OS não viram 264 riscos: o desenho para no teto',
   rotulos(h).length === TETO, rotulos(h).length);
ok('as maiores é que aparecem',
   rotulos(h)[0] === '563' && rotulos(h)[1] === '562', rotulos(h).slice(0, 3));
ok('e a última coluna diz quantas OS foram somadas nela',
   rotulos(h)[TETO - 1] === '+' + (264 - (TETO - 1)), rotulos(h)[TETO - 1]);
ok('a coluna do resto se explica na dica do mouse',
   /^\d+ OS menores, somadas: [\d.]+ produtos$/.test(dicas(h)[TETO - 1]), dicas(h)[TETO - 1]);
ok('e sai marcada como resto, para a cor não a confundir com um lote',
   (h.match(/dash-barra resto/g) || []).length === 1, h.match(/class="[^"]*"/g));
// Exatamente no teto, nada e agrupado.
h = g([quadro('A', Array.from({ length: TETO }, (_, k) => os('X' + k, 10 + k)))]);
ok('exatamente no teto, nenhuma OS é agrupada',
   rotulos(h).length === TETO && !/dash-barra resto/.test(h), rotulos(h).length);

/* ---------- 5. o desenho e honesto com a tela ---------- */

h = g([quadro('A', [os('0001', 10)]), quadro('B', [os('0002', 20)])]);
ok('as colunas são flexíveis, sem largura fixa (o painel encolhe até o celular)',
   (h.match(/class="dash-col"/g) || []).length === 2 && !/width="\d/.test(h), h.slice(0, 200));

/* ---------- 6. o formato do Power BI ---------- */

h = g([quadro('Unidade Descalvado', [os('0525', 1200)]), quadro('Unidade São Carlos', [os('0516', 5280)])]);
ok('o eixo tem 5 marcas, do zero ao topo redondo',
   [...h.matchAll(/dash-gr-ytick[^>]*>([^<]*)/g)].map(m => m[1]).join(' | ') === '0 | 1,5 mil | 3 mil | 4,5 mil | 6 mil',
   [...h.matchAll(/dash-gr-ytick[^>]*>([^<]*)/g)].map(m => m[1]));
ok('cada coluna traz o valor em cima',
   [...h.matchAll(/dash-col-val[^>]*>([^<]*)/g)].map(m => m[1]).join(' ') === '1,2 mil 5,3 mil',
   [...h.matchAll(/dash-col-val[^>]*>([^<]*)/g)].map(m => m[1]));
ok('a cor é do quadro: Descalvado s1, São Carlos s2',
   /dash-barra s1/.test(h) && /dash-barra s2/.test(h), h.match(/dash-barra[^"]*/g));
ok('dois quadros com colunas: tem legenda',
   (h.match(/<i class="s\d"><\/i>/g) || []).length === 2, h);
h = g([quadro('Unidade Descalvado', [os('0525', 1200)]), quadro('Unidade São Carlos', [])]);
ok('um quadro só com colunas: sem legenda de um item',
   !/dash-gr-leg/.test(h), h);
// Nome com aspas ou sinal de menor nao pode escapar para dentro do SVG.
h = g([quadro('A "B" <C>', [os('0001', 10)]), quadro('D', [os('0002', 20)])]);
ok('o nome do quadro sai escapado, e não vira marcação solta',
   !h.includes('<C>') && h.includes('&quot;B&quot;'), dicas(h));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
