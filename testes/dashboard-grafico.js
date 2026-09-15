/* Rode com:  node testes/dashboard-grafico.js

   O GRÁFICO DE COLUNAS de cada passo do dashboard.

   Os cartões dizem o número; o gráfico diz a PROPORÇÃO — qual metade do passo
   está carregada. É desenho, e desenho errado não dá erro: dá uma barra que
   mente sobre o tamanho da outra, e ninguém confere altura de retângulo com
   régua. Daí o teste.

   O que ele guarda:

     · a altura é PROPORCIONAL à maior coluna DO PRÓPRIO passo. A escala é do
       grupo, e não do painel: numa escala única, um passo de 200 peças ao lado
       de outro de 11.000 viraria uma linha rente ao chão;
     · coluna de valor zero vira um TRAÇO, e não some. O zero é resposta — "aqui
       não tem nada agora" —, e sumir com ela faria a barra ao lado parecer a
       única que existe;
     · sem o que comparar (uma coluna só, ou tudo zerado) não há gráfico. Gráfico
       que não compara é enfeite ocupando a tela;
     · o rótulo perde o prefixo que se repete em todas as colunas ("Unidade",
       "Recebido em") — ele é o que a linha do passo já disse, e repetido em cada
       coluna só rouba a largura de que o nome precisa. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const i = src.indexOf('function _dashGraficoColunas');
const j = src.indexOf('\n}', i);
if (i < 0 || j < 0) { console.error('nao achei _dashGraficoColunas no app.js'); process.exit(1); }
const desenhar = new Function('esc', 'cards',
  src.slice(i, j + 2) + '\nreturn _dashGraficoColunas(cards);');

const esc = (x) => String(x == null ? '' : x)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const g = (cards) => desenhar(esc, cards);
const card = (nome, pecas) => ({ nome, v: { pecas, os: 1 } });
const alturas = (h) => [...h.matchAll(/height="([\d.]+)"/g)].map(m => Number(m[1]));

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------- 1. quando NAO ha grafico ---------- */

ok('uma coluna só não vira gráfico (uma barra sozinha é sempre 100%)',
   g([card('Produto acabado', 500)]) === '');
ok('tudo zerado não vira gráfico',
   g([card('A', 0), card('B', 0)]) === '');
ok('lista vazia não quebra', g([]) === '' && g(null) === '');

/* ---------- 2. a altura e proporcional a MAIOR do passo ---------- */

// 1.944 contra 5.280: a menor tem de ficar em 36,8% da maior.
let h = g([card('Unidade Descalvado', 1944), card('Unidade São Carlos', 5280)]);
let a = alturas(h);
ok('duas colunas: a maior ocupa a altura toda', a[1] === 72, a);
ok('e a menor fica na proporção exata (1.944/5.280 = 36,8%)',
   Math.abs(a[0] / a[1] - 1944 / 5280) < 0.005, { proporcao: a[0] / a[1], esperado: 1944 / 5280 });

// A escala e do PASSO: os mesmos numeros, noutro grupo, dao a mesma figura.
const h2 = g([card('X', 972), card('Y', 2640)]);
ok('a escala é do passo, não do painel: 972/2.640 desenha igual a 1.944/5.280',
   Math.abs(alturas(h2)[0] - a[0]) < 0.01, { pequeno: alturas(h2), grande: a });

// O ROTULO e so o que esta no eixo: o nome INTEIRO segue na dica do mouse de
// cada coluna, que e onde ele nao rouba largura de ninguem.
const rotulos = (h) => [...h.matchAll(/dash-barra-rot[^>]*>([^<]*)/g)].map(m => m[1]);

/* ---------- 3. o zero e um traco, e nao um sumico ---------- */

h = g([card('Ida · manhã', 0), card('Ida · tarde', 120),
       card('Volta · manhã', 0), card('Volta · tarde', 60)]);
a = alturas(h);
ok('quatro turnos: quatro colunas, nenhuma sumiu', a.length === 4, a);
ok('as de zero viram traço rente ao eixo', a[0] === 1 && a[2] === 1, a);
ok('e saem marcadas como vazias, para a cor dizer que são zero',
   (h.match(/dash-barra vazia/g) || []).length === 2, h.match(/class="[^"]*"/g));
ok('a de 60 fica na metade da de 120', Math.abs(a[3] / a[1] - 0.5) < 0.005, a);

/* ---------- 4. os rotulos ---------- */

h = g([card('Unidade Descalvado', 10), card('Unidade São Carlos', 20)]);
ok('o rótulo perde o "Unidade", que se repete em todas as colunas',
   rotulos(h).join('|') === 'Descalvado|São Carlos', rotulos(h));
h = g([card('Recebido em Descalvado', 10), card('Recebido em São Carlos', 20)]);
ok('e perde o "Recebido em" pelo mesmo motivo',
   rotulos(h).join('|') === 'Descalvado|São Carlos', rotulos(h));
ok('mas o nome inteiro segue na dica do mouse da coluna',
   h.includes('<title>Recebido em Descalvado: 10 peças</title>'),
   h.match(/<title>[^<]*<\/title>/g));

/* ---------- 5. o desenho e honesto com a tela ---------- */

h = g([card('A', 10), card('B', 20)]);
ok('as larguras são em % (o painel encolhe até o celular)',
   /width="\d+\.\d+%"/.test(h) && /x="\d+\.\d+%"/.test(h), h.slice(0, 200));
ok('cada coluna carrega o número na dica do mouse',
   h.includes('<title>A: 10 peças</title>') && h.includes('<title>B: 20 peças</title>'),
   h.match(/<title>[^<]*<\/title>/g));
// Nome com aspas ou sinal de menor nao pode escapar para dentro do SVG.
h = g([card('A "B" <C>', 10), card('D', 20)]);
ok('o nome do cartão sai escapado, e não vira marcação solta',
   !h.includes('<C>') && h.includes('&quot;B&quot;'), h.match(/<title>[^<]*<\/title>/g));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
