/* Rode com:  node testes/ranking-quadro-unico.js

   O RANKING VIROU UM QUADRO SO (22/09/2026, Junior: "transforme todos os
   quadros do ranking de producao em um unico quadro, mas com as variaveis de
   grade, tipo, tamanho, cor e sku em filtros para o usuario escolher o que
   deseja ver no quadro com as variaveis cruzadas").

   Seis quadros com seis contas paralelas viraram UM monte de fatos agrupado de
   maneiras diferentes. O que este teste guarda e justamente o que a conta unica
   nao pode errar:

     - somar os fatos de uma OS devolve o total dela. E isso que faz QUALQUER
       agrupamento fechar com o mesmo total geral - foi o teste que se fez a mao
       no banco de 21/09: 170.563 produtos em 2026, por grade, por tipo, por
       tamanho, por cor, por SKU e por periodo;
     - a OS de duas cores reparte os produtos entre elas (contar o lote inteiro
       nas duas dobraria a fabrica);
     - dentro da OS os produtos se repartem pelos TAMANHOS na proporcao da folha;
     - OS sem distribuicao por tamanho cai em "(sem tamanho)" e nao some;
     - a contagem de OS e por numero DISTINTO e nao se soma: a mesma OS aparece
       em varios tamanhos;
     - tamanho sai na ordem da grade (P, M, G, GG, G1, G2, G3), e nao na
       alfabetica; cor e grade saem por VOLUME, a maior primeiro.

   Recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
/* Anda por LINHA ate o fim da declaracao: o app.js e gravado ora com LF, ora
   com CRLF (o git converte no checkout), e um corte que procura ";\n" devolve
   vazio no dia em que o arquivo esta com CRLF. */
function pegaConst(nome) {
  const i = src.search(new RegExp('^const ' + nome + ' = ', 'm'));
  if (i < 0) { console.error('nao achei a constante ' + nome); process.exit(1); }
  const out = [];
  for (const l of src.slice(i).split(/\r?\n/)) {
    out.push(l);
    if (l.replace(/\/\/[^\r\n]*$/, '').trimEnd().endsWith(';')) return out.join('\n');
  }
  console.error('nao achei o fim da constante ' + nome);
  process.exit(1);
}

/* O que resolve SKU, produtos e a distribuicao por tamanho entra DUBLADO: cada
   um tem teste proprio, e o que se prova aqui e o que o ranking faz com a
   resposta deles. */
function monta(mundo) {
  return new Function('M', `
    var STATE = M.STATE;
    function skusDaOS(o) { return M.skus[o.id] || []; }
    function produtosOS(o) { return M.produtos[o.id] || 0; }
    function totaisPorTamanhoTomOS(o) { return M.tam[o.id] || { totalGeral: 0, tamanhos: [], colTotal: () => 0 }; }
    function corNomeCurto(n) { return n; }
    function _gradeIdDaOS(o) { return o.gradeId || ''; }
    ${pegaConst('RANK_VARS')}
    ${pegaConst('_RANK_ORDEM_TAM')}
    ${pegaFuncao('_rankingFatos')}
    ${pegaFuncao('_rankOrdenarValores')}
    ${pegaFuncao('_rankingValores')}
    ${pegaFuncao('_rankOrdenarOS')}
    ${pegaFuncao('_rankingQuadro')}
    return { _rankingFatos, _rankingQuadro, _rankingValores };
  `)(mundo);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(Number(a) - Number(b)) < 0.001;

/* ---------------------- o mundo do teste ----------------------
   Tres OS no mesmo ano:
     0001  CM.LISA-PRE          600 produtos, P/M/G em 1-2-3
     0002  CM.LISA-PRE + -BRA   400 produtos (200 por cor), P/G em 1-1
     0003  BM.TRI-VER           100 produtos, SEM distribuicao por tamanho   */
const mundo = () => ({
  STATE: {
    ordens: [
      { id: 'a', os: '0001', data: '2026-03-10', gradeId: 'g1' },
      { id: 'b', os: '0002', data: '2026-04-11', gradeId: 'g1' },
      { id: 'c', os: '0003', data: '2026-04-20', gradeId: 'g2' }
    ],
    grades: [{ id: 'g1', nome: 'P ao G3 | CM.LISA | 117cm' },
             { id: 'g2', nome: 'GG  ao G3 | BM.TRI | 177cm' }],
    cores: [{ nome: 'Preto', siglaSku: 'PRE' }, { nome: 'Branco', siglaSku: 'BRA' },
            { nome: 'Vermelho', siglaSku: 'VER' }]
  },
  skus: { a: ['CM.LISA-PRE'], b: ['CM.LISA-PRE', 'CM.LISA-BRA'], c: ['BM.TRI-VER'] },
  produtos: { a: 600, b: 400, c: 100 },
  tam: {
    a: { totalGeral: 600, tamanhos: ['p', 'm', 'g'], colTotal: k => ({ p: 100, m: 200, g: 300 })[k] || 0 },
    b: { totalGeral: 400, tamanhos: ['p', 'g'], colTotal: k => ({ p: 200, g: 200 })[k] || 0 },
    c: { totalGeral: 0, tamanhos: [], colTotal: () => 0 }
  }
});

const M = mundo();
const api = monta(M);
const base = api._rankingFatos('2026', '');

/* ---------- 1. os fatos fecham com o total ---------- */
const soma = base.fatos.reduce((s, f) => s + f.produtos, 0);
ok('1. a soma de todos os fatos e o total produzido (1.100)', perto(soma, 1100), soma);
const daB = base.fatos.filter(f => f.os === '0002').reduce((s, f) => s + f.produtos, 0);
ok('2. a OS de duas cores soma o lote dela UMA vez (400, nao 800)', perto(daB, 400), daB);
const bPreto = base.fatos.filter(f => f.os === '0002' && f.cor === 'Preto')
  .reduce((s, f) => s + f.produtos, 0);
ok('3. e reparte os produtos entre as duas cores (200 cada)', perto(bPreto, 200), bPreto);

/* ---------- 2. a repartição por tamanho é a da folha ---------- */
const aG = base.fatos.find(f => f.os === '0001' && f.tamanho === 'G');
ok('4. o tamanho G da 0001 leva 300 dos 600 (a proporcao da folha)',
   aG && perto(aG.produtos, 300), aG);
ok('5. a OS sem distribuicao cai em "(sem tamanho)" e nao some',
   base.fatos.filter(f => f.os === '0003').every(f => f.tamanho === '(sem tamanho)')
   && perto(base.fatos.filter(f => f.os === '0003').reduce((s, f) => s + f.produtos, 0), 100),
   base.fatos.filter(f => f.os === '0003'));

/* ---------- 3. todo agrupamento fecha o mesmo total ---------- */
['grade', 'tipo', 'tamanho', 'cor', 'sku', 'periodo'].forEach(eixo => {
  const q = api._rankingQuadro(base.fatos, eixo, '');
  ok('6. linhas=' + eixo + ' fecha em 1.100 produtos e 3 OS',
     q.total === 1100 && q.totalOS === 3, { total: q.total, os: q.totalOS });
});

/* ---------- 4. cruzado: as celulas somam o total ---------- */
const q = api._rankingQuadro(base.fatos, 'tamanho', 'cor');
const somaCel = q.linhas.reduce((s, l) =>
  s + q.colunas.reduce((t, c) => t + ((q.cel(l.rotulo, c.rotulo) || { produtos: 0 }).produtos), 0), 0);
ok('7. no cruzado, a soma das celulas e o total', somaCel === q.total, { somaCel, total: q.total });
ok('8. P x Preto = 100 (da 0001) + 100 (metade dos 200 P da 0002) = 200',
   (q.cel('P', 'Preto') || {}).produtos === 200, q.cel('P', 'Preto'));
ok('9. a celula sabe quais OS a formaram',
   ((q.cel('P', 'Preto') || {}).os || []).join(',') === '0001,0002', q.cel('P', 'Preto'));

/* ---------- 5. a contagem de OS nao se soma ---------- */
const somaOS = q.linhas.reduce((s, l) => s + l.os.length, 0);
ok('10. somar as OS das linhas daria mais do que existem — por isso o total e por numero distinto',
   somaOS > q.totalOS && q.totalOS === 3, { somaOS, totalOS: q.totalOS });

/* ---------- 6. a ordem de cada variavel ---------- */
ok('11. tamanho sai na ordem da grade, com "(sem tamanho)" no fim',
   q.linhas.map(l => l.rotulo).join(' ') === 'P M G (sem tamanho)',
   q.linhas.map(l => l.rotulo));
ok('12. cor sai por VOLUME, a maior primeiro (Preto 800, Branco 200, Vermelho 100)',
   q.colunas.map(c => c.rotulo).join(' ') === 'Preto Branco Vermelho',
   q.colunas.map(c => c.rotulo + '=' + c.produtos));

/* ---------- 7. o filtro recorta ---------- */
const soPreto = base.fatos.filter(f => f.cor === 'Preto');
const qp = api._rankingQuadro(soPreto, 'tamanho', 'grade');
ok('13. recortado por cor=Preto sobram 800 produtos em 2 OS',
   qp.total === 800 && qp.totalOS === 2, { total: qp.total, os: qp.totalOS });

/* ---------- 8. os filtros oferecem so o que existe ---------- */
const v = api._rankingValores(base.fatos);
ok('14. o filtro de cor oferece as tres cores do periodo',
   v.cor.length === 3 && v.cor.includes('Preto'), v.cor);
ok('15. a grade chega sem o espaco duplo da digitacao ("GG  ao G3" = "GG ao G3")',
   v.grade.includes('GG ao G3'), v.grade);
ok('16. o periodo nao vira filtro de variavel (quem filtra periodo e o ano/mes)',
   v.periodo === undefined, Object.keys(v));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
