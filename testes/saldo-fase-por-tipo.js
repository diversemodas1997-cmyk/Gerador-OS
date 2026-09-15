/* Rode com:  node testes/saldo-fase-por-tipo.js

   O SALDO DE UM CAMPO ABERTO POR TIPO DE PRODUTO.

   Os campos de costura passaram a mostrar uma tabela de saldo POR TIPO
   (CM.LISA, BM.TRI, CO.JAGUAR…) em vez de uma só com tudo misturado. As tabelas
   são a MESMA conta de sempre — calcularSaldosFase — com um filtro de OS por
   cima, e é isso que este teste trava: abrir por tipo não pode inventar nem
   perder peça.

   O ponto delicado são os LANÇAMENTOS MANUAIS. Eles são por tecido e cor, sem
   OS e portanto sem tipo de produto: somados dentro de cada tabela, o mesmo
   ajuste seria contado uma vez por tipo e o campo passaria a ter mais peça do
   que tem. Por isso saem das tabelas por tipo (`semMov`) e ganham a sua
   (`soMov`) — e a soma das partes tem que devolver o total. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = [
  corta('function _normNome'),
  corta('function osEtapaMarcada'),
  corta('function componentesPorTecidoCorOS'),
  recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades'),
  cortaArr('const FASES_ESTOQUE'),
  // O campo "Estoque de corte" so conta OS com o status Ensacado (a `cond` da
  // fase), entao o motor precisa saber ler o status.
  cortaArr('const STATUS_OS'),
  cortaLinha('const STATUS_FIM'),
  corta('function _marcasDoStatus'),
  corta('function _statusDoChecklistOS'),
  corta('function _ultimaMarcacaoChecklist'),
  corta('function _statusOS'),
  corta('function _faseEntrouOS'),
  corta('function _nomeEtapaDaFase'),
  cortaLinha('function _faseIdxPorId'),
  cortaArr('const _TRANSITO_PERNAS'),
  corta('function _fracoesMovidasOS'),
  corta('function _transitoDaOS'),
  corta('function _fracaoRecebida'),
  corta('function _fracaoPerdida'),
  corta('function _expIso'),
  corta('function _expHoje'),
  corta('function _expDataEfetivaCarga'),
  corta('function calcularSaldosFase'),
  cortaLinha('const TERMINAL_ETAPA_RE'),
  corta('function faseAtualOS'),
  corta('function _expCancelSet'),
  corta('function _expEmbarcadoOS'),
  corta('function _skuDaGrade'),
  corta('function _skuDaOS')
].join('\n');

// Roda uma expressão contra o motor real, com o STATE dado.
function rodar(estado, expr) {
  const fn = new Function('STATE', `
    const corCanonicaPorTecido = (cor) => cor || '';
    const _gradeIdDaOS = (o) => o.gradeId || '';
    function _expPecasPacoteOS() {
      const mapa = new Map();
      return { mapa, total: 0, de: () => 0 };
    }
    ${motor}
    const IDX = FASES_ESTOQUE.findIndex(f => f.id === 'costurando');
    const soma = (d) => d.reduce((s, c) => s + c.estoque, 0);
    return (${expr});
  `);
  return fn(estado);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ---------------------- */

// Duas grades de tipos diferentes. O TIPO é o pedaço do meio do nome da grade —
// é de lá que _skuDaGrade o lê, porque não existe campo próprio para ele.
const GRADES = [
  { id: 'g1', nome: 'P ao G3 | CM.LISA | 117cm' },
  { id: 'g2', nome: 'P ao G3 | BM.TRI | 177cm' },
  { id: 'g3', nome: 'sem tipo nenhum' }        // nome fora do padrão: sem tipo
];

// Cada OS entra na costura de Descalvado (Costura marcada, sem passar por São
// Carlos) com as peças pedidas.
const os = (id, num, gradeId, pecas, tecido, cor) => ({
  id, os: num, modeloNome: 'Peça', data: '2026-09-14', gradeId,
  etapas: ['Corte', 'Costura'],
  progresso: { etapasCheck: { 'Corte': true, 'Costura': true }, etapasSeq: { 'Corte': 1, 'Costura': 2 } },
  componentes: [{ materialNome: tecido, corNome: cor, qtdTotal: pecas }]
});

const estado = (ordens, movs) => ({
  ordens,
  grades: GRADES,
  expedicaoCargas: [], expedicaoJanelas: [], expedicaoExcecoes: [],
  corteMov: [], costurandoMov: movs || [], corteScMov: [], costurandoScMov: [],
  fiosMov: [], expedicaoMov: []
});

const ORDENS = [
  os('a', '0501', 'g1', 100, 'Malha Algodão', 'Preto Malha Algodão'),
  os('b', '0502', 'g1', 50, 'Malha Algodão', 'Branco Malha Algodão'),
  os('c', '0503', 'g2', 200, 'Moletom Bulk', 'Bege Moletom Bulk'),
  os('d', '0504', 'g3', 30, 'Malha Algodão', 'Preto Malha Algodão')
];
const MOVS = [
  { id: 'm1', tipo: 'saida', qtd: 20, tecidoNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', data: '2026-09-15' }
];

/* ---------- 1. o tipo sai da GRADE cadastrada ---------- */

ok('o tipo de uma OS é o pedaço do meio do nome da grade',
  rodar(estado(ORDENS), `[_skuDaOS(STATE.ordens[0]), _skuDaOS(STATE.ordens[2]), _skuDaOS(STATE.ordens[3])]`)
    .join('|') === 'CM.LISA|BM.TRI|',
  rodar(estado(ORDENS), `[_skuDaOS(STATE.ordens[0]), _skuDaOS(STATE.ordens[2]), _skuDaOS(STATE.ordens[3])]`));

/* ---------- 2. cada tabela mostra só o seu tipo ---------- */

const porTipo = t => `soma(calcularSaldosFase(IDX, { semMov: true, filtroOS: o => (_skuDaOS(o) || '') === '${t}' }).detalhe)`;

ok('CM.LISA vê as suas 150 pç, e só elas',
  rodar(estado(ORDENS, MOVS), porTipo('CM.LISA')) === 150,
  rodar(estado(ORDENS, MOVS), porTipo('CM.LISA')));

ok('BM.TRI vê as suas 200 pç',
  rodar(estado(ORDENS, MOVS), porTipo('BM.TRI')) === 200,
  rodar(estado(ORDENS, MOVS), porTipo('BM.TRI')));

// Grade com nome fora do padrão não tem tipo — a OS não pode sumir por isso.
ok('a OS de grade sem tipo cai na tabela "sem tipo", e não no vazio',
  rodar(estado(ORDENS, MOVS), porTipo('')) === 30,
  rodar(estado(ORDENS, MOVS), porTipo('')));

/* ---------- 3. o lançamento manual não se multiplica ---------- */

ok('semMov deixa o ajuste manual de fora de cada tabela por tipo',
  rodar(estado(ORDENS, MOVS), porTipo('CM.LISA')) === 150,
  rodar(estado(ORDENS, MOVS), porTipo('CM.LISA')));

ok('soMov traz SÓ o ajuste manual (−20)',
  rodar(estado(ORDENS, MOVS), `soma(calcularSaldosFase(IDX, { soMov: true }).detalhe)`) === -20,
  rodar(estado(ORDENS, MOVS), `soma(calcularSaldosFase(IDX, { soMov: true }).detalhe)`));

// O fecho de tudo: as partes somam o total, e o total é o que a tela sempre
// mostrou. Se abrir por tipo tivesse contado o ajuste duas vezes, é aqui que
// apareceria.
const total = rodar(estado(ORDENS, MOVS), `soma(calcularSaldosFase(IDX).detalhe)`);
const partes = ['CM.LISA', 'BM.TRI', ''].reduce((s, t) => s + rodar(estado(ORDENS, MOVS), porTipo(t)), 0)
  + rodar(estado(ORDENS, MOVS), `soma(calcularSaldosFase(IDX, { soMov: true }).detalhe)`);
ok('a soma das tabelas por tipo + a dos manuais é igual ao total do campo',
  partes === total && total === 360, { partes, total });

/* ---------- 4. a conta antiga não mudou ---------- */

ok('sem opts, calcularSaldosFase continua sendo o total de sempre',
  rodar(estado(ORDENS), `soma(calcularSaldosFase(IDX).detalhe)`) === 380,
  rodar(estado(ORDENS), `soma(calcularSaldosFase(IDX).detalhe)`));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
