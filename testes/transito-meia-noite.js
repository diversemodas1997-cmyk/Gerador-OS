/* Rode com:  node testes/transito-meia-noite.js

   O TRÂNSITO NÃO ATRAVESSA A MEIA-NOITE (09/10/2026, Junior: "é impossível
   começar qualquer período com qualquer número maior que zero na coluna Início
   em trânsito").

   No histórico (_dashIntervalos, que o Início e o Relatório produção usam), o
   intervalo de um cartão de trânsito fecha às 23:59:59 do dia em que começou.
   Se a OS ainda aparece na estrada num instante de outro dia (outra carga),
   abre um intervalo novo NAQUELE instante — nunca à meia-noite.

   Recorta do app.js a função real; a linha do tempo da OS vem pronta. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const corta = (de) => {
  const i = src.indexOf(de), j = src.indexOf('\n}', i);
  if (i < 0 || j < 0) { console.error('nao achei ' + de); process.exit(1); }
  return src.slice(i, j + 2);
};
const constante = (nome) => src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'))[0];

const H = (d, h, m) => new Date(2026, 9, d, h, m || 0).getTime();   // outubro/2026
const FIM = d => new Date(2026, 9, d + 1).getTime() - 1000;

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const _dashLinhaDoTempoOS = (o) => ctx.linhas[o.id];
  const _dashCartoesDaOS = () => [];
  const produtosOS = () => 0;
  const _dataFinalizacaoOS = () => '';
  const _expAgora = () => ctx.agora;
  ${constante('DASH_CARTOES_TRANSITO')}
  ${corta('function _dashIntervalos')}
  return _dashIntervalos;
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const ev = (t, cartoes) => ({ t, cartoes: new Map(Object.entries(cartoes)) });

// Saiu dia 2 às 09:30 e a chegada só foi marcada dia 7: fecha no fim do dia 2.
let ctx = { agora: H(9, 12), STATE: { ordens: [{ id: 'a', os: '0513' }] },
  linhas: { a: [ev(H(2, 9, 30), { idaManha: 100 }), ev(H(7, 17), { corteSC: 100 })] } };
let iv = monta(ctx)();
ok('o intervalo de transito fecha as 23:59:59 do dia em que comecou',
   iv.idaManha.length === 1 && iv.idaManha[0].ate === FIM(2), iv.idaManha);
ok('e o cartao de destino continua entrando na hora marcada', iv.corteSC[0].de === H(7, 17), iv.corteSC);

// Na estrada de novo num instante de outro dia (segunda carga, dia 5 às 14h).
ctx.linhas.a = [ev(H(2, 9, 30), { idaManha: 100 }), ev(H(5, 14), { idaManha: 100 }), ev(H(5, 18), { corteSC: 100 })];
iv = monta(ctx)();
ok('outra carga em outro dia abre intervalo novo no instante dela, nao a meia-noite',
   iv.idaManha.length === 2 && iv.idaManha[0].ate === FIM(2)
   && iv.idaManha[1].de === H(5, 14) && iv.idaManha[1].ate === H(5, 18), iv.idaManha);

// Ainda na estrada depois do ultimo instante.
ctx.linhas.a = [ev(H(8, 9, 30), { idaManha: 100 })];
iv = monta(ctx)();
ok('o que seguia na estrada fecha no fim do dia, se ele ja passou', iv.idaManha[0].ate === FIM(8), iv.idaManha);
ctx.agora = H(8, 15);
iv = monta(ctx)();
ok('no proprio dia da viagem segue aberto', iv.idaManha[0].ate === null, iv.idaManha);

// Os outros cartoes nao mudam: o corte de Descalvado atravessa dias a vontade.
ctx.agora = H(9, 12);
ctx.linhas.a = [ev(H(2, 9), { corte: 100 }), ev(H(7, 9), { costurando: 100 })];
iv = monta(ctx)();
ok('cartao que nao e de transito atravessa a meia-noite como sempre', iv.corte[0].ate === H(7, 9), iv.corte);

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
