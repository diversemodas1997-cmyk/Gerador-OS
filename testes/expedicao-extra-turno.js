/* Rode com:  node testes/expedicao-extra-turno.js

   MAIS UMA EXPEDIÇÃO NO MESMO TURNO (06/11/2026): o botão "+ Expedição no
   turno" cria uma janela de data fixa no dia da OE, no turno dela, em outro
   horário.
     · o turno sai da hora da ida: até 11:59 é manhã, de 12:00 em diante, tarde;
     · as horas de ida já usadas no dia são as das expedições não canceladas
       que acontecem NAQUELE dia (a remarcada para outro dia não ocupa).

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(de, ate) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + de + ' no app.js'); process.exit(1); }
  return src.slice(i, j + ate.length);
}

let falhas = 0;
const ok = (cond, msg) => { if (cond) console.log('ok   ' + msg); else { falhas++; console.log('FALHA ' + msg); } };

const OCS = [
  { data: '2026-11-06', horaIda: '07:00' },
  { data: '2026-11-06', horaIda: '13:30' },
  { data: '2026-11-06', horaIda: '09:00', cancelada: true },   // cancelada: não ocupa
  { data: '2026-11-07', horaIda: '08:00' }                     // remarcada para o dia seguinte
];
const f = new Function('OCS', `
  const ocorrenciasExpedicao = () => OCS;
  ${corta('const _expTurnoDaHora', '};\n')}
  ${corta('function _expHorasIdaDoDia', '\n}')}
  return { _expTurnoDaHora, _expHorasIdaDoDia };
`)(OCS);

ok(f._expTurnoDaHora('07:00') === 'manha', '07:00 é manhã');
ok(f._expTurnoDaHora('11:59') === 'manha', '11:59 ainda é manhã');
ok(f._expTurnoDaHora('12:00') === 'tarde', '12:00 já é tarde');
ok(f._expTurnoDaHora('') === 'manha', 'sem hora conta como manhã (como o painel)');
const ocupadas = f._expHorasIdaDoDia('2026-11-06');
ok(ocupadas.join(',') === '07:00,13:30', 'horas ocupadas no dia: ' + ocupadas.join(','));
ok(!ocupadas.includes('09:00'), 'expedição cancelada não ocupa horário');
ok(!ocupadas.includes('08:00'), 'expedição de outro dia não ocupa horário');

// O botão aparece no quadro e na linha da OE, e o salvar trata o tipo novo.
ok(/abrirModalExpExtra\('\$\{esc\(oc\.janela\.id\)\}'/.test(src), 'botão chama abrirModalExpExtra com a OE');
ok(src.includes("ctx.tipo === 'extra'"), 'salvarModalExpedicao grava a expedição extra');

console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
