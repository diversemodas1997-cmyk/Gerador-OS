/* Rode com:  node testes/folha-oe-do-dia.js

   FOLHA DE OE DO DIA INTEIRO (05/11/2026, Junior: "as OE do mesmo dia, mas em
   períodos diferentes, aparecem na mesma folha de impressão").

   Aberta pelo quadro de uma OE, a folha traz todas as OE da mesma data efetiva:
     · a manhã e a tarde do dia entram juntas, a do dia seguinte não;
     · OE remarcada PARA aquele dia entra; a remarcada para outro dia sai;
     · os números saem em ordem e sem repetir (cabeçalho e nome do PDF).

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
const montar = (STATE, chave) => new Function('STATE', `
  let expFolhaOe = ${JSON.stringify(chave)};
  ${corta('const _expChaveOe', ';\n')}
  ${corta('function _oeNumeroDe', '\n}')}
  ${corta('function _expFolhaDia', '\n}')}
  ${corta('const _expNaFolha', ';\n')}
  ${corta('const _expNumerosDaFolha', ';\n')}
  return { _expNaFolha, _expNumerosDaFolha, _expFolhaDia };
`)(STATE);

let falhas = 0;
const ok = (cond, msg) => { if (cond) console.log('ok   ' + msg); else { falhas++; console.log('FALHA ' + msg); } };
const oc = (janelaId, dataOrig, data) => ({ janela: { id: janelaId }, dataOrig, data: data || dataOrig });

const STATE = {
  expedicaoCargas: [
    { id: 'a', janelaId: 'M', data: '2026-11-04', oe: 85 },
    { id: 'b', janelaId: 'T', data: '2026-11-04', oe: 86 },
    { id: 'c', janelaId: 'M', data: '2026-11-05', oe: 87 },
    { id: 'd', janelaId: 'T', data: '2026-11-03', oe: 84 }
  ],
  expedicaoExcecoes: [{ janelaId: 'T', data: '2026-11-03', tipo: 'remarcada', novaData: '2026-11-04' }]
};
const manha = oc('M', '2026-11-04'), tarde = oc('T', '2026-11-04'), amanha = oc('M', '2026-11-05');
const remarcada = oc('T', '2026-11-03', '2026-11-04');

const f = montar(STATE, 'M|2026-11-04');
ok(f._expNaFolha(manha) && f._expNaFolha(tarde), 'clicando na manhã: manhã e tarde do dia na folha');
ok(!f._expNaFolha(amanha), 'o dia seguinte fica fora');
ok(f._expNaFolha(remarcada), 'OE remarcada para o dia entra');
const nums = f._expNumerosDaFolha([tarde, manha, remarcada, manha]);
ok(JSON.stringify(nums) === '[84,85,86]', 'números em ordem e sem repetir (deu ' + JSON.stringify(nums) + ')');

const g = montar(STATE, 'T|2026-11-04');
ok(g._expNaFolha(manha) && g._expNaFolha(tarde), 'clicando na tarde: também o dia inteiro');

const h = montar(STATE, 'T|2026-11-03');
ok(h._expFolhaDia() === '2026-11-04', 'clicando na remarcada: o dia é o da remarcação');

const sem = montar(STATE, null);
ok(sem._expNaFolha(amanha) && sem._expFolhaDia() === null, 'sem OE escolhida: folha do período, como antes');

if (falhas) { console.log('\n' + falhas + ' falha(s)'); process.exit(1); }
console.log('\ntudo certo');
