/* Rode com:  node testes/oe-numero-sem-buraco.js

   OE SEM BURACO NA SEQUÊNCIA (05/11/2026 — a OE 0086 foi pulada: nasceu em
   04/11 com uma OS que depois saiu da carga, e a próxima já saiu 0087).

   O que o teste guarda:
     · número que ficou sem carga no TOPO é dado de novo à próxima OE;
     · contador adiantado (gasto à toa) não pula número;
     · expedição que já tem número continua com ele;
     · carga movida que era a única da OE leva o número junto.

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(nome) {
  const i = src.indexOf(nome);
  const j = src.indexOf('\n}', i);
  if (i < 0 || j < 0) { console.error('nao achei ' + nome + ' no app.js'); process.exit(1); }
  return src.slice(i, j) + '\n}';
}
const montar = (STATE) => new Function('STATE', `
  ${corta('function _oeNumeroDe')}
  ${corta('function _oeNumeroGarantir')}
  ${corta('function _oeCargasComNumero')}
  return { _oeNumeroDe, _oeNumeroGarantir, _oeCargasComNumero };
`)(STATE);

let falhas = 0;
const ok = (cond, msg) => { if (cond) console.log('ok   ' + msg); else { falhas++; console.log('FALHA ' + msg); } };
const carga = (id, janelaId, data, oe) => ({ id, janelaId, data, oe });

// O caso de 04/11: 85 existe, o contador ficou em 86 (a OS saiu), nenhuma carga com 86.
{
  const S = { expedicaoCargas: [carga('a', 'M', '2026-11-04', 85)], meta: { oeCounter: 86 } };
  const f = montar(S);
  const r = f._oeNumeroGarantir('T', '2026-11-04');
  ok(r.n === 86, 'número órfão no topo volta: próxima OE = 86 (deu ' + r.n + ')');
  ok(S.meta.oeCounter === 86, 'contador acompanha o número dado');
}
// Expedição que já tem número fica com ele.
{
  const S = { expedicaoCargas: [carga('a', 'M', '2026-11-04', 85)], meta: { oeCounter: 85 } };
  const r = montar(S)._oeNumeroGarantir('M', '2026-11-04');
  ok(r.n === 85 && !r.novo, 'mesma expedição continua 85');
}
// Sequência normal.
{
  const S = { expedicaoCargas: [carga('a', 'M', '2026-11-04', 85)], meta: { oeCounter: 85 } };
  const r = montar(S)._oeNumeroGarantir('T', '2026-11-04');
  ok(r.n === 86 && r.novo, 'nova expedição = 86 e grava o contador');
}
// Carga movida, única da OE 86, para uma expedição sem número: leva o 86.
{
  const S = { expedicaoCargas: [carga('a', 'M', '2026-11-04', 85), carga('b', 'T', '2026-11-04', 86)], meta: { oeCounter: 86 } };
  const f = montar(S);
  ok(f._oeCargasComNumero(86, 'b') === 0, 'sem a carga movida, ninguém mais tem o 86');
  const r = f._oeNumeroGarantir('T', '2026-11-05', 86);
  ok(r.n === 86, 'carga movida leva o próprio número (deu ' + r.n + ')');
}
// Movida para uma expedição que já tem número: entra na de lá.
{
  const S = { expedicaoCargas: [carga('a', 'M', '2026-11-05', 87), carga('b', 'T', '2026-11-04', 86)], meta: { oeCounter: 87 } };
  const r = montar(S)._oeNumeroGarantir('M', '2026-11-05', 86);
  ok(r.n === 87, 'destino com número manda (87)');
}

if (falhas) { console.log('\n' + falhas + ' falha(s)'); process.exit(1); }
console.log('\ntudo certo');
