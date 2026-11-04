/* Rode com:  node testes/folha-conjugada-nova.js

   Folha agregada das OS conjugadas só para OS NOVAS (04/11/2026). As OS que já
   existiam continuam com a folha separada; a marca `folhaConjAgregada` é posta
   quando a OS nasce. O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + de); process.exit(1); }
  return src.slice(i, j);
}
const codigo = recorte("const FOLHA_CONJ_AGREGADA", "// Os desenhos do grupo");

let falhou = 0, n = 0;
function ok(cond, nome) { n++; console.log((cond ? '  ok  ' : '  FALHOU ') + n + '. ' + nome); if (!cond) falhou++; }

function montar(ordens) {
  const STATE = { ordens };
  const _statusOS = o => o.status || 'ativo';
  return new Function('STATE', '_statusOS', codigo + '\nreturn { _grupoConjugadoDaFolha, _grupoConjugadoDaOS };')(STATE, _statusOS);
}
const ids = g => g.membros.map(m => m.id).join(',');

// 1. Grupo antigo (nenhuma marca): cada uma na sua folha.
{
  const f = montar([{ id: 'a', os: '0600' }, { id: 'b', os: '0599', conjugadaPaiId: 'a' }]);
  ok(ids(f._grupoConjugadoDaFolha({ id: 'a' })) === 'a', 'matriz antiga: folha separada');
  ok(ids(f._grupoConjugadoDaFolha({ id: 'b' })) === 'a', 'conjugada antiga: nao entra em grupo agregado');
  ok(ids(f._grupoConjugadoDaOS({ id: 'a' })) === 'a,b', 'a amarra dos dados segue intacta');
}
// 2. Grupo novo: agrega.
{
  const f = montar([{ id: 'a', os: '0700', folhaConjAgregada: true },
                    { id: 'b', os: '0699', conjugadaPaiId: 'a', folhaConjAgregada: true }]);
  ok(ids(f._grupoConjugadoDaFolha({ id: 'a' })) === 'a,b', 'matriz nova: agrega o grupo');
}
// 3. Matriz antiga com conjugada nova: a matriz manda, fica separado.
{
  const f = montar([{ id: 'a', os: '0600' },
                    { id: 'b', os: '0701', conjugadaStatusPaiId: 'a', folhaConjAgregada: true }]);
  ok(ids(f._grupoConjugadoDaFolha({ id: 'a' })) === 'a', 'matriz antiga + conjugada nova: separado');
}
// 4. Matriz nova com uma antiga conjugada à mão: a antiga fica fora.
{
  const f = montar([{ id: 'a', os: '0700', folhaConjAgregada: true },
                    { id: 'b', os: '0699', conjugadaPaiId: 'a', folhaConjAgregada: true },
                    { id: 'c', os: '0600', conjugadaStatusPaiId: 'a' }]);
  ok(ids(f._grupoConjugadoDaFolha({ id: 'a' })) === 'a,b', 'matriz nova: a OS antiga conjugada a mao fica fora');
}
// 5. A OS nasce marcada: _mesclarComOSExistente sem anterior, e na aprovação.
ok(/if \(!ant\) \{\s*data\[FOLHA_CONJ_AGREGADA\] = true;/.test(src), 'OS nova pelo formulario nasce marcada');
ok(/os\.criadoPor = _obsQuemSou\(\);\s*os\[FOLHA_CONJ_AGREGADA\] = true;/.test(src), 'OS aprovada no planejamento nasce marcada');
ok(/delete copia\[FOLHA_CONJ_AGREGADA\]/.test(src), 'copia de OS nao leva a marca da original');

console.log(falhou ? `\n${falhou} falha(s)` : '\ntudo certo');
process.exit(falhou ? 1 : 0);
