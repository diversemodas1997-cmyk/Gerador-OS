/* Rode com:  node testes/avisos.js

   O MURAL DE AVISOS: o que os outros escreveram e carimbaram nas OS.

   O mural NÃO guarda evento nenhum — ele é lido das próprias OS (as notas da
   folha e o carimbo de status), e a marca de "já li" mora no navegador de cada
   um. Isso resolve o egresso (o blob desce inteiro a cada abertura), mas põe
   toda a responsabilidade em três contas que este teste protege:

     · o que entra na lista (janela de 30 dias, nota sem data fica fora);
     · o que conta como NÃO LIDO — o que chegou depois da última visita e não
       foi escrito por mim (ninguém precisa ser avisado de si mesmo);
     · a PRIMEIRA visita não pode nascer com trezentos avisos: um contador que
       nasce cheio ninguém zera, ignora.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

// localStorage de mentira: um objeto. E o que o mural usa para lembrar da
// ultima visita, entao o teste precisa de um por cenario.
const monta = (ctx) => new Function('ctx', `
  const localStorage = {
    getItem: (k) => (k in ctx.ls ? ctx.ls[k] : null),
    setItem: (k, v) => { ctx.ls[k] = String(v); }
  };
  const currentUser = ctx.login ? { email: ctx.login } : null;
  const STATE = ctx.STATE;
  ${constante('AVISOS_DIAS')}
  ${constante('AVISOS_MAX')}
  ${constante('STATUS_OS')}
  ${recorte('function _obsNotas', 'as notas da OS')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _statusOS', 'a leitura do status')}
  ${recorte('function _avisosEventos', 'a lista de eventos')}
  ${recorte('function _avisosChaveVistos', 'a chave da ultima visita')}
  ${recorte('function _avisosMarcarVistos', 'a marca da visita')}
  ${recorte('function _avisosVistos', 'a leitura da ultima visita')}
  ${recorte('function _avisosNaoLidos', 'a conta do que e novo')}
  return { _avisosEventos, _avisosNaoLidos, _avisosVistos, _avisosMarcarVistos };
`)(ctx);

const diasAtras = (d) => new Date(Date.now() - d * 24 * 60 * 60 * 1000).toISOString();
const ctxDe = (login, ordens, ls = {}) => {
  const ctx = { login, ls, STATE: { ordens } };
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const ordens = () => ([
  { id: 'a', os: '0483', obsNotas: [{ login: 'costura@diverse.local', texto: 'faltou pano na 3a fase', em: diasAtras(1) }] },
  { id: 'b', os: '0484', statusOS: 'parado', statusOSPor: 'enfesto.corte@diverse.local', statusOSEm: diasAtras(2) },
  { id: 'c', os: '0485', obsNotas: [{ login: 'admin@diverse.local', texto: 'recado velho', em: diasAtras(60) }] },
  { id: 'd', os: '0486', obsNotas: [{ login: 'costura@diverse.local', texto: 'migrada, sem data', anterior: true }] },
  { id: 'e', os: '0487' },                                    // nada aconteceu nela
  { id: 'f', os: '0488', statusOS: 'nao-iniciado' }           // status que e a ausencia de status
]);

console.log('-- o que entra no mural --');
let t = ctxDe('admin@diverse.local', ordens());
let ev = t.api._avisosEventos();
ok('1. entra a observacao e o carimbo de status', ev.length === 2, JSON.stringify(ev.map(e => e.os)));
ok('2. mais recente em cima', ev[0].os === '0483' && ev[1].os === '0484', JSON.stringify(ev.map(e => e.os)));
ok('3. a observacao traz OS, quem, quando e o texto',
   ev[0].tipo === 'obs' && ev[0].osId === 'a' && ev[0].quem === 'costura@diverse.local'
   && /faltou pano/.test(ev[0].texto), JSON.stringify(ev[0]));
ok('4. o status traz o estado carimbado',
   ev[1].tipo === 'status' && ev[1].status === 'parado'
   && ev[1].quem === 'enfesto.corte@diverse.local', JSON.stringify(ev[1]));
ok('5. recado de 60 dias fica fora da janela', !ev.some(e => e.os === '0485'), JSON.stringify(ev.map(e => e.os)));
ok('6. nota migrada, sem data, nao vira aviso (nao da para dizer se e nova)',
   !ev.some(e => e.os === '0486'), JSON.stringify(ev.map(e => e.os)));
ok('7. "nao iniciado" nao e evento — e a ausencia de carimbo',
   !ev.some(e => e.os === '0488'), JSON.stringify(ev.map(e => e.os)));

console.log('');
console.log('-- o que e NAO LIDO --');
t = ctxDe('admin@diverse.local', ordens());
ok('8. primeira visita nasce lida (ninguem zera contador que nasce cheio)',
   t.api._avisosNaoLidos().length === 0, String(t.api._avisosNaoLidos().length));
ok('9. e a marca ficou guardada para a proxima vez',
   typeof t.ctx.ls['avisosVistos:admin@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));

// Visita antiga guardada: tudo o que veio depois e novidade.
t = ctxDe('admin@diverse.local', ordens(), { 'avisosVistos:admin@diverse.local': diasAtras(5) });
ok('10. depois de uma visita antiga, os dois eventos sao novos',
   t.api._avisosNaoLidos().length === 2, String(t.api._avisosNaoLidos().length));

// O proprio recado nunca e aviso.
t = ctxDe('costura@diverse.local', ordens(), { 'avisosVistos:costura@diverse.local': diasAtras(5) });
const naoLidos = t.api._avisosNaoLidos();
ok('11. o que EU escrevi nao me avisa',
   naoLidos.length === 1 && naoLidos[0].tipo === 'status',
   JSON.stringify(naoLidos.map(e => e.os + ':' + e.tipo)));

// A marca e por login: dois turnos no mesmo computador nao herdam o "ja li".
t = ctxDe('enfesto.corte@diverse.local', ordens(), { 'avisosVistos:costura@diverse.local': diasAtras(5) });
ok('12. a marca de um login nao vale para o outro',
   t.api._avisosNaoLidos().length === 0
   && typeof t.ctx.ls['avisosVistos:enfesto.corte@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));

// Marcou como visto agora: nada pendente ate acontecer coisa nova.
t = ctxDe('admin@diverse.local', ordens(), { 'avisosVistos:admin@diverse.local': diasAtras(5) });
t.api._avisosMarcarVistos(new Date().toISOString());
ok('13. abrir o mural zera o que estava pendente',
   t.api._avisosNaoLidos().length === 0, String(t.api._avisosNaoLidos().length));
const comNovo = ordens();
comNovo[0].obsNotas[0].editadoEm = new Date(Date.now() + 1000).toISOString();
t.ctx.STATE.ordens = comNovo;
ok('14. e a observacao corrigida depois disso volta a avisar',
   t.api._avisosNaoLidos().length === 1, String(t.api._avisosNaoLidos().length));

console.log('');
console.log('-- sem login --');
t = ctxDe('', ordens(), {});
ok('15. quem nao entrou na conta ve o mural, sem nada pendente',
   t.api._avisosEventos().length === 2 && t.api._avisosNaoLidos().length === 0,
   String(t.api._avisosNaoLidos().length));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
