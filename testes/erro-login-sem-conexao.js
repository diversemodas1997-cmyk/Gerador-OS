/* Rode com:  node testes/erro-login-sem-conexao.js
   Quando o login nem chega ao servidor ("Failed to fetch"), a tela tem de dizer
   que NAO e a senha e para onde ir — e a orientacao muda conforme o lado que
   falhou (servidor da fabrica ou nuvem). Ver traduzirErroAuth em app.js. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('function traduzirErroAuth(msg)');
const fim = src.indexOf('async function submeterAuth()');
if (ini < 0 || fim < 0) { console.error('nao achei os limites do trecho'); process.exit(1); }

let _modoServidor;
const MODO_LOCAL = 'local', MODO_NUVEM = 'nuvem';
eval(src.slice(ini, fim) + `
  globalThis.__traduzir = traduzirErroAuth;
  globalThis.__modo = m => { _modoServidor = m; };
`);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

__modo(MODO_NUVEM);
const nuvem = __traduzir('Failed to fetch');
ok('1. nuvem: nao mostra o texto cru', nuvem !== 'Failed to fetch', nuvem);
ok('1b. nuvem: diz que nao e a senha', /não é a sua senha/i.test(nuvem), nuvem);
ok('1c. nuvem: manda abrir pelo atalho da fabrica', /Gerador-OS/.test(nuvem), nuvem);
ok('1d. Firefox (NetworkError) tem a mesma traducao',
   __traduzir('NetworkError when attempting to fetch resource.') === nuvem);
ok('1e. Safari (Load failed) tem a mesma traducao', __traduzir('Load failed') === nuvem);
ok('1f. prazo estourado tem a mesma traducao',
   __traduzir('O servidor não respondeu em 15s') === nuvem);

__modo(MODO_LOCAL);
const local = __traduzir('Failed to fetch');
ok('2. local: fala do servidor da fabrica', /servidor da fábrica não respondeu/i.test(local), local);
ok('2b. local: diz que nao e a senha', /não é a sua senha/i.test(local), local);
ok('2c. local: nao manda trocar de atalho', !/atalho/.test(local), local);

ok('3. senha errada continua sendo senha errada',
   __traduzir('Invalid login credentials') === 'Nome ou senha incorretos');
ok('3b. erro desconhecido passa como veio', __traduzir('algo novo') === 'algo novo');

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
