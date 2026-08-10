/* Rode com:  node testes/modo-servidor.js
   Verifica a escolha entre servidor da fabrica e nuvem, e — o mais importante —
   que a gravacao so e bloqueada no caso certo: servidor local CONFIGURADO mas
   fora do ar. Sem servidor local configurado, tudo grava como sempre gravou. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('const SUPA_URL');   // as constantes da nuvem entram no recorte
const fim = src.indexOf('/* ---- Tela de Configurações');
if (ini < 0 || fim < 0) { console.error('nao achei os limites do trecho'); process.exit(1); }

/* ------------------------------ stubs ------------------------------------ */
const guardado = {};
global.localStorage = {
  getItem: k => (k in guardado ? guardado[k] : null),
  setItem: (k, v) => { guardado[k] = String(v); },
  removeItem: k => { delete guardado[k]; }
};
let respostaFetch = { ok: false, erro: true };   // servidor local: como responde
let urlSondada = null;
global.fetch = async (url) => {
  urlSondada = url;
  if (respostaFetch.erro) throw new Error('ECONNREFUSED');
  return { ok: respostaFetch.ok };
};
global.AbortController = class { constructor() { this.signal = {}; } abort() {} };
global.document = { getElementById: () => null };
global.window = { supabase: { createClient: (u, k) => ({ _url: u, _key: k }) } };

eval(src.slice(ini, fim) + `
  globalThis.__modo = () => _modoServidor;
  globalThis.__supa = () => supa;
`);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};
const CFG = { url: 'http://192.168.0.50:8000', key: 'chave-local' };

(async () => {
  // 1. Nada configurado: nuvem, e GRAVA normalmente (comportamento de hoje).
  definirServidorLocal(null, null);
  await resolverServidor();
  ok('1. sem servidor local -> nuvem', __modo() === 'nuvem');
  ok('1b. sem servidor local a gravacao CONTINUA liberada', podeGravar() === true);
  ok('1c. cliente aponta para a nuvem', __supa()._url.includes('supabase.co'));

  // 2. Configurado e de pe: fala com a fabrica e grava.
  definirServidorLocal(CFG.url, CFG.key);
  respostaFetch = { ok: true, erro: false };
  await resolverServidor();
  ok('2. servidor local de pe -> modo local', __modo() === 'local');
  ok('2b. grava', podeGravar() === true);
  ok('2c. cliente aponta para a fabrica', __supa()._url === CFG.url);
  ok('2d. sondou o endereco certo', urlSondada === CFG.url + '/rest/v1/', urlSondada);

  // 3. Configurado mas fora do ar: cai para a nuvem em modo CONSULTA.
  respostaFetch = { ok: false, erro: true };
  await resolverServidor();
  ok('3. servidor local fora -> nuvem', __modo() === 'nuvem');
  ok('3b. gravacao BLOQUEADA (e o que impede os dois lados de divergir)',
     podeGravar() === false);
  ok('3c. cliente aponta para a nuvem', __supa()._url.includes('supabase.co'));

  // 4. De pe mas recusando a chave (401): nao serve como servidor.
  respostaFetch = { ok: false, erro: false };
  await resolverServidor();
  ok('4. servidor local recusando a chave -> nuvem, sem gravar',
     __modo() === 'nuvem' && podeGravar() === false);

  // 5. Barra sobrando no endereco nao pode virar // na sondagem.
  definirServidorLocal('http://192.168.0.50:8000///', CFG.key);
  respostaFetch = { ok: true, erro: false };
  await resolverServidor();
  ok('5. barra extra no endereco e normalizada', urlSondada === CFG.url + '/rest/v1/', urlSondada);

  // 6. Config corrompida no localStorage nao pode derrubar o programa.
  guardado['servidorLocal'] = '{lixo';
  await resolverServidor();
  ok('6. config corrompida -> nuvem, gravando (nao trava o app)',
     __modo() === 'nuvem' && podeGravar() === true);

  // 7. Config pela metade (sem chave) conta como nao configurada.
  guardado['servidorLocal'] = JSON.stringify({ url: CFG.url });
  await resolverServidor();
  ok('7. config sem a chave -> tratada como ausente', podeGravar() === true);

  console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
  process.exit(falhas ? 1 : 0);
})();
