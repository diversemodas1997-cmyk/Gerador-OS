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
// O status importa tanto quanto o ok: a raiz /rest/v1/ responde 403 a chave
// anonima com o servidor inteiro funcionando, e so o 401 significa chave
// recusada. Um falso que devolvesse so "ok" nao distinguiria os dois.
let respostaFetch = { ok: false, erro: true, status: 0 };
let urlSondada = null;
// O que o servidor devolve em /servidor-local.json. null = nao existe (404),
// que e o caso da nuvem e de qualquer servidor mais velho.
let servidorLocalJson = null;
global.fetch = async (url) => {
  if (String(url).endsWith('/servidor-local.json')) {
    return servidorLocalJson
      ? { ok: true, status: 200, json: async () => servidorLocalJson }
      : { ok: false, status: 404, json: async () => ({}) };
  }
  urlSondada = url;
  if (respostaFetch.erro) throw new Error('ECONNREFUSED');
  return { ok: respostaFetch.ok, status: respostaFetch.status };
};
global.AbortController = class { constructor() { this.signal = {}; } abort() {} };
global.document = { getElementById: () => null };
global.window = { supabase: { createClient: (u, k) => ({ _url: u, _key: k }) } };
global.location = { origin: 'https://193.168.0.8', protocol: 'https:' };

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
  respostaFetch = { ok: true, erro: false, status: 200 };
  await resolverServidor();
  ok('2. servidor local de pe -> modo local', __modo() === 'local');
  ok('2b. grava', podeGravar() === true);
  ok('2c. cliente aponta para a fabrica', __supa()._url === CFG.url);
  ok('2d. sondou o endereco certo', urlSondada === CFG.url + '/rest/v1/', urlSondada);

  // 3. Configurado mas fora do ar: cai para a nuvem em modo CONSULTA.
  respostaFetch = { ok: false, erro: true, status: 0 };
  await resolverServidor();
  ok('3. servidor local fora -> nuvem', __modo() === 'nuvem');
  ok('3b. gravacao BLOQUEADA (e o que impede os dois lados de divergir)',
     podeGravar() === false);
  ok('3c. cliente aponta para a nuvem', __supa()._url.includes('supabase.co'));

  // 4. De pe mas recusando a chave (401): nao serve como servidor.
  respostaFetch = { ok: false, erro: false, status: 401 };
  await resolverServidor();
  ok('4. servidor local recusando a chave (401) -> nuvem, sem gravar',
     __modo() === 'nuvem' && podeGravar() === false);

  // 4b. 403 na raiz do REST e o comportamento NORMAL do Supabase atual: essa
  // rota virou so de admin, e a chave anonima e recusada ali com o servidor
  // inteiro funcionando. Enquanto isto contava como "fora do ar", TODA maquina
  // da fabrica caia na nuvem em modo consulta - o servidor local nunca era
  // usado, por mais bem configurado que estivesse.
  respostaFetch = { ok: false, erro: false, status: 403 };
  await resolverServidor();
  ok('4b. 403 na raiz do REST -> segue sendo o servidor da fabrica, gravando',
     __modo() === 'local' && podeGravar() === true);

  // 5. Barra sobrando no endereco nao pode virar // na sondagem.
  definirServidorLocal('http://192.168.0.50:8000///', CFG.key);
  respostaFetch = { ok: true, erro: false, status: 200 };
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

  // ---- Descoberta automatica pelo servidor que serviu a pagina ----------
  // Sem isto existe um no sem saida: para configurar o servidor e preciso
  // chegar em Configuracoes, que exige estar logado, e o login vai para a
  // nuvem - que pode estar fora do ar, restrita ou sem a conta da pessoa.
  // Abrindo o programa PELO servidor da fabrica, ele se apresenta.
  definirServidorLocal(null, null);
  servidorLocalJson = { url: 'https://193.168.0.8', key: 'chave-publicada' };
  respostaFetch = { ok: true, erro: false, status: 200 };
  await resolverServidor();
  ok('11. maquina virgem que abre pelo servidor se conecta sozinha',
     __modo() === 'local' && __supa()._url === 'https://193.168.0.8');
  ok('11b. usa a chave que o servidor publicou', __supa()._key === 'chave-publicada');
  ok('11c. guarda a descoberta para as proximas aberturas',
     (JSON.parse(guardado['servidorLocal'] || '{}')).key === 'chave-publicada');

  // Servido pela nuvem (ou por um servidor mais velho): o arquivo nao existe,
  // e tudo tem de seguir como sempre foi.
  definirServidorLocal(null, null);
  servidorLocalJson = null;
  await resolverServidor();
  ok('12. sem o arquivo publicado -> nuvem, gravando como antes',
     __modo() === 'nuvem' && podeGravar() === true);

  // Publicado, mas o servidor recusa a chave: nao adianta ter se apresentado.
  definirServidorLocal(null, null);
  servidorLocalJson = { url: 'https://193.168.0.8', key: 'chave-errada' };
  respostaFetch = { ok: false, erro: false, status: 401 };
  await resolverServidor();
  ok('13. chave publicada que o servidor recusa -> nuvem, e NAO fica guardada',
     __modo() === 'nuvem' && !guardado['servidorLocal']);

  // O IP do servidor mudou. A maquina tem o endereco VELHO guardado, e ele nao
  // responde mais. Ela precisa aceitar o endereco novo de quem serviu a pagina,
  // senao a troca de IP obrigaria a passar computador por computador - e ate la
  // todos estariam na nuvem.
  definirServidorLocal('https://193.168.0.8', 'chave-publicada');
  servidorLocalJson = { url: 'https://193.168.0.200', key: 'chave-publicada' };
  let tentativas = 0;
  global.fetch = async (url) => {
    if (String(url).endsWith('/servidor-local.json')) {
      return { ok: true, status: 200, json: async () => servidorLocalJson };
    }
    urlSondada = url;
    tentativas++;
    // O endereco velho nao atende; o novo atende.
    if (String(url).startsWith('https://193.168.0.200')) return { ok: true, status: 200 };
    throw new Error('ECONNREFUSED');
  };
  await resolverServidor();
  ok('14. endereco velho morto -> adota o endereco novo que o servidor anuncia',
     __modo() === 'local' && __supa()._url === 'https://193.168.0.200');
  ok('14b. grava o endereco novo, para nao redescobrir toda vez',
     (JSON.parse(guardado['servidorLocal'] || '{}')).url === 'https://193.168.0.200');
  ok('14c. tentou o velho antes de trocar (nao troca por capricho)', tentativas >= 2);

  // Volta o fetch normal para o resto do arquivo.
  global.fetch = async (url) => {
    if (String(url).endsWith('/servidor-local.json')) {
      return servidorLocalJson
        ? { ok: true, status: 200, json: async () => servidorLocalJson }
        : { ok: false, status: 404, json: async () => ({}) };
    }
    urlSondada = url;
    if (respostaFetch.erro) throw new Error('ECONNREFUSED');
    return { ok: respostaFetch.ok, status: respostaFetch.status };
  };

  // Volta ao estado neutro para o resto do arquivo.
  servidorLocalJson = null;
  definirServidorLocal(null, null);

  // ---- Endereco das imagens dos desenhos -------------------------------
  // O dado guarda so o NOME do arquivo; o endereco e montado contra o servidor
  // em uso. E o que faz o mesmo desenho aparecer na fabrica e fora dela.
  const NOME = '1777290341246_1bsz6o.png';
  const NA_NUVEM = 'https://ckkqrjkhorvaahyazqsr.supabase.co/storage/v1/object/public/desenhos/' + NOME;

  definirServidorLocal(null, null);
  respostaFetch = { ok: false, erro: true };
  await resolverServidor();
  ok('8. nome vira endereco da NUVEM quando e a nuvem que atende',
     urlDesenho(NOME) === NA_NUVEM, urlDesenho(NOME));

  definirServidorLocal(CFG.url, CFG.key);
  respostaFetch = { ok: true, erro: false };
  await resolverServidor();
  ok('8b. o MESMO nome vira endereco da FABRICA quando ela atende',
     urlDesenho(NOME) === CFG.url + '/storage/v1/object/public/desenhos/' + NOME,
     urlDesenho(NOME));

  ok('9. endereco completo antigo passa intacto (dado velho nao some)',
     urlDesenho(NA_NUVEM) === NA_NUVEM);
  ok('9b. imagem em base64 passa intacta',
     urlDesenho('data:image/png;base64,iVBOR') === 'data:image/png;base64,iVBOR');
  ok('9c. vazio nao vira endereco quebrado',
     urlDesenho('') === '' && urlDesenho(null) === '' && urlDesenho(undefined) === '');
  ok('9d. nome com espaco e escapado', urlDesenho('a b.png').endsWith('a%20b.png'));

  // A volta: usada pela migracao que encurta os enderecos ja gravados.
  ok('10. extrai o nome de um endereco da nuvem', _nomeDoArquivoDesenho(NA_NUVEM) === NOME);
  ok('10b. extrai o nome de um endereco da fabrica',
     _nomeDoArquivoDesenho(CFG.url + '/storage/v1/object/public/desenhos/' + NOME) === NOME);
  ok('10c. nome ja encurtado NAO e mexido de novo', _nomeDoArquivoDesenho(NOME) === null);
  ok('10d. base64 nao e mexido', _nomeDoArquivoDesenho('data:image/png;base64,iVBOR') === null);
  ok('10e. endereco de outro bucket nao e mexido',
     _nomeDoArquivoDesenho('https://x.co/storage/v1/object/public/outro/a.png') === null);
  ok('10f. ida e volta fecha', _nomeDoArquivoDesenho(urlDesenho(NOME)) === NOME);

  console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
  process.exit(falhas ? 1 : 0);
})();
