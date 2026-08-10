/* Rode com:  node testes/leitura-parcial.js
   Exercita a camada de leitura do app contra um Supabase falso.
   Recorta do app.js real o trecho que vai de _adotarServidorPreservandoEdicoes
   ate _cloudLoadCompleto, e roda com stubs no lugar dos globais do navegador. */
const fs = require('fs');
const path = require('path');

const APP = path.join(__dirname, '..', 'app.js');
const src = fs.readFileSync(APP, 'utf8');
const ini = src.indexOf('// `chavesDoServidor` e a lista AUTORITATIVA'.replace('e a', 'é a'));
const fim = src.indexOf('async function cloudFlush()');
if (ini < 0 || fim < 0) { console.error('nao achei os limites do trecho'); process.exit(1); }
const trecho = src.slice(ini, fim);

/* ---- Supabase falso: simula a projecao do PostgREST, inclusive data->chave --- */
function criarSupa(estado) {
  estado.lidos = [];                       // bytes "baixados" por consulta
  const projetar = (tabela, sel) => {
    if (tabela === 'sync_signal') {
      return { data: estado.semMapa ? { key_versions: null }
                 : { key_versions: JSON.parse(JSON.stringify(estado.versoes)) }, error: null };
    }
    if (estado.erroLeitura) return { data: null, error: { code: 'PGRST301' } };
    const out = {};
    for (const campo of sel.split(',')) {
      const m = campo.match(/^([A-Za-z_][A-Za-z0-9_]*):data->([A-Za-z_][A-Za-z0-9_]*)$/);
      if (m) {
        if (estado.selParcialQuebrado) continue;   // simula PostgREST que ignora ->
        out[m[1]] = estado.blob[m[2]];
      } else if (campo === 'updated_at') out.updated_at = estado.updated_at;
      else if (campo === 'data') out.data = estado.blob;
    }
    estado.lidos.push(JSON.stringify(out).length);
    return { data: out, error: null };
  };
  return {
    from(tabela) {
      const q = {
        _sel: '',
        select(s) { q._sel = s; return q; },
        eq() { return q; },
        maybeSingle() { return Promise.resolve(projetar(tabela, q._sel)); },
        then(res, rej) { return Promise.resolve(projetar(tabela, q._sel)).then(res, rej); }
      };
      return q;
    },
    auth: { refreshSession: async () => ({ data: {} }) }
  };
}

/* ---------------------------- stubs dos globais --------------------------- */
let supa = null, currentUser = { id: 'u1' }, cloudCache = null;
let _dirtyKeys = new Set(), _baseline = {}, _cloudLoadErro = false;
let _ultimoUpdatedAtServidor = null;
const DEVICE_ID = 'ESTA-ABA';
const toast = () => {};
global.setTimeout = global.setTimeout;

// `let` declarado dentro de eval fica preso ao escopo do eval — de fora nao da
// para ler nem zerar _versoesConhecidas. Estes acessores expoem a variavel REAL
// que as funcoes do app usam; sem eles o teste mediria outra coisa.
eval(trecho + `
  globalThis.__getVersoes = () => _versoesConhecidas;
  globalThis.__setVersoes = v => { _versoesConhecidas = v; };
`);

/* ------------------------------- cenarios --------------------------------- */
// Proporcoes parecidas com as reais: `operacoes` e `desenhos` sao o peso morto
// que a leitura parcial precisa deixar de baixar quando so `ordens` muda.
const GORDO = (etiqueta, n) => JSON.stringify([{ id: 1, t: etiqueta.repeat(n) }]);
const BLOB = () => ({
  ordens:    '[{"id":1,"n":"OS-1"}]',
  operacoes: GORDO('op', 1500),
  desenhos:  GORDO('dz', 800),
  cores:     '["preto"]',
  _device:   'OUTRA-MAQUINA'
});
const V1 = { ordens: 't1', operacoes: 't1', desenhos: 't1', cores: 't1' };

function cenario(over) {
  const est = Object.assign({
    blob: BLOB(), versoes: Object.assign({}, V1), updated_at: 't1',
    semMapa: false, selParcialQuebrado: false, erroLeitura: false
  }, over || {});
  supa = criarSupa(est);
  cloudCache = null; _baseline = {}; _dirtyKeys = new Set();
  __setVersoes(null); _cloudLoadErro = false;
  return est;
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

(async () => {
  // 1. Primeira leitura: sem base de comparacao, tem que baixar tudo.
  let e = cenario();
  await cloudLoad();
  ok('1. primeira leitura baixa o blob inteiro', cloudCache.ordens === BLOB().ordens
     && cloudCache.operacoes === BLOB().operacoes && cloudCache.cores === BLOB().cores);
  ok('1b. guarda o mapa de versoes como base', __getVersoes()
     && __getVersoes().ordens === 't1');

  // 2. So `ordens` mudou: a parcial tem que baixar so `ordens`.
  e.blob.ordens = '[{"id":1,"n":"OS-1"},{"id":2,"n":"OS-2"}]';
  e.versoes.ordens = 't2'; e.updated_at = 't2'; e.lidos = [];
  const antesOperacoes = cloudCache.operacoes;
  await cloudLoad();
  ok('2. baixou a ordens nova', cloudCache.ordens === e.blob.ordens);
  ok('2b. operacoes NAO desceu de novo', cloudCache.operacoes === antesOperacoes);
  const V_ORDENS = e.blob.ordens.length, V_TUDO = JSON.stringify(BLOB()).length;
  ok('2c. baixou so o tamanho de ordens, nao o blob inteiro',
     e.lidos.length === 1 && e.lidos[0] < V_ORDENS + 200 && e.lidos[0] < V_TUDO / 10,
     'baixou ' + e.lidos[0] + 'B; ordens=' + V_ORDENS + 'B; blob=' + V_TUDO + 'B');
  console.log('       -> ' + e.lidos[0] + ' bytes em vez de ' + V_TUDO
     + ' (' + (V_TUDO / e.lidos[0]).toFixed(1) + 'x menos)');

  // 3. Chave suja localmente nao pode ser sobrescrita pela parcial.
  e = cenario();
  await cloudLoad();
  cloudCache.cores = '["local-nao-salvo"]';
  _dirtyKeys.add('cores');
  e.blob.cores = '["do-servidor"]'; e.versoes.cores = 't2'; e.updated_at = 't2';
  await cloudLoad();
  ok('3. edicao local pendente sobrevive a leitura parcial',
     cloudCache.cores === '["local-nao-salvo"]', 'ficou ' + cloudCache.cores);

  // 4. Chave apagada no servidor sai do cache local, mesmo sem descer nada.
  e = cenario();
  await cloudLoad();
  delete e.blob.cores; delete e.versoes.cores; e.updated_at = 't2';
  await cloudLoad();
  ok('4. chave removida no servidor some do cache', !('cores' in cloudCache));
  ok('4b. as outras continuam', cloudCache.ordens === BLOB().ordens);

  // 5. `_device` precisa vir na parcial, senao o polling confunde a gravacao
  //    de outra maquina com a nossa e a tela nao atualiza.
  e = cenario();
  await cloudLoad();
  cloudCache._device = DEVICE_ID;                 // como fica apos o nosso flush
  e.blob._device = 'MAQUINA-DO-CORTE';
  e.blob.ordens = '[{"id":3}]'; e.versoes.ordens = 't2'; e.updated_at = 't2';
  await cloudLoad();
  ok('5. _device vem junto na parcial', cloudCache._device === 'MAQUINA-DO-CORTE',
     'ficou ' + cloudCache._device);

  // 6. Servidor sem o mapa (coluna ainda nao criada): cai na leitura completa.
  e = cenario({ semMapa: true });
  await cloudLoad();
  ok('6. sem mapa no servidor, baixa tudo e funciona',
     cloudCache.ordens === BLOB().ordens && cloudCache.cores === BLOB().cores);

  // 7. Mapa igual mas updated_at andou (aba com versao antiga do app gravou):
  //    nao da pra saber o que mudou -> tem que baixar tudo.
  e = cenario();
  await cloudLoad();
  e.blob.ordens = '[{"id":99}]'; e.updated_at = 't2';   // mapa NAO mexeu
  await cloudLoad();
  ok('7. gravacao de app antigo ainda e capturada', cloudCache.ordens === '[{"id":99}]',
     'ficou ' + cloudCache.ordens);

  // 8. O select parcial nao devolve o que se esperava -> cai para a completa.
  e = cenario();
  await cloudLoad();
  e.blob.ordens = '[{"id":7}]'; e.versoes.ordens = 't2'; e.updated_at = 't2';
  e.selParcialQuebrado = true;
  await cloudLoad();
  ok('8. select parcial furado cai para a leitura completa',
     cloudCache.ordens === '[{"id":7}]' && cloudCache.cores === BLOB().cores,
     'ficou ' + cloudCache.ordens);

  // 9. Nome de chave fora do padrao derruba para a completa (nada interpolado).
  e = cenario();
  await cloudLoad();
  e.versoes['chave com espaco'] = 't2'; e.blob['chave com espaco'] = 'x';
  e.updated_at = 't2';
  ok('9. nome estranho manda baixar tudo', _chavesParaBaixar(e.versoes) === null);

  console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
  process.exit(falhas ? 1 : 0);
})();
