/* Rode com:  node testes/prazo-gravacao.js

   O "SALVANDO..." QUE NUNCA TERMINAVA.

   O fetch do navegador não tem prazo. Uma requisição que sai e não volta —
   servidor no meio de um reinício, Wi-Fi que caiu com a conexão TCP de pé, Kong
   esperando um lock no Postgres, e o nginx da fábrica esperando 3600s por ele —
   fica pendurada para sempre, sem erro e sem resposta.

   O estrago não era a requisição perdida: era o que vinha depois. `_flushing`
   nunca era liberado, e cloudFlush devolve na hora quando há gravação no ar. A
   partir dali TODA gravação da sessão só se reagendava, o aviso ficava eterno em
   "☁ Salvando..." e o polling também emudecia (ele não relê enquanto há flush no
   ar). O programa parecia vivo, dizia estar salvando, e não salvava mais nada até
   alguém recarregar a página.

   Este teste guarda as duas defesas: TODA chamada tem prazo, e uma gravação dada
   por perdida não pode calar as próximas. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = [
  cortaLinha('const PRAZO_LEITURA_MS'),
  cortaLinha('const PRAZO_ESCRITA_MS'),
  cortaLinha('const FLUSH_ZUMBI_MS'),
  corta('function _prazoDaRequisicao'),
  corta('function _flushPodeComecar'),
  corta('function _fetchComPrazo')
].join('\n');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

// Ambiente: um fetch que registra o que recebeu e NUNCA responde (é o caso que
// este teste existe para cobrir), e um setTimeout que dispara na hora — assim o
// prazo é exercido sem o teste esperar 20 segundos de verdade.
function ambiente({ fetchQueNuncaVolta = true, prazoImediato = true } = {}) {
  const registro = { chamadas: [], abortos: [] };
  const fn = new Function('REG', 'OPC', `
    const AbortController = globalThis.AbortController;
    const DOMException = globalThis.DOMException;
    const setTimeout = OPC.prazoImediato ? (cb) => { cb(); return 0; } : globalThis.setTimeout;
    const clearTimeout = () => {};
    // Dublê fiel do fetch no que importa aqui: honra o signal. Cancelado, ele
    // REJEITA — é assim que o navegador se comporta, e é disso que depende a
    // requisição pendurada virar erro em vez de espera eterna.
    const motivo = (sig) => String((sig.reason && sig.reason.message) || sig.reason || 'abortado');
    const fetch = (entrada, init) => {
      REG.chamadas.push({ entrada, init });
      const sig = init && init.signal;
      if (sig && sig.aborted) {
        REG.abortos.push(motivo(sig));
        return Promise.reject(sig.reason || new Error('abortado'));
      }
      if (OPC.fetchQueNuncaVolta) {
        return new Promise((_, rej) => {
          if (sig) sig.addEventListener('abort', () => {
            REG.abortos.push(motivo(sig));
            rej(sig.reason || new Error('abortado'));
          });
        });   // sai e não volta, a não ser que alguém cancele
      }
      if (sig) sig.addEventListener('abort', () => REG.abortos.push(motivo(sig)));
      return Promise.resolve({ ok: true, status: 200 });
    };
    ${motor}
    return {
      _prazoDaRequisicao, _flushPodeComecar, _fetchComPrazo,
      PRAZO_LEITURA_MS, PRAZO_ESCRITA_MS, FLUSH_ZUMBI_MS
    };
  `);
  return { api: fn(registro, { fetchQueNuncaVolta, prazoImediato }), registro };
}

const { api } = ambiente({ fetchQueNuncaVolta: false, prazoImediato: false });

/* ---------- 1. cada tipo de chamada tem o seu prazo ---------- */

eq('leitura (GET) usa o prazo curto', api._prazoDaRequisicao('GET'), api.PRAZO_LEITURA_MS);
eq('HEAD também é leitura', api._prazoDaRequisicao('HEAD'), api.PRAZO_LEITURA_MS);
eq('sem método declarado vale como leitura', api._prazoDaRequisicao(undefined), api.PRAZO_LEITURA_MS);
eq('gravação (POST) usa o prazo longo', api._prazoDaRequisicao('post'), api.PRAZO_ESCRITA_MS);
eq('PATCH também é gravação', api._prazoDaRequisicao('PATCH'), api.PRAZO_ESCRITA_MS);
ok('a escrita tem folga sobre a leitura', api.PRAZO_ESCRITA_MS > api.PRAZO_LEITURA_MS,
  [api.PRAZO_LEITURA_MS, api.PRAZO_ESCRITA_MS]);
ok('nenhum prazo é infinito', api.PRAZO_LEITURA_MS > 0 && api.PRAZO_ESCRITA_MS > 0,
  [api.PRAZO_LEITURA_MS, api.PRAZO_ESCRITA_MS]);

/* ---------- 2. a requisição que não volta é cancelada ---------- */

(async () => {
  const { api: apiTravada, registro } = ambiente({ fetchQueNuncaVolta: true, prazoImediato: true });
  let rejeitou = false, msg = '';
  try {
    await apiTravada._fetchComPrazo('https://servidor/rest/v1/shared_data', { method: 'POST', body: 'x' });
  } catch (e) { rejeitou = true; msg = (e && e.message) || String(e); }
  ok('a chamada pendurada termina em erro, e não em espera eterna', rejeitou, msg);
  ok('o sinal de cancelamento foi disparado', registro.abortos.length === 1, registro.abortos);
  ok('o erro diz que o servidor não respondeu', /não respondeu/.test(registro.abortos[0] || ''), registro.abortos);
  ok('o fetch recebeu um signal', !!(registro.chamadas[0] && registro.chamadas[0].init.signal), registro.chamadas.length);
  // O corpo e o método não podem ser perdidos no caminho: é a gravação de verdade.
  eq('o método é repassado', registro.chamadas[0].init.method, 'POST');
  eq('o corpo é repassado', registro.chamadas[0].init.body, 'x');

  // Cancelar por fora (o supabase-js faz isso) tem de continuar valendo.
  const { api: api2, registro: reg2 } = ambiente({ fetchQueNuncaVolta: true, prazoImediato: false });
  const ctrlExterno = new AbortController();
  const p = api2._fetchComPrazo('https://servidor/rest/v1/x', { method: 'GET', signal: ctrlExterno.signal });
  ctrlExterno.abort(new Error('cancelado por quem chamou'));
  let msg2 = '';
  try { await p; } catch (e) { msg2 = (e && e.message) || String(e); }
  ok('cancelamento de fora chega ao fetch', reg2.abortos.length === 1, reg2.abortos);
  ok('e é o motivo de fora que aparece', /quem chamou/.test(reg2.abortos[0] || ''), reg2.abortos);

  /* ---------- 3. gravação dada por perdida não cala as próximas ---------- */

  const AGORA = 1000000;
  eq('sem gravação no ar, pode salvar', api._flushPodeComecar(false, 0, AGORA), true);
  eq('com gravação no ar há 1s, espera a vez',
    api._flushPodeComecar(true, AGORA - 1000, AGORA), false);
  eq('com gravação no ar há pouco menos do prazo, ainda espera',
    api._flushPodeComecar(true, AGORA - (api.FLUSH_ZUMBI_MS - 1), AGORA), false);
  eq('passado o prazo de zumbi, a próxima gravação sai',
    api._flushPodeComecar(true, AGORA - (api.FLUSH_ZUMBI_MS + 1), AGORA), true);
  // Sem carimbo de início não há como saber a idade: não travar o programa é o
  // que importa, então deixa passar.
  eq('gravação sem carimbo de início não bloqueia para sempre',
    api._flushPodeComecar(true, 0, AGORA), true);
  ok('o prazo de zumbi é maior que o de escrita (uma gravação sadia nunca é dada por perdida)',
    api.FLUSH_ZUMBI_MS > api.PRAZO_ESCRITA_MS, [api.FLUSH_ZUMBI_MS, api.PRAZO_ESCRITA_MS]);

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
