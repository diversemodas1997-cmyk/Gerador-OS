/* Rode com:  node testes/espelho.js
   O espelho ESCREVE na nuvem. Um defeito aqui não deixa a tela estranha: apaga
   a cópia de fora do prédio. Estes testes cobrem principalmente o que ele tem
   que se RECUSAR a fazer. */
const { espelhar, blobVazio } = require('../servidor/espelho');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

const LOCAL = 'http://localhost:8000', NUVEM = 'https://nuvem.supabase.co';
const CHEIO = { ordens: '[{"id":1}]', desenhos: '[{"id":9,"img":"a.png"}]', cores: '["preto"]' };
const VAZIO = { ordens: '[]', desenhos: '[]' };

/* Servidor de mentira: guarda o estado dos dois lados e anota o que foi escrito. */
function montar(estado) {
  const escritas = [];
  const resposta = (corpo, ok_) => Promise.resolve({
    ok: ok_ !== false, status: ok_ === false ? 500 : 200,
    text: async () => (corpo === undefined ? '' : JSON.stringify(corpo)),
    arrayBuffer: async () => Buffer.from('imagem-binaria'),
    headers: { get: () => 'image/png' }
  });
  const buscar = async (url, op) => {
    const metodo = (op && op.method) || 'GET';
    const lado = url.startsWith(NUVEM) ? 'nuvem' : 'fabrica';
    if (estado[lado + 'Fora']) throw new Error('ECONNREFUSED');

    if (url.includes('/rest/v1/shared_data')) {
      if (metodo === 'GET') {
        const d = estado[lado].dados;
        return resposta(d ? [d] : []);
      }
      escritas.push({ alvo: lado, tipo: 'dados', corpo: JSON.parse(op.body) });
      estado[lado].dados = JSON.parse(op.body);
      return resposta();
    }
    if (url.includes('/rest/v1/sync_signal')) {
      if (metodo === 'GET') return resposta(estado[lado].sinal ? [estado[lado].sinal] : []);
      escritas.push({ alvo: lado, tipo: 'sinal', corpo: JSON.parse(op.body) });
      return resposta();
    }
    if (url.includes('/rest/v1/mensagens')) {
      if (estado[lado + 'SemMensagens']) return resposta({ message: 'relation does not exist' }, false);
      const lista = estado[lado].mensagens || [];
      if (metodo === 'GET') {
        // A nuvem so pede os ids; a fabrica manda a mensagem inteira.
        return resposta(url.includes('select=id') ? lista.map(m => ({ id: m.id })) : lista);
      }
      if (metodo === 'DELETE') {
        const ids = (url.match(/id=in\.\(([^)]*)\)/) || [, ''])[1].split(',').filter(Boolean);
        escritas.push({ alvo: lado, tipo: 'msg-apagar', ids });
        estado[lado].mensagens = lista.filter(m => !ids.includes(m.id));
        return resposta();
      }
      const novas = JSON.parse(op.body);
      escritas.push({ alvo: lado, tipo: 'msg-enviar', ids: novas.map(m => m.id) });
      estado[lado].mensagens = lista.concat(novas);
      return resposta();
    }
    if (url.includes('/storage/v1/object/list/')) {
      const offset = JSON.parse(op.body).offset;
      return resposta(offset ? [] : (estado[lado].imagens || []).map(n => ({ name: n })));
    }
    if (url.includes('/storage/v1/object/public/')) return resposta();
    if (url.includes('/storage/v1/object/')) {
      const nome = decodeURIComponent(url.split('/').pop());
      escritas.push({ alvo: lado, tipo: 'imagem', nome });
      estado[lado].imagens = (estado[lado].imagens || []).concat([nome]);
      return resposta();
    }
    throw new Error('rota não prevista no teste: ' + url);
  };
  return { buscar, escritas };
}

const cenario = (fabrica, nuvem, extra) => Object.assign({
  fabrica: Object.assign({ dados: null, sinal: null, imagens: [], mensagens: [] }, fabrica),
  nuvem: Object.assign({ dados: null, sinal: null, imagens: [], mensagens: [] }, nuvem)
}, extra || {});
const msg = (id, texto) => ({ id, criado_em: '2026-08-26T12:0' + id + ':00Z',
                              autor_id: 'u1', autor: 'costura@diverse.local', texto });

const rodar = (est, op) => {
  const { buscar, escritas } = montar(est);
  return espelhar(Object.assign({
    local: LOCAL, localKey: 'kl', nuvem: NUVEM, nuvemKey: 'kn', buscar
  }, op || {})).then(rel => ({ rel, escritas }));
};

(async () => {
  // 1. Caso normal: fábrica mudou, nuvem recebe dados, sinal e imagem nova.
  let est = cenario(
    { dados: { data: CHEIO, updated_at: 't2' }, sinal: { updated_at: 't2', key_versions: { ordens: 't2' } }, imagens: ['a.png', 'b.png'] },
    { dados: { data: CHEIO, updated_at: 't1' }, imagens: ['a.png'] });
  let { rel, escritas } = await rodar(est);
  ok('1. espelhou os dados', rel.dados === 'espelhado');
  ok('1b. gravou na NUVEM, nunca na fábrica',
     escritas.every(e => e.alvo === 'nuvem'), JSON.stringify(escritas.map(e => e.alvo)));
  ok('1c. levou o carimbo da fábrica',
     escritas.find(e => e.tipo === 'dados').corpo.updated_at === 't2');
  ok('1d. levou o sinal (para a nuvem tambem ler so o que mudou)',
     !!escritas.find(e => e.tipo === 'sinal'));
  ok('1e. subiu SO a imagem que faltava', rel.imagens === 1
     && escritas.filter(e => e.tipo === 'imagem').length === 1
     && escritas.find(e => e.tipo === 'imagem').nome === 'b.png');

  // 2. Nada mudou: não escreve nada. Espelhar de novo o mesmo não pode custar.
  est = cenario({ dados: { data: CHEIO, updated_at: 't2' }, imagens: ['a.png'] },
                { dados: { data: CHEIO, updated_at: 't2' }, imagens: ['a.png'] });
  ({ rel, escritas } = await rodar(est));
  ok('2. carimbo igual -> não reescreve', rel.dados === 'nao-mexeu'
     && !escritas.some(e => e.tipo === 'dados'), rel.motivo);

  // 3. A TRAVA QUE IMPORTA: fábrica vazia não pode apagar a nuvem cheia.
  est = cenario({ dados: { data: VAZIO, updated_at: 't9' } },
                { dados: { data: CHEIO, updated_at: 't1' } });
  ({ rel, escritas } = await rodar(est));
  ok('3. fábrica vazia NÃO apaga a nuvem cheia', rel.dados === 'bloqueado');
  ok('3b. e não escreveu absolutamente nada', escritas.length === 0);
  ok('3c. o motivo é explicado', /única cópia boa/.test(rel.motivo || ''), rel.motivo);

  // 4. Primeiro espelho de uma fábrica ainda vazia: pode, não há o que perder.
  est = cenario({ dados: { data: VAZIO, updated_at: 't1' } }, { dados: null });
  ({ rel } = await rodar(est));
  ok('4. nuvem vazia aceita o primeiro espelho', rel.dados === 'espelhado');

  // 5. Fábrica sem linha nenhuma: não inventa dado na nuvem.
  est = cenario({ dados: null }, { dados: { data: CHEIO, updated_at: 't1' } });
  ({ rel, escritas } = await rodar(est));
  ok('5. fábrica sem dados -> não mexe na nuvem',
     rel.dados === 'nao-mexeu' && escritas.length === 0);

  // 6. Internet fora: falha limpa, sem ter escrito pela metade.
  est = cenario({ dados: { data: CHEIO, updated_at: 't2' } }, { dados: null }, { nuvemFora: true });
  let erro = null;
  try { await rodar(est); } catch (e) { erro = e.message; }
  ok('6. nuvem inacessível falha sem escrever', erro !== null, erro);

  // 7. --forcar reescreve mesmo com carimbo igual (para reconstruir a nuvem).
  est = cenario({ dados: { data: CHEIO, updated_at: 't2' }, imagens: [] },
                { dados: { data: CHEIO, updated_at: 't2' }, imagens: [] });
  ({ rel } = await rodar(est, { forcar: true }));
  ok('7. --forcar reescreve mesmo sem mudança', rel.dados === 'espelhado');

  // 8. blobVazio: o critério da trava precisa estar certo.
  ok('8. blobVazio reconhece cheio e vazio',
     blobVazio(VAZIO) === true && blobVazio(CHEIO) === false
     && blobVazio(null) === true && blobVazio({}) === true);
  ok('8b. só desenhos já conta como não-vazio',
     blobVazio({ ordens: '[]', desenhos: '[{"id":1}]' }) === false);

  /* AS MENSAGENS. Elas moram em tabela propria (o canal de recados), e nao no
     blob: recado novo NAO muda o carimbo do shared_data, entao este passo tem
     de rodar mesmo quando os dados nao mudaram. */
  est = cenario({ dados: { data: CHEIO, updated_at: 't9' }, mensagens: [msg('1', 'oi'), msg('2', 'faltou pano')] },
                { dados: { data: CHEIO, updated_at: 't9' }, mensagens: [msg('1', 'oi')] });
  ({ rel, escritas } = await rodar(est));
  ok('9. mensagem nova sobe mesmo com os dados iguais (carimbo nao muda com recado)',
     rel.dados === 'nao-mexeu' && rel.mensagens === 1, JSON.stringify(rel));
  ok('10. e sobe SO a que faltava',
     escritas.filter(e => e.tipo === 'msg-enviar').flatMap(e => e.ids).join(',') === '2',
     JSON.stringify(escritas.filter(e => e.tipo === 'msg-enviar')));

  // Recado apagado na fabrica sai da nuvem: quem apagou o proprio recado fez
  // isso de proposito, e deixa-lo legivel de fora desfaria a decisao.
  est = cenario({ dados: { data: CHEIO, updated_at: 't9' }, mensagens: [msg('1', 'oi')] },
                { dados: { data: CHEIO, updated_at: 't9' }, mensagens: [msg('1', 'oi'), msg('3', 'engano')] });
  ({ rel, escritas } = await rodar(est));
  ok('11. o que foi apagado na fabrica e apagado na nuvem',
     rel.mensagensApagadas === 1
     && escritas.some(e => e.tipo === 'msg-apagar' && e.ids.join(',') === '3'), JSON.stringify(rel));

  // Um lado sem a tabela nao pode derrubar o espelho dos DADOS, que e o que importa.
  est = cenario({ dados: { data: CHEIO, updated_at: 't10' }, mensagens: [msg('1', 'oi')] },
                { dados: { data: CHEIO, updated_at: 't9' } }, { nuvemSemMensagens: true });
  ({ rel } = await rodar(est));
  ok('12. servidor sem a tabela de mensagens nao derruba o espelho dos dados',
     rel.dados === 'espelhado' && rel.mensagens === 0, JSON.stringify(rel));

  // E da para desligar as mensagens sem desligar o resto.
  est = cenario({ dados: { data: CHEIO, updated_at: 't11' }, mensagens: [msg('1', 'oi')] },
                { dados: { data: CHEIO, updated_at: 't9' } });
  ({ rel, escritas } = await rodar(est, { semMensagens: true }));
  ok('13. --sem-mensagens espelha os dados e nao toca no canal',
     rel.dados === 'espelhado' && !escritas.some(e => String(e.tipo).startsWith('msg-')),
     JSON.stringify(escritas.map(e => e.tipo)));

  console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
  process.exit(falhas ? 1 : 0);
})();
