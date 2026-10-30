/* Rode com:  node testes/vendas-dre.js

   VENDAS DAS PLATAFORMAS PELO PAINEL DRE (30/10/2026). O que este teste
   protege (carrega o estoque-confeccao.js e o vendas-dre.js de verdade, com
   um Painel DRE de mentira):
     · o SKU da plataforma (TAMANHO-TIPO.ESTAMPA-COR) vira o SKU daqui;
     · loja desligada e pedido com data antes do corte não lançam nada;
     · o mesmo pedido chegando de novo não baixa duas vezes;
     · pedido que deixa de ser faturado (cancelado) é estornado;
     · falta de estoque lança o que der e completa quando a peça entra;
     · SKU desconhecido vira pendência e a regra "De → Para" resolve
       (inclusive kit e item sem SKU, pela descrição);
     · o movimento apagado à mão dispensa o pedido;
     · o cursor anda página por página, é guardado como veio, e o cursor
       parado não vira laço (nem lança nada);
     · com a trava de outro computador valendo, esta aba não sincroniza. */
const fs = require('fs');
const path = require('path');
const raiz = path.join(__dirname, '..');
const modEstoque = fs.readFileSync(path.join(raiz, 'estoque-confeccao.js'), 'utf8');
const modDre = fs.readFileSync(path.join(raiz, 'vendas-dre.js'), 'utf8');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const INICIO = '2026-10-01T00:00:00.000Z';
function carrega(paginas) {
  const STATE = {
    meta: { acabConfig: { inicio: INICIO } },
    acabProdutos: [
      { id: 'p1', sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', tamanhos: ['P', 'M', 'G', 'GG', 'G1'] },
      { id: 'p2', sku: 'CM.TRI-CAQUI', linha: 'CM.TRI', cor: 'CAQUI', tamanhos: ['P', 'M', 'G'] },
      { id: 'p3', sku: 'BM.TRI-BEGE', linha: 'BM.TRI', cor: 'BEGE', tamanhos: ['P', 'M', 'G'] },
      { id: 'p4', sku: 'PM.LISA-AZUL', linha: 'PM.LISA', cor: 'AZUL', tamanhos: ['P', 'M', 'G', 'GG'] },
      { id: 'p5', sku: 'BM.LISA-PRE', linha: 'BM.LISA', cor: 'PRE', tamanhos: ['P', 'M', 'G'] }
    ],
    acabMov: [], ordens: [], desenhos: [], modelos: [], cores: []
  };
  // Estoque de partida: 10 de cada, menos o CM.TRI-CAQUI G (1 peça).
  let n = 0;
  STATE.acabProdutos.forEach(p => p.tamanhos.forEach(t => STATE.acabMov.push({ id: 'e' + (++n), tipo: 'entrada', origem: 'manual',
    produtoId: p.id, tamanho: t, data: '2026-10-02', qtd: p.id === 'p2' && t === 'G' ? 1 : 10, custo: 0, criadoEm: '2026-10-02T10:00:00.000Z' })));
  const window = {};
  const gravou = [], toasts = [], chamadas = [];
  const fila = (paginas || []).slice();
  Object.assign(globalThis, {
    window, STATE,
    document: { getElementById: () => null, querySelector: () => null },
    exigirEdicao: () => true,
    toast: (t, k) => toasts.push((k || '') + ':' + t),
    saveState: async k => { gravou.push(k); },
    desfazerNomearAcao: () => {},
    confirm: () => true,
    goto: () => {},
    formatDate: iso => iso,
    esc: s => String(s == null ? '' : s),
    uid: () => 'id_' + (++n),
    totaisPorTamanhoTomOS: () => ({ colTotal: () => 0 }),
    skusDaOS: () => [], _statusOS: () => '',
    catalogoSkus: [],
    currentUser: { email: 'admin@x' }, currentRole: 'admin',
    podeGravar: () => true, contaTemAcesso: () => true,
    _obsQuemSou: () => 'admin',
    fetch: async (url, op) => {
      chamadas.push({ url, chave: op && op.headers && op.headers['x-chave'] });
      const r = fila.length ? fila.shift() : { pedidos: [], cursor: null, maisPaginas: false };
      return { status: 200, ok: true, json: async () => r };
    }
  });
  new Function('window', 'STATE', modEstoque)(window, STATE);
  new Function('window', 'STATE', modDre)(window, STATE);
  const I = window._vendasDreInterno;
  const c = I.dre();
  c.chave = 'chave-teste';
  c.lojas = { L1: { ligada: true, corte: '2026-10-20', nome: 'Diverse Estilos', plataforma: 'shopify' },
    L2: { ligada: false, corte: '', nome: 'Hungria Store', plataforma: 'mercadolivre' } };
  return { STATE, window, I, c, gravou, toasts, chamadas, fila };
}
const pedido = (chave, itens, extra) => Object.assign({ id: chave, chave, integracao_id: 'L1', plataforma: 'shopify', loja: 'Diverse Estilos',
  numero: chave.replace(/\D/g, ''), status: 'pago', faturado: true, data: '2026-10-25T15:00:00.000Z', atualizado_em: '2026-10-25T15:00:00.000Z',
  itens }, extra || {});
const saidasDe = (STATE, chave) => STATE.acabMov.filter(m => m.origem === 'dre' && (!chave || m.pedidoChave === chave));
const somaSaidas = (STATE, chave) => saidasDe(STATE, chave).reduce((s, m) => s + m.qtd, 0);

(async () => {
  console.log('-- o SKU da plataforma vira o SKU daqui --');
  {
    const { I, STATE } = carrega();
    const cv = sku => { const r = I.converter({ sku, descricao: '', quantidade: 1 }, []); return r.erro ? 'ERRO: ' + r.erro : r.partes.map(x => x.produto.sku + ' ' + x.tamanho).join(','); };
    ok('M-BM.TRI-BEGE → BM.TRI-BEGE M', cv('M-BM.TRI-BEGE') === 'BM.TRI-BEGE M', cv('M-BM.TRI-BEGE'));
    ok('G1-CM.TRI.LA.MANGA-PRE → CM.TRI-PRE (não cadastrado = erro)', /CM\.TRI-PRE não está cadastrado/.test(cv('G1-CM.TRI.LA.MANGA-PRE')), cv('G1-CM.TRI.LA.MANGA-PRE'));
    ok('GG-PM.76.ITALY-AZUL → PM.LISA-AZUL GG', cv('GG-PM.76.ITALY-AZUL') === 'PM.LISA-AZUL GG', cv('GG-PM.76.ITALY-AZUL'));
    ok('P-CM.ITALIA.MILANO-MARINHO → CM.LISA-MARINHO (não cadastrado)', /CM\.LISA-MARINHO não está cadastrado/.test(cv('P-CM.ITALIA.MILANO-MARINHO')), cv('P-CM.ITALIA.MILANO-MARINHO'));
    ok('G-CM.TRI.LA.MANGA-CÁQUI → CM.TRI-CAQUI G (acento sai)', cv('G-CM.TRI.LA.MANGA-CÁQUI') === 'CM.TRI-CAQUI G', cv('G-CM.TRI.LA.MANGA-CÁQUI'));
    ok('CM.TRIX não é CM.TRI (vai para CM.LISA)', I.linhaDe('CM.TRIX.Y') === 'CM.LISA', I.linhaDe('CM.TRIX.Y'));
    ok('item sem SKU → erro pedindo regra', /sem SKU/.test(cv('')), cv(''));
    ok('tamanho na descrição: "Camiseta Paris - GG"', I.tamDaDescricao('Camiseta Paris - GG') === 'GG');
    ok('"KIT 2 CAMISETAS M" → M, não 2', I.tamDaDescricao('KIT 2 CAMISETAS M') === 'M', I.tamDaDescricao('KIT 2 CAMISETAS M'));
    void STATE;
  }

  console.log('-- loja desligada e pedido antes do corte --');
  {
    const { STATE, I, c } = carrega();
    I.processar([pedido('A1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 2 }], { integracao_id: 'L2' }),
      pedido('A2', [{ sku: 'M-BM.TRI-BEGE', quantidade: 2 }], { data: '2026-10-19T23:00:00.000Z' })], c);
    ok('nada lançado', saidasDe(STATE).length === 0, saidasDe(STATE));
    // 20/10 00:30 em São Paulo é 20/10 03:30Z; 19/10 23:00Z é 19/10 20:00 em SP.
    I.processar([pedido('A3', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }], { data: '2026-10-20T03:30:00.000Z' })], c);
    ok('o dia do corte conta (fuso de São Paulo)', somaSaidas(STATE, 'A3') === 1);
    ok('a data do movimento é o dia do pedido em SP', saidasDe(STATE, 'A3')[0].data === '2026-10-20', saidasDe(STATE, 'A3')[0]);
  }

  console.log('-- repetição, cancelamento e formato do movimento --');
  {
    const { STATE, I, c } = carrega();
    const p = pedido('B1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 2 }, { sku: 'GG-PM.76.ITALY-AZUL', quantidade: 1 }]);
    I.processar([p], c);
    ok('lança 3 peças em 2 movimentos', somaSaidas(STATE, 'B1') === 3 && saidasDe(STATE, 'B1').length === 2, saidasDe(STATE, 'B1'));
    const m = saidasDe(STATE, 'B1')[0];
    ok('movimento: saida, cliente, valor 0, pedido, loja e plataforma', m.tipo === 'saida' && m.destino === 'cliente' && m.valor === 0
      && m.pedidoChave === 'B1' && m.pedidoNumero === '1' && m.loja === 'Diverse Estilos' && m.plataforma === 'shopify' && m.origem === 'dre', m);
    I.processar([p], c); I.processar([p], c);
    ok('o mesmo pedido 3 vezes: continua 3 peças', somaSaidas(STATE, 'B1') === 3, somaSaidas(STATE, 'B1'));
    I.processar([Object.assign({}, p, { itens: [{ sku: 'M-BM.TRI-BEGE', quantidade: 3 }, { sku: 'GG-PM.76.ITALY-AZUL', quantidade: 1 }] })], c);
    ok('pedido que passou de 2 para 3: lança só +1', somaSaidas(STATE, 'B1') === 4 && saidasDe(STATE, 'B1').length === 2, saidasDe(STATE, 'B1'));
    I.processar([Object.assign({}, p, { faturado: false, status: 'cancelado' })], c);
    ok('cancelado: estorna tudo (os movimentos somem)', saidasDe(STATE, 'B1').length === 0, saidasDe(STATE, 'B1'));
    I.processar([p], c);
    ok('voltou a ser faturado: lança de novo', somaSaidas(STATE, 'B1') === 3);
  }

  console.log('-- falta de estoque --');
  {
    const { STATE, I, c } = carrega();
    const p = pedido('C1', [{ sku: 'G-CM.TRI.LA.MANGA-CÁQUI', quantidade: 3 }]);
    I.processar([p], c);
    ok('há 1: lança 1', somaSaidas(STATE, 'C1') === 1, somaSaidas(STATE, 'C1'));
    const pe = c.pendentes.C1;
    ok('o resto (2) fica em pendências como falta', pe && pe.problemas.length === 1 && pe.problemas[0].tipo === 'falta' && pe.problemas[0].qtd === 2, pe);
    STATE.acabMov.push({ id: 'e-novo', tipo: 'entrada', origem: 'manual', produtoId: 'p2', tamanho: 'G', data: '2026-10-26', qtd: 5, custo: 0, criadoEm: '2026-10-26T10:00:00.000Z' });
    I.processar(Object.values(c.pendentes), c);
    ok('a peça entrou: completa os 3', somaSaidas(STATE, 'C1') === 3, somaSaidas(STATE, 'C1'));
    ok('e sai das pendências', !c.pendentes.C1, c.pendentes.C1);
    ok('continua um movimento só do pedido', saidasDe(STATE, 'C1').length === 1, saidasDe(STATE, 'C1'));
  }

  console.log('-- SKU desconhecido e a tabela De → Para --');
  {
    const { STATE, I, c } = carrega();
    const p = pedido('D1', [{ sku: 'M-KIT.PARIS-PRE', quantidade: 1 }, { sku: '', descricao: 'Camiseta Lisa Preta - G', quantidade: 2 }, { sku: 'M-BM.TRI-BEGE', quantidade: 1 }]);
    I.processar([p], c);
    ok('o item conhecido sai', somaSaidas(STATE, 'D1') === 1);
    ok('os dois desconhecidos ficam pendentes', c.pendentes.D1 && c.pendentes.D1.problemas.filter(x => x.tipo === 'sku').length === 2, c.pendentes.D1);
    const kit = I.lerPara('CM.LISA-PRE x1, BM.LISA-PRE x1');
    ok('lerPara do kit', !kit.erro && kit.para.length === 2, kit);
    ok('lerPara recusa SKU não cadastrado', !!I.lerPara('XX.Y-ZZ').erro);
    c.mapa.push({ id: 'r1', tipo: 'sku', de: 'KIT.PARIS', para: kit.para });
    c.mapa.push({ id: 'r2', tipo: 'descricao', de: 'CAMISETA LISA PRETA', para: [{ sku: 'CM.LISA-PRE', qtd: 1 }] });
    I.processar(Object.values(c.pendentes), c);
    const porSku = STATE.acabMov.filter(m => m.pedidoChave === 'D1').map(m => m.produtoId + m.tamanho + ':' + m.qtd).sort().join(' ');
    ok('kit (CM.LISA-PRE M + BM.LISA-PRE M) e descrição (CM.LISA-PRE G x2)', porSku === 'p1G:2 p1M:1 p3M:1 p5M:1', porSku);
    ok('pendência resolvida', !c.pendentes.D1, c.pendentes.D1);
    c.mapa.push({ id: 'r3', tipo: 'sku', de: 'BRINDE', ignorar: true });
    I.processar([pedido('D2', [{ sku: 'U-BRINDE', quantidade: 1 }, { sku: 'M-BRINDE.ADESIVO', quantidade: 1 }])], c);
    ok('regra "ignorar": nada lançado, nada pendente', somaSaidas(STATE, 'D2') === 0 && !c.pendentes.D2, c.pendentes.D2);
  }

  console.log('-- apagar à mão dispensa o pedido --');
  {
    const { STATE, I, c, window } = carrega();
    const p = pedido('E1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }]);
    I.processar([p], c);
    const id = saidasDe(STATE, 'E1')[0].id;
    await window.acabApagarMov(id);
    ok('apagado', saidasDe(STATE, 'E1').length === 0);
    ok('pedido dispensado', !!c.dispensados.E1);
    I.processar([p], c);
    ok('a sincronização não relança', saidasDe(STATE, 'E1').length === 0, saidasDe(STATE, 'E1'));
  }

  console.log('-- a rodada: páginas, cursor guardado como veio, laço --');
  {
    const cur1 = { desde: '2026-10-25T15:00:00.123456+00:00', apos: 'B1' };
    const cur2 = { desde: '2026-10-25T16:00:00.654321+00:00', apos: 'B3' };
    const { STATE, I, c, chamadas, gravou } = carrega([
      { pedidos: [pedido('B1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }]), pedido('Z9', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }], { data: '2026-06-01T10:00:00.000Z' })], cursor: cur1, maisPaginas: true },
      { pedidos: [pedido('B3', [{ sku: 'P-CM.LISA.QQ-PRE', quantidade: 2 }])], cursor: cur2, maisPaginas: false }
    ]);
    await I.rodar();
    ok('duas páginas pedidas, com a chave no cabeçalho', chamadas.length === 2 && chamadas.every(x => x.chave === 'chave-teste'), chamadas);
    ok('a primeira começa no corte da loja', /desde=2026-10-20T03%3A00%3A00.000Z/.test(chamadas[0].url), chamadas[0].url);
    ok('a segunda usa o cursor como veio (microssegundos e apos)', chamadas[1].url.includes(encodeURIComponent(cur1.desde)) && chamadas[1].url.includes('apos=B1'), chamadas[1].url);
    ok('cursor final guardado sem conversão', c.cursor.desde === cur2.desde && c.cursor.apos === 'B3', c.cursor);
    ok('pedido antigo (junho) ignorado', somaSaidas(STATE, 'Z9') === 0);
    ok('B1 e B3 baixados', somaSaidas(STATE, 'B1') === 1 && somaSaidas(STATE, 'B3') === 2);
    ok('sem erro no status', !c.status.erro, c.status);
    ok('gravou acabMov e meta', gravou.includes('acabMov') && gravou.includes('meta'), gravou);
  }
  {
    const preso = { desde: '2026-10-25T15:00:00.1+00:00', apos: 'X' };
    const { STATE, I, c } = carrega([
      { pedidos: [pedido('F1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }])], cursor: preso, maisPaginas: true },
      { pedidos: [pedido('F1', [{ sku: 'M-BM.TRI-BEGE', quantidade: 1 }])], cursor: preso, maisPaginas: true }
    ]);
    c.cursor = Object.assign({}, preso);
    await I.rodar();
    ok('cursor parado: para com erro', /não andou/.test(c.status.erro || ''), c.status);
    ok('e não lança nada', somaSaidas(STATE) === 0);
  }
  {
    const { I, c, chamadas } = carrega([]);
    c.cursor = { desde: '2026-10-28T00:00:00.000Z', apos: '' };
    c.lojas.L2.ligada = true; c.lojas.L2.corte = '2026-10-22';
    // ligar/recuar o corte volta o cursor (ver recuarCursor) — aqui direto:
    await window.vendasDreCorte('L2', '2026-10-22');
    ok('recuar o corte volta o cursor para o corte', c.cursor.desde === new Date('2026-10-22T00:00:00-03:00').toISOString(), c.cursor);
    void I; void chamadas;
  }

  console.log('-- a trava --');
  {
    const { I, c, chamadas } = carrega([]);
    c.trava = { quem: 'nathaly', aba: 'outra', ate: new Date(Date.now() + 60000).toISOString() };
    I._set({ ultima: 0, pegouEm: 0 });
    await window.vendasDreRodar(true);
    ok('trava de outro valendo: não pede nada', chamadas.length === 0, chamadas);
    c.trava.ate = new Date(Date.now() - 1000).toISOString();
    await window.vendasDreRodar(true);
    ok('trava vencida: assume e roda', chamadas.length === 1 && c.trava.aba === I.ABA, { chamadas, trava: c.trava });
  }
  {
    // Sem forçar: pega a trava e espera uma leitura antes de rodar.
    const { I, c, chamadas } = carrega([]);
    I._set({ ultima: 0, pegouEm: 0 });
    globalThis.Date.now = (orig => () => orig() + 10 * 60000)(Date.now.bind(Date));
    await window.vendasDreRodar();
    ok('primeira passada só pega a trava', chamadas.length === 0 && c.trava && c.trava.aba === I.ABA, { chamadas, trava: c.trava });
    await window.vendasDreRodar();
    ok('logo depois ainda espera a confirmação', chamadas.length === 0, chamadas);
  }

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
