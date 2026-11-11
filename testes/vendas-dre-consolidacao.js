/* Rode com:  node testes/vendas-dre-consolidacao.js [backup.json]

   CONSOLIDAÇÃO DAS VENDAS ANTIGAS (11/11/2026). Carrega o estoque-confeccao.js
   e o vendas-dre.js de verdade e confere que:
     · as saídas do DRE de mais de 30 dias viram totais, e as recentes ficam;
     · o saldo, o custo médio e o CMV de cada SKU não mudam — conferido aqui
       com a conta do próprio estoque, independente da conferência interna;
     · o pedido antigo que o DRE reenviar não baixa de novo;
     · o pedido pendente fica de fora e continua podendo se completar;
     · roda uma vez por dia.
   Com um backup como argumento, mede a redução com os dados reais. */
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

function carrega(STATE) {
  const window = {};
  let n = 0;
  Object.assign(globalThis, {
    window, STATE,
    document: { getElementById: () => null, querySelector: () => null },
    exigirEdicao: () => true, toast: () => {}, saveState: async () => {}, desfazerNomearAcao: () => {},
    confirm: () => true, goto: () => {}, formatDate: iso => iso, esc: s => String(s == null ? '' : s),
    uid: () => 'novo_' + (++n),
    totaisPorTamanhoTomOS: () => ({ colTotal: () => 0 }), skusDaOS: () => [], _statusOS: () => '',
    catalogoSkus: [], currentUser: { email: 'admin@x' }, currentRole: 'admin',
    podeGravar: () => true, contaTemAcesso: () => true, _obsQuemSou: () => 'admin',
    fetch: async () => ({ status: 200, ok: true, json: async () => ({ pedidos: [], cursor: null, maisPaginas: false }) })
  });
  new Function('window', 'STATE', modEstoque)(window, STATE);
  new Function('window', 'STATE', modDre)(window, STATE);
  return { I: window._vendasDreInterno, acab: window._acabInterno };
}

// A conta do estoque, para comparar antes e depois.
function retrato(acab) {
  const ms = acab.todosMovs();
  const S = acab.saldos(ms);
  const out = {};
  const pares = new Set(ms.map(m => m.produtoId + '|' + m.tamanho));
  pares.forEach(k => {
    const [pid, t] = k.split('|');
    const r = acab.cmp(ms, pid, t);
    out[k] = [S.de(pid, t).total, +r.medio.toFixed(6), +r.cmv.toFixed(6)];
    const g = acab.cmp(ms, pid, null);
    out[pid + '|*'] = [+g.medio.toFixed(6), +g.cmv.toFixed(6)];
  });
  return out;
}

(async () => {
  console.log('-- vendas espalhadas por dois meses, com entradas de custo diferente no meio --');
  const STATE = {
    meta: { acabConfig: { inicio: '2026-08-01T00:00:00.000Z' } },
    acabProdutos: [
      { id: 'p1', sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', tamanhos: ['P', 'M', 'G'] },
      { id: 'p2', sku: 'CM.TRI-CAQUI', linha: 'CM.TRI', cor: 'CAQUI', tamanhos: ['P', 'M', 'G'] }
    ],
    acabMov: [], ordens: [], desenhos: [], modelos: [], cores: []
  };
  let k = 0;
  const ent = (pid, t, data, qtd, custo) => STATE.acabMov.push({ id: 'e' + (++k), tipo: 'entrada', origem: 'manual', produtoId: pid, tamanho: t, data, qtd, custo, criadoEm: data + 'T10:00:00.000Z' });
  const sai = (pid, t, data, qtd, plat, chave) => STATE.acabMov.push({ id: 's' + (++k), tipo: 'saida', origem: 'dre', produtoId: pid, tamanho: t, data, qtd, valor: 0,
    destino: 'cliente', pedidoChave: chave || ('L1:' + k), pedidoNumero: String(k), loja: 'Loja ' + plat, plataforma: plat, criadoEm: data + 'T15:00:00.000Z' });
  ['P', 'M', 'G'].forEach(t => { ent('p1', t, '2026-09-01', 100, 20); ent('p2', t, '2026-09-01', 60, 30); });
  ent('p1', 'M', '2026-09-20', 50, 26);                   // entrada no meio: corta o trecho do p1
  ent('p2', 'G', '2026-10-05', 40, 35);
  ent('p1', 'P', '2026-10-15', 30, 22);                   // depois do corte
  const plats = ['shopify', 'mercadolivre', 'shopee', 'tiktok'];
  for (let d = 2; d <= 70; d++) {                         // 02/09 a 10/11
    const dia = new Date(Date.UTC(2026, 8, d)).toISOString().slice(0, 10);
    ['P', 'M', 'G'].forEach((t, i) => {
      sai('p1', t, dia, 1 + ((d + i) % 2), plats[(d + i) % 4]);
      if (d % 3 === 0) sai('p2', t, dia, 1, plats[d % 4]);
    });
  }
  // Um pedido antigo ainda pendente (lançou parte; falta o resto).
  sai('p2', 'P', '2026-09-10', 1, 'shopify', 'L1:PENDENTE');
  // Um movimento de antes do início da contagem: não conta, não entra.
  STATE.acabMov.push({ id: 'velho', tipo: 'saida', origem: 'dre', produtoId: 'p1', tamanho: 'P', data: '2026-07-20', qtd: 5,
    pedidoChave: 'L1:VELHO', criadoEm: '2026-07-20T10:00:00.000Z' });

  const { I, acab } = carrega(STATE);
  const c = I.dre();
  c.lojas = { L1: { ligada: true, corte: '2026-09-01', nome: 'Loja', plataforma: 'shopify' } };
  c.pendentes = { 'L1:PENDENTE': { chave: 'L1:PENDENTE', data: '2026-09-10T15:00:00.000Z', integracao_id: 'L1', itens: [] } };

  const antes = retrato(acab);
  const nAntes = STATE.acabMov.length;
  const dreAntigas = STATE.acabMov.filter(m => m.origem === 'dre' && m.data < '2026-10-12' && m.pedidoChave !== 'L1:PENDENTE' && m.id !== 'velho').length;
  const recentesAntes = STATE.acabMov.filter(m => m.origem === 'dre' && m.data >= '2026-10-12').length;

  const r = I.consolidar(c, '2026-11-11');
  ok('o corte é 30 dias antes de hoje (12/10)', r && r.ate === '2026-10-12', r);
  ok('consolidou e gravaria', r.mudou && !r.erro, r);
  ok(`juntou as ${dreAntigas} saídas antigas`, r.antes === dreAntigas, r);
  ok('o acabMov encolheu', STATE.acabMov.length < nAntes - dreAntigas / 2, [nAntes, STATE.acabMov.length]);
  ok('as recentes (12/10 em diante) ficam pedido a pedido',
    STATE.acabMov.filter(m => m.origem === 'dre' && m.data >= '2026-10-12').length === recentesAntes);
  ok('o pendente fica de fora', STATE.acabMov.some(m => m.pedidoChave === 'L1:PENDENTE'));
  ok('o de antes do início fica como está', STATE.acabMov.some(m => m.id === 'velho'));
  const depois = retrato(acab);
  const dif = Object.keys(antes).filter(x => JSON.stringify(antes[x]) !== JSON.stringify(depois[x]));
  ok('saldo, custo médio e CMV iguais em todo SKU e tamanho', dif.length === 0 && Object.keys(depois).length === Object.keys(antes).length,
    dif.slice(0, 3).map(x => [x, antes[x], depois[x]]));
  const cons = STATE.acabMov.filter(m => m.origem === 'dre-consolidado');
  ok('nenhum total atravessa uma entrada do produto (p1 em 20/09)',
    !cons.some(m => m.produtoId === 'p1' && m.data < '2026-09-20' && m.ate >= '2026-09-20'));
  ok('o total guarda pedidos e plataformas', cons.every(m => m.pedidos > 0 && Object.values(m.porPlataforma).reduce((a, b) => a + b, 0) === m.qtd));
  ok('consolidadoAte anotado', c.consolidadoAte === '2026-10-12', c.consolidadoAte);

  console.log('\n-- o DRE reenvia pedidos antigos --');
  const pedido = (chave, data) => ({ chave, id: chave, integracao_id: 'L1', plataforma: 'shopify', loja: 'Loja', numero: '1', faturado: true, data,
    itens: [{ sku: 'M-CM.LISA-PRE', quantidade: 1 }] });
  ok('pedido de antes do corte é ignorado', !I.relevante(pedido('L1:9', '2026-10-01T15:00:00.000Z'), c));
  ok('pedido recente continua valendo', I.relevante(pedido('L1:10', '2026-10-20T15:00:00.000Z'), c));
  ok('o pendente antigo continua valendo', I.relevante(pedido('L1:PENDENTE', '2026-09-10T15:00:00.000Z'), c));
  const nMov = STATE.acabMov.length;
  I.processar([pedido('L1:5', '2026-09-03T15:00:00.000Z')], c);
  ok('reenviado não baixa de novo', STATE.acabMov.length === nMov && JSON.stringify(retrato(acab)) === JSON.stringify(depois));

  console.log('\n-- uma vez por dia --');
  ok('segunda chamada no mesmo dia não faz nada', I.consolidar(c, '2026-11-11') === null);
  const r2 = I.consolidar(c, '2026-11-12');
  ok('no dia seguinte anda mais um dia, e a conta segue igual',
    r2 && !r2.erro && JSON.stringify(retrato(acab)) === JSON.stringify(depois), r2);

  const bkp = process.argv[2];
  if (bkp) {
    console.log('\n-- dados reais:', path.basename(bkp), '--');
    const d = JSON.parse(fs.readFileSync(bkp, 'utf8'));
    const B = d.data || d;
    const parse = v => (typeof v === 'string' ? JSON.parse(v) : v);
    const S2 = { meta: parse(B.meta) || {}, acabProdutos: parse(B.acabProdutos) || [], acabMov: parse(B.acabMov) || [],
      ordens: [], desenhos: [], modelos: [], cores: [] };
    const { I: I2, acab: a2 } = carrega(S2);
    const c2 = I2.dre();
    const bytes = () => JSON.stringify(S2.acabMov).length;
    const b0 = bytes(), n0 = S2.acabMov.length, f0 = retrato(a2);
    // "Hoje" 30 dias depois do último dia do backup: tudo dele já é antigo.
    const ult = S2.acabMov.map(m => String(m.data || '')).sort().pop();
    const hoje = I2.diaMenos(ult, -31);
    const r3 = I2.consolidar(c2, hoje);
    const f1 = retrato(a2);
    const dif3 = Object.keys(f0).filter(x => JSON.stringify(f0[x]) !== JSON.stringify(f1[x]));
    console.log(`  hoje simulado ${hoje}: ${r3.antes} saídas do DRE -> ${r3.depois} totais${r3.erro ? ' · ' + r3.erro : ''}`);
    console.log(`  acabMov: ${n0} movimentos, ${(b0 / 1048576).toFixed(2)} MB -> ${S2.acabMov.length}, ${(bytes() / 1048576).toFixed(2)} MB`);
    ok('dados reais: conta igual em todo SKU', dif3.length === 0 && !r3.erro, dif3.slice(0, 3));
  }

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
