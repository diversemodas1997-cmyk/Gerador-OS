/* Rode com:  node testes/estoque-confeccao.js

   O ESTOQUE CONFECÇÃO (21/10/2026): o produto acabado que veio do programa
   Estoque-Confeccao (Render) para dentro do Gerador-OS. O que este teste
   protege:
     · a OS que marca a etapa Estoque entra sozinha, por tamanho, pelo SKU dela
       — e só a partir do INÍCIO do estoque (começou do zero aqui);
     · a OS que não dá para lançar sozinha vai para "esperando SKU", com o motivo;
     · o saldo (entradas − saídas + ajustes) e o custo médio ponderado, com as
       mesmas regras do original;
     · a importação do ERP: casa pelo código e pela linha-cor, soma linhas
       repetidas, não inventa SKU sem pedir;
     · a Contabilidade recebe o estoque no formato que o Render publicava;
     · a ligação com o app.js (chaves, área de acesso, telas, snapshot).
   Carrega o estoque-confeccao.js de verdade, com o resto do programa de mentira. */
const fs = require('fs');
const path = require('path');

const raiz = path.join(__dirname, '..');
const mod = fs.readFileSync(path.join(raiz, 'estoque-confeccao.js'), 'utf8');
const app = fs.readFileSync(path.join(raiz, 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const css = fs.readFileSync(path.join(raiz, 'styles.css'), 'utf8');

const INICIO = '2026-10-21T12:00:00.000Z';
const ms = iso => Date.parse(iso);

function carrega(STATE) {
  const window = {};
  let n = 0;
  new Function('window', 'STATE', 'formatDate', 'esc', 'uid', 'totaisPorTamanhoTomOS', 'skusDaOS', '_statusOS',
    'catalogoSkus', 'currentRole', mod)(
    window, STATE,
    iso => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; },
    s => String(s == null ? '' : s),
    () => 'id_' + (++n),
    o => ({ colTotal: k => (o.porTam || {})[k] || 0 }),
    o => o.skus || [],
    o => o.statusOS || '',
    [], 'admin');
  return window;
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const osDe = (id, num, quando, extra) => Object.assign({
  id, os: num, modeloNome: 'Camiseta Básica', etapas: ['Corte', 'Estoque'],
  progresso: { etapasCheck: { Estoque: true }, etapasSeq: { Estoque: ms(quando) } },
  porTam: { m: 94, g: 94 }, skus: ['CM.LISA-PRE']
}, extra || {});
const P1 = { id: 'p1', codigo: 12, sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', descricao: 'CAMISETA BÁSICA PRETO', tamanhos: ['P', 'M', 'G'], preco: 19.9 };

console.log('-- a OS que marca a etapa Estoque --');
let STATE = {
  meta: { acabConfig: { inicio: INICIO } },
  acabProdutos: [P1], acabMov: [],
  ordens: [
    osDe('o1', '0656', '2026-10-21T13:00:00Z'),
    osDe('o2', '0600', '2026-10-20T09:00:00Z'),                                   // antes do início
    osDe('o3', '0659', '2026-10-21T14:00:00Z', { skus: ['CM.LISA-BRA'] }),       // SKU só do programa
    osDe('o4', '0661', '2026-10-21T15:00:00Z', { skus: ['BM.TRI-AZUL', 'BM.TRI-BEGE'] }),  // duas cores
    osDe('o5', '0662', '2026-10-21T15:00:00Z', { statusOS: 'cancelado' }),
    osDe('o6', '0663', '2026-10-21T15:00:00Z', { progresso: { etapasCheck: {}, etapasSeq: {} } }),  // não marcou
    osDe('o7', '0664', '2026-10-21T16:00:00Z', { skus: [] })                       // sem SKU
  ]
};
let w = carrega(STATE);
let r = w._acabInterno._entradasDasOS();
ok('1. a OS 0656 lança uma entrada por tamanho', r.entradas.filter(e => e.os === '0656').map(e => e.tamanho).join() === 'M,G', r.entradas);
ok('2. com a quantidade do Total por tamanho e o SKU da OS', r.entradas.find(e => e.tamanho === 'M').qtd === 94 && r.entradas[0].produtoId === 'p1', r.entradas);
ok('3. na data do carimbo da etapa', r.entradas[0].data === '2026-10-21', r.entradas[0].data);
ok('4. OS marcada antes do início não lança (o estoque começou do zero)', !r.entradas.some(e => e.os === '0600') && !r.pendencias.some(p => p.os === '0600'));
// Desde 23/10/2026 (5f52fdd) o SKU da OS é lido do programa ('sku:' + SKU): a
// OS entra no estoque sem esperar o SKU ser cadastrado aqui.
ok('5. SKU que só o programa conhece entra sozinho, sem "esperando SKU"',
   r.entradas.filter(e => e.os === '0659').every(e => e.produtoId === 'sku:CM.LISA-BRA')
   && r.entradas.filter(e => e.os === '0659').map(e => e.tamanho).join() === 'M,G'
   && !r.pendencias.some(p => p.os === '0659'), { entradas: r.entradas, pendencias: r.pendencias });
ok('6. OS de duas cores vai para "esperando SKU"', r.pendencias.some(p => p.os === '0661' && /mais de uma cor/.test(p.motivo)));
ok('7. OS sem SKU vai para "esperando SKU"', r.pendencias.some(p => p.os === '0664' && /não tem SKU/.test(p.motivo)));
ok('8. cancelada e não marcada ficam de fora', !r.entradas.concat(r.pendencias).some(x => x.os === '0662' || x.os === '0663'));
STATE.meta.acabConfig.osDispensadas = ['o3'];
ok('9. OS dispensada sai da lista', !w._acabInterno._entradasDasOS().pendencias.some(p => p.os === '0659'));
STATE.meta.acabConfig = {};
r = w._acabInterno._entradasDasOS();
ok('10. sem início marcado, nenhuma OS lança (o lado seguro)', r.semInicio && !r.entradas.length);

console.log('\n-- saldo e custo médio --');
STATE = { meta: { acabConfig: { inicio: INICIO } }, acabProdutos: [P1], ordens: [], acabMov: [
  { id: 'a', tipo: 'entrada', produtoId: 'p1', tamanho: 'M', data: '2026-10-21', qtd: 10, custo: 10 },
  { id: 'b', tipo: 'entrada', produtoId: 'p1', tamanho: 'M', data: '2026-10-22', qtd: 10, custo: 20 },
  { id: 'c', tipo: 'saida', produtoId: 'p1', tamanho: 'M', data: '2026-10-23', qtd: 5, valor: 39.9, destino: 'cliente' },
  { id: 'd', tipo: 'ajuste', produtoId: 'p1', tamanho: 'M', data: '2026-10-24', contado: 14, anterior: 15, diferenca: -1 }
] };
w = carrega(STATE);
const I = w._acabInterno;
const S = I.saldos(STATE.acabMov);
ok('11. saldo = entradas − saídas + ajustes', S.de('p1', 'M').total === 14, S.de('p1', 'M'));
ok('12. a última contagem fica guardada', S.de('p1', 'M').ult && S.de('p1', 'M').ult.contado === 14);
const c = I.cmp(STATE.acabMov, 'p1', 'M');
ok('13. custo médio ponderado (10×10 + 10×20) / 20 = 15', Math.abs(c.medio - 15) < 1e-9, c.medio);
ok('14. a saída sai pelo custo médio: CMV 5 × 15 = 75', Math.abs(c.cmv - 75) < 1e-9, c.cmv);
ok('15. sem entrada com custo, vale o preço do cadastro', I.cmp(STATE.acabMov, 'p1', 'G').medio === 19.9);

// 16 a 25 eram a importação do relatório do ERP (planoErp, lerCsvErp,
// extrairTipoCor): a tela Saídas do ERP saiu em 30/10/2026 — as vendas saem por
// Vendas das plataformas (testes/vendas-dre.js).

console.log('\n-- a Contabilidade --');
STATE.ordens = [osDe('o1', '0656', '2026-10-21T13:00:00Z')];
const snap = I.acabSnapshotContab();
ok('26. por SKU, a quantidade (lançamentos + OS), no formato do Render',
   snap.total === 14 + 188 && snap.itens.length === 1 && snap.itens[0].sku === 'CM.LISA-PRE' && snap.itens[0].codigo === 12 && 'preco' in snap.itens[0], snap);

console.log('\n-- os avisos do sino --');
{
  const st = { meta: { acabConfig: { inicio: INICIO } }, acabProdutos: [P1], ordens: [
      osDe('o1', '0656', '2026-10-21T13:00:00Z'),
      // A que espera SKU é a OS sem SKU nenhum (a de SKU só do programa já entra).
      osDe('o7', '0664', '2026-10-21T14:00:00Z', { skus: [] })],
    acabMov: [
      { id: 'e1', tipo: 'saida', origem: 'erp', produtoId: 'p1', tamanho: 'M', data: '2026-10-21', qtd: 3, erpArquivo: 'vendas.pdf', criadoEm: '2026-10-21T15:00:10Z', criadoPor: 'escritorio@diverse.local' },
      { id: 'e2', tipo: 'saida', origem: 'erp', produtoId: 'p1', tamanho: 'G', data: '2026-10-21', qtd: 2, erpArquivo: 'vendas.pdf', criadoEm: '2026-10-21T15:00:40Z', criadoPor: 'escritorio@diverse.local' },
      { id: 'e3', tipo: 'saida', origem: 'manual', produtoId: 'p1', tamanho: 'G', data: '2026-10-21', qtd: 1, criadoEm: '2026-10-21T15:05:00Z' }] };
  const av = carrega(st)._acabInterno.acabAvisos(ms('2026-10-01T00:00:00Z'));
  const tipos = av.map(e => e.tipo).sort().join(',');
  ok('35a. a OS que entrou, a que espera SKU e a importação do ERP', tipos === 'acab-erp,acab-os,acab-pend', av);
  ok('35b. a OS que entrou diz o SKU e as peças', av.find(e => e.tipo === 'acab-os').texto === 'CM.LISA-PRE · 188 peças', av);
  ok('35c. a importação vira um aviso só, com quem importou', av.filter(e => e.tipo === 'acab-erp').length === 1
     && /2 saídas · 5 peças/.test(av.find(e => e.tipo === 'acab-erp').texto) && av.find(e => e.tipo === 'acab-erp').quem === 'escritorio@diverse.local', av);
  const velho = carrega(st)._acabInterno.acabAvisos(ms('2026-10-22T00:00:00Z'));
  ok('35d. fora da janela de dias some, mas a OS esperando SKU continua', velho.map(e => e.tipo).join() === 'acab-pend', velho);
  ok('35e. os três tipos são da área Estoque confecção, com a coluna Avisar', /'acab-os': 'estoque-confeccao', 'acab-pend': 'estoque-confeccao', 'acab-erp': 'estoque-confeccao'/.test(app)
     && /'estoque-confeccao': 'OS que entrou no estoque/.test(app) && /acabAvisos\(limite\)/.test(app));
}

console.log('\n-- a ligação com o programa --');
ok('27. as chaves entram nas duas listas do app.js', (app.match(/'costuraProd','acabProdutos','acabMov'/g) || []).length === 2);
// Desde 28/10/2026 o Inventário é a parte de baixo de Entradas e saídas; desde
// 30/10/2026 há Vendas das plataformas (vendas-dre.js): seis telas.
ok('28. a área Estoque confecção existe, com as cinco telas (sem Saídas do ERP, com Vendas das plataformas)', /k: 'estoque-confeccao', rotulo: 'Estoque de produtos'/.test(app)
   && /'estoque-confeccao': \['acab-estoque', 'acab-mov', 'acab-produtos', 'acab-dre', 'acab-cmp'\]/.test(app)
   && /if \(page === 'acab-erp'\) page = 'acab-dre';/.test(app));
ok('29. as ações das telas têm área', ['lançar no estoque de produtos', 'apagar lançamento do estoque de produtos', 'cadastrar SKU no estoque de produtos', 'configurar o estoque de produtos']
   .every(a => app.includes(`'${a}': 'estoque-confeccao'`)) && ['lançar no estoque de produtos', 'apagar lançamento do estoque de produtos', 'cadastrar SKU no estoque de produtos', 'configurar o estoque de produtos']
   .every(a => mod.includes(`exigirEdicao('${a}')`)));
ok('30. o snapshot da Contabilidade leva o produto acabado', /produtosAcabados = acabSnapshotContab\(\)/.test(app) && /'acabProdutos', 'acabMov'\];/.test(app));
// Desde 28/10/2026 (Jones) a pasta é própria, de primeiro nível, logo acima
// de Estoques — que passou a se chamar "Estoque de materiais".
const _iniProd = html.indexOf('<div class="nav-group" data-group="estoque-confeccao">');
const nav = html.slice(_iniProd, html.indexOf('<div class="nav-group" data-group="estoques">'));
ok('31. a pasta é própria, fora de Estoques, na barra lateral', _iniProd > 0 && nav.length > 0
   && /<span class="nav-toggle">▼<\/span> Estoque de produtos<\/button>/.test(nav)
   && /data-group="estoques">\s*<button[^>]*><span class="nav-toggle">▼<\/span> Estoque de materiais<\/button>/.test(html)
   && !nav.includes('data-page="acab-inventario"')
   && !nav.includes('data-page="acab-erp"')
   && ['acab-estoque', 'acab-mov', 'acab-produtos', 'acab-dre', 'acab-cmp'].every(p => nav.includes(`data-page="${p}"`)));
ok('32. cada tela tem a sua seção e o painel', ['acab-estoque', 'acab-mov', 'acab-produtos', 'acab-dre', 'acab-cmp']
   .every(p => html.includes(`<section class="page hidden" data-page="${p}">`) && html.includes(`id="${p}-painel"`)));
ok('32a. o inventário fica no pé de Entradas e saídas, e o endereço antigo leva para lá',
   /<div id="acab-inventario-painel"><\/div>`;\s*renderInventario\(\);/.test(mod) && /if \(page === 'acab-inventario'\) page = 'acab-mov';/.test(app));
ok('33. o script e o modal estão no index.html', /<script src="estoque-confeccao\.js\?v=/.test(html) && /id="modal-acab-produto"/.test(html));
ok('34. os botões de editar aparecem para quem tem a área', /body\.pode-estoque-confeccao \[data-page\^="acab-"\] \.admin-only/.test(css));
const onclicks = [...new Set((mod.match(/on(?:click|change|input)="(acab\w+)\(/g) || []).map(x => x.match(/"(acab\w+)/)[1]))];
ok('35. todo handler das telas existe no window', onclicks.length > 10 && onclicks.every(n => typeof w[n] === 'function'), onclicks.filter(n => typeof w[n] !== 'function'));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
