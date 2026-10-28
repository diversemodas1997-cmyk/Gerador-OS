/* Rode com:  node testes/estoque-limites-linha.js

   O MÍNIMO E O MÁXIMO DE CADA LINHA DO ESTOQUE ATUAL (28/10/2026, Jones:
   "Insira coluna configurável para estoque mínimo e outra para estoque máximo
   no quadro Estoque atual. Retire o quadro de configuração de estoque mínimo e
   máximo geral. Os estoques mínimos e máximos devem ser configurados para cada
   linha do Estoque atual, de acordo com o sku, cor e tamanho"). Protege:
     · a tabela tem as colunas Mín. e Máx., editáveis, por SKU e tamanho;
     · o limite de uma linha não vale para outro tamanho nem outro SKU;
     · sem limite, a linha fica "Sem limite" (o par geral antigo não vale mais);
     · gravar o limite de um SKU lido do programa dá registro próprio a ele;
     · máximo menor que o mínimo é recusado; apagar o número tira o limite;
     · Produtos não tem mais os campos gerais de mínimo e máximo.
   Carrega o estoque-confeccao.js de verdade, com uma página de mentira. */
const fs = require('fs');
const path = require('path');
const mod = fs.readFileSync(path.join(__dirname, '..', 'estoque-confeccao.js'), 'utf8');

function carrega(STATE) {
  const window = {};
  const els = {};
  const el = id => (els[id] = els[id] || { id, value: '', innerHTML: '' });
  const gravou = [], toasts = [];
  Object.assign(globalThis, {
    document: { getElementById: el, querySelector: () => null, querySelectorAll: () => [] },
    exigirEdicao: () => true, temAcesso: () => true,
    toast: (t, k) => toasts.push((k || '') + ':' + t),
    saveState: async k => { gravou.push(k); }, desfazerNomearAcao: () => {}, confirm: () => true, goto: () => {}
  });
  new Function('window', 'STATE', 'formatDate', 'esc', 'uid', 'totaisPorTamanhoTomOS', 'skusDaOS', '_statusOS',
    'catalogoSkus', 'currentRole', mod)(
    window, STATE, iso => iso, s => String(s == null ? '' : s), () => 'id',
    () => ({ colTotal: () => 0 }), () => [], () => '', [], 'admin');
  return { window, els, gravou, toasts, I: window._acabInterno };
}
// O campo editado: el.closest('tr') devolve a linha com a célula de status.
const campo = (valor) => {
  const st = { innerHTML: '' };
  return { value: valor, st, closest: () => ({ querySelector: () => st }) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

(async () => {
  const STATE = {
    meta: { acabConfig: { inicio: '2026-10-22T00:00:00.000Z', estoqueMin: 150, estoqueMax: 1000 } },
    acabProdutos: [{ id: 'p1', sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', tamanhos: ['P', 'G'], preco: 0 }],
    acabMov: [
      { id: 'm1', tipo: 'entrada', produtoId: 'p1', tamanho: 'P', qtd: 40, data: '2026-10-28', criadoEm: '2026-10-28T10:00:00.000Z' },
      { id: 'm2', tipo: 'entrada', produtoId: 'p1', tamanho: 'G', qtd: 300, data: '2026-10-28', criadoEm: '2026-10-28T10:00:00.000Z' }
    ],
    ordens: [], cores: [], modelos: [],
    desenhos: [{ id: 'd1', desc: 'Camiseta Básica | Branco', skuLinha: 'CM.LISA-BRA' }]
  };
  const u = carrega(STATE);
  const { I } = u;

  u.window.renderEstoqueConfeccao('acab-estoque');
  let html = u.els['acab-estoque-painel'].innerHTML;
  ok('1. a tabela tem as colunas Min. e Max.', /<th[^>]*>Mín\.<\/th><th[^>]*>Máx\.<\/th>/.test(html));
  ok('2. cada linha tem os dois campos editaveis', (html.match(/acabSalvarLimite\('p1','P','min'/g) || []).length === 1
    && (html.match(/acabSalvarLimite\('p1','G','max'/g) || []).length === 1);
  ok('3. sem limite, a linha fica "Sem limite" (o 150/1000 geral nao vale)', !/Estoque mínimo<\/span>/.test(html) && /Sem limite/.test(html));

  let c = campo('50');
  await u.window.acabSalvarLimite('p1', 'P', 'min', c);
  ok('4. o minimo grava no cadastro do SKU, por tamanho', JSON.stringify(STATE.acabProdutos[0].limites) === '{"P":{"min":50,"max":null}}'
    && u.gravou.includes('acabProdutos'), STATE.acabProdutos[0].limites);
  ok('5. e a linha passa a Estoque minimo (40 <= 50)', /Estoque mínimo/.test(c.st.innerHTML), c.st.innerHTML);
  ok('6. o limite do P nao vale para o G', I.limiteDe(STATE.acabProdutos[0], 'G').min === null);

  c = campo('200');
  await u.window.acabSalvarLimite('p1', 'G', 'max', c);
  ok('7. o maximo marca Excedente (300 > 200)', /Excedente/.test(c.st.innerHTML), c.st.innerHTML);
  c = campo('30');
  await u.window.acabSalvarLimite('p1', 'P', 'max', c);
  ok('8. maximo menor que o minimo e recusado', I.limiteDe(STATE.acabProdutos[0], 'P').max === null && /não pode ser menor/.test(u.toasts.pop()));
  c = campo('');
  await u.window.acabSalvarLimite('p1', 'G', 'max', c);
  ok('9. apagar o numero tira o limite', !('G' in STATE.acabProdutos[0].limites) && /Sem limite/.test(c.st.innerHTML), STATE.acabProdutos[0].limites);

  await Promise.resolve();
  c = campo('10');
  await u.window.acabSalvarLimite('sku:CM.LISA-BRA', 'M', 'min', c);
  const bra = STATE.acabProdutos.find(p => p.sku === 'CM.LISA-BRA');
  ok('10. o SKU do programa ganha registro proprio com o limite', !!bra && bra.id === 'sku:CM.LISA-BRA' && bra.limites.M.min === 10, STATE.acabProdutos);
  ok('11. outro SKU nao herda o limite', I.limiteDe(STATE.acabProdutos[0], 'M').min === null);

  u.window.renderEstoqueConfeccao('acab-produtos');
  html = u.els['acab-produtos-painel'].innerHTML;
  ok('12. Produtos nao tem mais o minimo e o maximo gerais', !/acab-cfg-min|acab-cfg-max/.test(html) && /acab-cfg-inicio/.test(html));
  ok('13. o filtro de status tem "Sem limite"', /value="sem"/.test(u.els['acab-estoque-painel'].innerHTML));

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
