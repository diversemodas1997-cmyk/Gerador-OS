/* Rode com:  node testes/estoque-confeccao-manual.js

   ENTRADA E SAÍDA À MÃO no Estoque confecção (21/10/2026, Junior: "insira
   entrada/saída manual no estoque confecção"). O que este teste protege:
     · a tela de Entradas e saídas mostra os formulários mesmo sem nenhum SKU
       cadastrado (o estoque começou do zero e a tela só dizia "cadastre");
     · a entrada de um SKU que ainda não existe cadastra o SKU (o do catálogo
       com linha, cor e descrição) e lança a entrada, gravando as duas chaves;
     · SKU fora do catálogo pede confirmação; tamanho fora da linha é barrado;
     · a saída exige SKU cadastrado e estoque no tamanho;
     · o "+ entrada"/"− saída" do Estoque atual abre a tela com o SKU e o
       tamanho já escolhidos.
   Carrega o estoque-confeccao.js de verdade, com uma página de mentira. */
const fs = require('fs');
const path = require('path');

const mod = fs.readFileSync(path.join(__dirname, '..', 'estoque-confeccao.js'), 'utf8');

// Uma página de mentira: cada id é um elemento com value/innerHTML. O
// innerHTML de um painel cria os campos que ele declara (id="..." value="...").
function pagina() {
  const els = {};
  const el = id => (els[id] = els[id] || {
    id, value: '', _html: '', focus() { pagina.foco = id; }, scrollIntoView() {}, setSelectionRange() {},
    get innerHTML() { return this._html; },
    set innerHTML(h) {
      this._html = h;
      for (const m of h.matchAll(/<(input|select)[^>]*\bid="([^"]+)"[^>]*>/g)) {
        const e = el(m[2]);
        const v = m[0].match(/\bvalue="([^"]*)"/);
        e.value = m[1] === 'input' ? (v ? v[1] : '') : '';
      }
      // A caixa de tamanho: vale a opção marcada, ou a primeira.
      if (/<option/.test(h) && !/<(input|select)/.test(h)) {
        const sel = h.match(/<option value="([^"]*)" selected/) || h.match(/<option value="([^"]*)"/);
        this.value = sel ? sel[1] : '';
      }
    },
    get textContent() { return this._html; }, set textContent(t) { this._html = t; }
  });
  return { getElementById: el, querySelector: () => null, els };
}

function carrega(STATE, catalogo, respostaConfirm) {
  const window = {};
  const doc = pagina();
  const gravou = [], toasts = [], foi = [];
  let n = 0;
  Object.assign(globalThis, {
    document: doc,
    exigirEdicao: () => true,
    toast: (t, k) => toasts.push((k || '') + ':' + t),
    saveState: async k => { gravou.push(k); },
    desfazerNomearAcao: () => {},
    confirm: () => respostaConfirm,
    goto: p => { foi.push(p); window.renderEstoqueConfeccao(p); }
  });
  new Function('window', 'STATE', 'formatDate', 'esc', 'uid', 'totaisPorTamanhoTomOS', 'skusDaOS', '_statusOS',
    'catalogoSkus', 'currentRole', mod)(
    window, STATE, iso => iso, s => String(s == null ? '' : s), () => 'id_' + (++n),
    () => ({ colTotal: () => 0 }), () => [], () => '', catalogo || [], 'admin');
  doc.getElementById('acab-mov-painel');
  doc.getElementById('acab-estoque-painel');
  return { window, doc, gravou, toasts, foi };
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const estado = () => ({ meta: { acabConfig: { inicio: '2026-10-21T12:00:00.000Z' } }, acabProdutos: [], acabMov: [], ordens: [] });
const CAT = [{ item: 'CM.LISA-PRE', tipo: 'CM.LISA', cor: 'PRE', descricao: 'CAMISETA BÁSICA PRETO' }];

(async () => {
  console.log('-- sem nenhum SKU cadastrado --');
  let STATE = estado();
  let t = carrega(STATE, CAT, true);
  t.window.renderEstoqueConfeccao('acab-mov');
  const painel = t.doc.els['acab-mov-painel'].innerHTML;
  ok('1. a tela mostra os dois formulários', /Registrar entrada/.test(painel) && /Registrar saída/.test(painel));
  ok('2. e a lista de SKUs sugere os do catálogo', /<datalist id="acab-skus">[\s\S]*value="CM.LISA-PRE"/.test(painel));

  const campo = (id, v) => { t.doc.getElementById(id).value = v; };
  campo('acab-ent-produto', 'cm.lisa-pre');
  t.window.acabTrocarProduto('acab-ent');
  ok('3. SKU do catálogo: a caixa de tamanho traz os da linha', /value="P"/.test(t.doc.els['acab-ent-tam'].innerHTML) && /value="G3"/.test(t.doc.els['acab-ent-tam'].innerHTML));
  ok('4. e avisa que o SKU é novo', /SKU novo \(do catálogo/.test(t.doc.els['acab-ent-info'].innerHTML), t.doc.els['acab-ent-info'].innerHTML);
  campo('acab-ent-tam', 'M'); campo('acab-ent-qtd', '30'); campo('acab-ent-custo', '12,50');
  await t.window.acabRegistrarEntrada();
  const p = STATE.acabProdutos[0];
  ok('5. a entrada cadastra o SKU com linha, cor e descrição do catálogo', STATE.acabProdutos.length === 1 && p.sku === 'CM.LISA-PRE' && p.linha === 'CM.LISA' && p.cor === 'PRE' && p.descricao === 'CAMISETA BÁSICA PRETO', p);
  ok('6. e lança a entrada no SKU novo', STATE.acabMov.length === 1 && STATE.acabMov[0].produtoId === p.id && STATE.acabMov[0].qtd === 30 && STATE.acabMov[0].custo === 12.5 && STATE.acabMov[0].tamanho === 'M' && STATE.acabMov[0].origem === 'manual', STATE.acabMov);
  ok('7. grava os produtos e os lançamentos', t.gravou.join(',') === 'acabProdutos,acabMov', t.gravou);

  console.log('\n-- saída --');
  t.gravou.length = 0;
  campo('acab-sai-produto', 'CM.LISA-PRE'); t.window.acabTrocarProduto('acab-sai');
  campo('acab-sai-tam', 'M'); campo('acab-sai-qtd', '31'); campo('acab-sai-destino', 'cliente');
  await t.window.acabRegistrarSaida();
  ok('8. não sai mais do que há', STATE.acabMov.length === 1 && /insuficiente/.test(t.toasts.pop()));
  campo('acab-sai-qtd', '10');
  await t.window.acabRegistrarSaida();
  ok('9. a saída dentro do estoque é lançada', STATE.acabMov.length === 2 && STATE.acabMov[1].tipo === 'saida' && STATE.acabMov[1].qtd === 10 && t.gravou.join(',') === 'acabMov', STATE.acabMov);
  campo('acab-sai-produto', 'CM.LISA-AZU'); t.window.acabTrocarProduto('acab-sai');
  campo('acab-sai-qtd', '1');
  await t.window.acabRegistrarSaida();
  ok('10. saída de SKU não cadastrado é barrada', STATE.acabMov.length === 2 && /não está cadastrado/.test(t.toasts.pop()));

  console.log('\n-- SKU fora do catálogo --');
  STATE = estado();
  t = carrega(STATE, CAT, false);
  t.window.renderEstoqueConfeccao('acab-mov');
  campo('acab-ent-produto', 'BM.MOL-CIN'); t.window.acabTrocarProduto('acab-ent');
  campo('acab-ent-tam', 'G'); campo('acab-ent-qtd', '5');
  await t.window.acabRegistrarEntrada();
  ok('11. pede confirmação, e sem ela não cadastra nem lança', STATE.acabProdutos.length === 0 && STATE.acabMov.length === 0);
  t = carrega(STATE, CAT, true);
  t.window.renderEstoqueConfeccao('acab-mov');
  campo('acab-ent-produto', 'BM.MOL-CIN'); t.window.acabTrocarProduto('acab-ent');
  campo('acab-ent-tam', 'G'); campo('acab-ent-qtd', '5');
  await t.window.acabRegistrarEntrada();
  ok('12. confirmado, cadastra pela linha e cor do SKU', STATE.acabProdutos.length === 1 && STATE.acabProdutos[0].linha === 'BM.MOL' && STATE.acabProdutos[0].cor === 'CIN' && STATE.acabMov.length === 1, STATE.acabProdutos);
  campo('acab-ent-produto', 'CJ.INF-PRE'); t.window.acabTrocarProduto('acab-ent');
  campo('acab-ent-tam', 'GG'); campo('acab-ent-qtd', '5');
  await t.window.acabRegistrarEntrada();
  ok('13. tamanho fora da linha (adulto no infantil) é barrado', STATE.acabProdutos.length === 1 && /não é da linha/.test(t.toasts.pop()));
  campo('acab-ent-produto', 'CM LISA'); campo('acab-ent-tam', 'G');
  await t.window.acabRegistrarEntrada();
  ok('14. SKU com espaço é recusado', STATE.acabProdutos.length === 1 && /SKU inválido/.test(t.toasts.pop()));

  console.log('\n-- do Estoque atual --');
  t.window.renderEstoqueConfeccao('acab-estoque');
  const est = t.doc.els['acab-estoque-painel'].innerHTML;
  const pid = STATE.acabProdutos[0].id;
  ok('15. cada linha tem "+ entrada" e "− saída"', est.includes(`acabLancar('ent','${pid}','G')`) && est.includes(`acabLancar('sai','${pid}','G')`));
  t.window.acabLancar('sai', pid, 'G');
  ok('16. "− saída" abre Entradas e saídas com o SKU e o tamanho escolhidos', t.foi.pop() === 'acab-mov'
     && t.doc.els['acab-sai-produto'].value === 'BM.MOL-CIN' && t.doc.els['acab-sai-tam'].value === 'G' && pagina.foco === 'acab-sai-qtd',
     [t.doc.els['acab-sai-produto'].value, t.doc.els['acab-sai-tam'].value, pagina.foco]);
  ok('17. e mostra quanto há no tamanho', /em estoque no G: <b>5<\/b>/.test(t.doc.els['acab-sai-info'].innerHTML), t.doc.els['acab-sai-info'].innerHTML);

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
