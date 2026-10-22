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

function carrega(STATE, catalogo, respostaConfirm, daOS) {
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
    (daOS && daOS.totais) || (() => ({ colTotal: () => 0 })), (daOS && daOS.skus) || (() => []), () => '', catalogo || [], 'admin');
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

  console.log('\n-- a entrada à mão abate as OS em Estoque (22/10/2026) --');
  {
    const st = estado();
    st.acabProdutos = [{ id: 'pp', sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', tamanhos: ['P', 'M', 'G'] }];
    // Duas OS em Estoque no M: a 0601 entrou por último (22/10), a 0598 antes (21/10).
    const osE = (id, num, quando, m) => ({ id, os: num, etapas: ['Corte', 'Estoque'], _m: m,
      progresso: { etapasCheck: { Estoque: true }, etapasSeq: { Estoque: Date.parse(quando) } } });
    st.ordens = [osE('o98', '0598', '2026-10-21T15:00:00Z', 10), osE('o01', '0601', '2026-10-22T10:00:00Z', 8)];
    const daOS = { totais: o => ({ colTotal: k => (k === 'm' ? o._m : 0) }), skus: () => ['CM.LISA-PRE'] };
    const u = carrega(st, CAT, true, daOS);
    const c = (id, v) => { u.doc.getElementById(id).value = v; };
    u.window.renderEstoqueConfeccao('acab-mov');
    c('acab-ent-produto', 'CM.LISA-PRE'); u.window.acabTrocarProduto('acab-ent');
    c('acab-ent-tam', 'M'); c('acab-ent-qtd', '12'); u.window.acabInfoSku('acab-ent');
    ok('18. a tela avisa o que vai abater, da OS mais recente para a mais antiga',
       /Vai abater das OS em Estoque: <b>OS 0601 \(8\), OS 0598 \(4\)<\/b>/.test(u.doc.els['acab-ent-info'].innerHTML), u.doc.els['acab-ent-info'].innerHTML);
    await u.window.acabRegistrarEntrada();
    const m = st.acabMov[0];
    ok('19. a entrada grava o abatimento', JSON.stringify(m.abate) === JSON.stringify([{ osId: 'o01', os: '0601', qtd: 8 }, { osId: 'o98', os: '0598', qtd: 4 }]), m.abate);
    u.window.renderEstoqueConfeccao('acab-estoque');
    ok('20. o saldo não conta duas vezes: 18 das OS + 12 à mão − 12 abatidas = 18',
       /font-weight:700;">18</.test(u.doc.els['acab-estoque-painel'].innerHTML), u.doc.els['acab-estoque-painel'].innerHTML.match(/font-weight:700;[^<]*</g));
    u.window.renderEstoqueConfeccao('acab-mov');
    const pm = u.doc.els['acab-mov-painel'].innerHTML;
    ok('21. a lista mostra a OS abatida e a entrada que abateu', /de 10 · −4 à mão/.test(pm) && /abateu OS 0601 \(8\), OS 0598 \(4\)/.test(pm));
    c('acab-ent-produto', 'CM.LISA-PRE'); u.window.acabTrocarProduto('acab-ent');
    c('acab-ent-tam', 'M'); c('acab-ent-qtd', '10');
    await u.window.acabRegistrarEntrada();
    ok('22. o que passa do saldo das OS entra como peça nova (abate só as 6 que restavam)',
       JSON.stringify(st.acabMov[1].abate) === JSON.stringify([{ osId: 'o98', os: '0598', qtd: 6 }]), st.acabMov[1].abate);
    c('acab-ent-qtd', '3'); u.doc.getElementById('acab-ent-abate').checked = false;
    await u.window.acabRegistrarEntrada();
    ok('23. com a caixa desmarcada não abate', st.acabMov.length === 3 && !st.acabMov[2].abate, st.acabMov[2]);
    st.acabMov.splice(0, 1);
    u.window.renderEstoqueConfeccao('acab-mov');
    ok('24. apagar a entrada devolve as peças às OS', !/−4 à mão/.test(u.doc.els['acab-mov-painel'].innerHTML) && /de 10 · −6 à mão/.test(u.doc.els['acab-mov-painel'].innerHTML));
  }

  console.log('\n-- o botão dá o início da contagem (22/10/2026) --');
  {
    const st = estado();
    delete st.meta.acabConfig.inicio;
    st.acabProdutos = [{ id: 'pp', sku: 'CM.LISA-PRE', linha: 'CM.LISA', cor: 'PRE', tamanhos: ['P', 'M', 'G'] }];
    const u = carrega(st, CAT, true);
    const c = (id, v) => { u.doc.getElementById(id).value = v; };
    u.window.renderEstoqueConfeccao('acab-estoque');
    ok('25. abrir a tela não marca o início sozinho', !st.meta.acabConfig.inicio);
    ok('26. sem início, o Estoque atual mostra o botão', /acabIniciarContagem\(\)">▶ Iniciar contagem do estoque/.test(u.doc.els['acab-estoque-painel'].innerHTML));
    u.window.renderEstoqueConfeccao('acab-mov');
    c('acab-ent-produto', 'CM.LISA-PRE'); u.window.acabTrocarProduto('acab-ent');
    c('acab-ent-tam', 'M'); c('acab-ent-qtd', '5');
    await u.window.acabRegistrarEntrada();
    ok('27. sem início, a entrada à mão é barrada', !st.acabMov.length && /Inicie a contagem/.test(u.toasts.pop()));
    st.acabMov.push({ id: 'velho', tipo: 'entrada', origem: 'manual', produtoId: 'pp', tamanho: 'M', data: '2026-10-01', qtd: 7, criadoEm: '2026-10-01T10:00:00.000Z' });
    await u.window.acabIniciarContagem();
    ok('28. o botão grava o início e quem deu', !!st.meta.acabConfig.inicio && st.meta.acabConfig.inicioPor === 'admin', st.meta.acabConfig);
    c('acab-ent-produto', 'CM.LISA-PRE'); u.window.acabTrocarProduto('acab-ent');
    c('acab-ent-tam', 'M'); c('acab-ent-qtd', '5');
    await u.window.acabRegistrarEntrada();
    u.window.renderEstoqueConfeccao('acab-estoque');
    const pe = u.doc.els['acab-estoque-painel'].innerHTML;
    ok('29. o lançamento de antes do início não conta (só os 5 de agora)', /font-weight:700;">5</.test(pe) && st.acabMov.length === 2, pe.match(/font-weight:700;[^<]*</g));
    ok('30. depois de iniciada, a tela mostra quando e oferece reiniciar', /Contagem iniciada em/.test(pe) && /Reiniciar contagem/.test(pe));
  }

  console.log('\n-- os SKUs do programa entram sozinhos (22/10/2026) --');
  {
    const st = estado();
    st.desenhos = [{ id: 'd1', desc: 'Camiseta lisa preta', skuLinha: 'CM.LISA-PRE' }, { id: 'd2', desc: 'Calça', skuLinha: 'CJM' }];
    st.ordens = [{ id: 'o1', os: '0700', modeloNome: 'Bermuda moletom', etapas: [] }];
    const daOS = { skus: o => (o.id === 'o1' ? ['BM.LISA-BEGE'] : []) };
    globalThis.podeGravar = () => true;
    const u = carrega(st, CAT, true, daOS);
    u.window.renderEstoqueConfeccao('acab-mov');
    await new Promise(r => setTimeout(r, 0));
    const skus = st.acabProdutos.map(p => p.sku).sort().join(',');
    ok('31. abrir a tela cadastra os SKUs dos desenhos e das OS (a linha sozinha não)', skus === 'BM.LISA-BEGE,CM.LISA-PRE', skus);
    const pre = st.acabProdutos.find(p => p.sku === 'CM.LISA-PRE');
    ok('32. com a descrição do catálogo, linha, cor e tamanhos', pre.descricao === 'CAMISETA BÁSICA PRETO' && pre.linha === 'CM.LISA' && pre.cor === 'PRE' && pre.tamanhos.includes('G3') && pre.origem === 'programa', pre);
    ok('33. e o aviso "Nenhum SKU cadastrado" some', !/Nenhum SKU cadastrado/.test(u.doc.els['acab-mov-painel'].innerHTML));
    const antes = u.gravou.length;
    u.window.renderEstoqueConfeccao('acab-mov');
    await new Promise(r => setTimeout(r, 0));
    ok('34. sem SKU novo, não grava de novo', u.gravou.length === antes && st.acabProdutos.length === 2);
    delete globalThis.podeGravar;
  }

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
