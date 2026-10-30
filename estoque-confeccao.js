/* =====================================================================
   ESTOQUE CONFECÇÃO — o produto acabado, dentro do Gerador-OS (21/10/2026)
   =====================================================================
   (Junior: "Analise uma forma de copiar o programa Estoque confecção e insira
   os campos do programa que não forem redundantes ao programa gerador OS, de
   forma que seja uma pasta na barra lateral, dentro da pasta Estoque".)

   O original (Desktop\Estoque-Confeccao, servidor Node no Render) controla a
   prateleira do PRODUTO ACABADO: SKU × tamanho, com entradas, saídas (venda,
   cancelamento, devolução, defeito), contagem de inventário, custo médio
   ponderado e a importação das vendas do ERP. Decidido com o Junior: o
   Gerador-OS passa a valer e o estoque COMEÇA DO ZERO (sem o histórico do
   Render); a lista de SKUs pode ser trazida do catálogo que o Render publicou.

   O QUE VEIO (uma tela para cada, na pasta "Estoque confecção", hoje "Estoque de produtos"):
     · Estoque atual       — a tabela SKU × tamanho, os totais, e o mínimo e o
                             máximo de cada linha;
     · Entradas e saídas   — os lançamentos à mão e a lista dos últimos;
     · Inventário          — a contagem física que acerta o saldo (desde
                             28/10/2026, a parte de baixo de Entradas e saídas);
     · Produtos (SKUs)     — o cadastro e o início do estoque;
     · Saídas do ERP       — o relatório de vendas (PDF ou CSV) vira saída;
     · Relatório CMP       — o custo médio ponderado, movimento a movimento.

   O QUE NÃO VEIO, porque o Gerador-OS já tem:
     · login e usuários         → as contas do programa e as concessões
                                  (área "Estoque de produtos" em Configurações);
     · backup, restaurar, zerar → Dados, backup e pastas (e o Desfazer);
     · a sincronização com o Gerador-OS (o Render lia a etapa "Estoque" das
       OS pela nuvem a cada minuto) → aqui a entrada da OS é lida DIRETO da
       própria OS (ver _entradasDasOS): marcar a etapa Estoque lança, desmarcar
       estorna, sem cópia nem atraso;
     · a publicação para a Contabilidade → vai junto do contabSnapshot que o
       Gerador-OS já grava (produtosAcabados, no mesmo formato do Render).

   ONDE MORA: os SKUs são os do programa (desenhos, modelos e OS, lidos na
   hora) mais STATE.acabProdutos (o que só o estoque tem), e STATE.acabMov (entradas, saídas e
   ajustes lançados aqui ou pelo ERP), mais STATE.meta.acabConfig (mínimo,
   máximo, início). As entradas das OS NÃO são gravadas: são calculadas da OS a
   cada desenho, e por isso não há como divergirem dela.

   Arquivo separado do app.js, como o colorimetro.js. As telas são as
   section[data-page^="acab-"] do index.html; o goto chama renderEstoqueConfeccao.
   ===================================================================== */
(function () {
  'use strict';

  // Os tamanhos na ordem da prateleira: adulto (com o PP, que só a bermuda
  // feminina usa) e o infantil, que o Render não tinha e as OS têm.
  const TAMS = ['PP', 'P', 'M', 'G', 'GG', 'G1', 'G2', 'G3', '2', '4', '6', '8', '10', '12', '14', '16'];
  const TAMS_PADRAO = ['P', 'M', 'G', 'GG', 'G1', 'G2', 'G3'];
  const TAMS_INFANTIL = ['2', '4', '6', '8', '10', '12', '14', '16'];
  const TAM_DA_GRADE = { p: 'P', m: 'M', g: 'G', gg: 'GG', g1: 'G1', g2: 'G2', g3: 'G3',
    t2: '2', t4: '4', t6: '6', t8: '8', t10: '10', t12: '12', t14: '14', t16: '16' };
  const DESTINOS = [
    { k: 'cliente', rotulo: 'Estoque → Cliente' },
    { k: 'cancelamento', rotulo: 'Estoque → Cancelamento' },
    { k: 'devolucao', rotulo: 'Estoque → Devolução' },
    { k: 'defeitos', rotulo: 'Estoque → Defeitos' }
  ];
  const ETAPA_ESTOQUE_RE = /estoque/i;

  /* ---------------------------------------------------------------- base */
  const lista = k => (Array.isArray(STATE[k]) ? STATE[k] : (STATE[k] = []));
  /* OS SKUs SÃO LIDOS DO PROGRAMA, NÃO COPIADOS (23/10/2026, Junior: "não deve
     ser um programa à parte do Gerador OS, mas receber a origem das
     informações do mesmo lugar que todos os campos do programa recebem, para
     não precisar ser sincronizado"). Antes, os SKUs dos desenhos e das OS eram
     COPIADOS para STATE.acabProdutos ao abrir a tela (sincronizarSkusPrograma)
     e, entre uma cópia e outra, o estoque não via o SKU novo.
     Agora a lista é montada na hora: o que o programa define (skusDoPrograma:
     desenho técnico, modelo e OS) mais o que só existe aqui (`cadastro()`:
     SKU do catálogo, do ERP, digitado na entrada, e o código do ERP, o preço e
     os tamanhos que alguém acertou à mão). O SKU do programa tem id fixo
     ('sku:' + SKU); ele só é gravado no cadastro quando ganha algo próprio
     daqui — um lançamento, um código, um preço (ver fixarProduto). */
  const cadastro = () => lista('acabProdutos');
  // A lista é refeita uma vez por rodada de código (não a cada busca de SKU).
  let _produtosCache = null;
  const esquecerProdutos = () => { _produtosCache = null; };
  function produtos() {
    if (_produtosCache) return _produtosCache;
    const cad = cadastro();
    const tem = new Set(cad.map(p => norm(p.sku)));
    const virtuais = [];
    skusDoPrograma().forEach((desc, sku) => {
      if (tem.has(sku)) return;
      const c = doCatalogo(sku);
      const pt = partesSku(sku);
      virtuais.push({ id: 'sku:' + sku, codigo: null, sku, linha: norm(c && c.tipo) || pt.linha, cor: norm(c && c.cor) || pt.cor,
        descricao: (c && c.descricao) || desc || '', tamanhos: tamsDaLinha((c && c.tipo) || pt.linha), preco: 0,
        origem: 'programa', doPrograma: true });
    });
    _produtosCache = cad.concat(virtuais);
    Promise.resolve().then(esquecerProdutos);
    return _produtosCache;
  }
  // O SKU do programa passa a ter registro próprio aqui (mesmo id), para
  // guardar o que só o estoque sabe e não sumir se o desenho trocar de SKU.
  function fixarProduto(p) {
    if (!p || !p.doPrograma || cadastro().some(x => x.id === p.id)) return false;
    const fixo = Object.assign({}, p, { criadoEm: new Date().toISOString(), criadoPor: quemSou() });
    delete fixo.doPrograma;
    cadastro().push(fixo);
    esquecerProdutos();
    return fixo;
  }
  const movs = () => lista('acabMov');
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase();
  const rankTam = t => { const i = TAMS.indexOf(t); return i < 0 ? 99 : i; };
  const ordenarTams = arr => [...new Set(arr)].sort((a, b) => rankTam(a) - rankTam(b));
  const asTam = v => { const t = norm(v); return TAMS.includes(t) ? t : ''; };
  const quemSou = () => (typeof _cpQuemSou === 'function' ? _cpQuemSou() : '');
  const hoje = () => (typeof _aviHoje === 'function' ? _aviHoje() : new Date().toISOString().slice(0, 10));
  const fmtN = v => Math.round(Number(v) || 0).toLocaleString('pt-BR');
  const fmtR = v => (Number(v) || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  const fmtD = iso => (iso ? formatDate(String(iso).slice(0, 10)) : '—');
  const num = v => { const n = parseFloat(String(v == null ? '' : v).replace(/\./g, '').replace(',', '.')); return Number.isFinite(n) ? n : NaN; };
  const val = id => ((document.getElementById(id) || {}).value || '');
  const mono = "text-align:right;font-family:'IBM Plex Mono',monospace;white-space:nowrap;";

  function cfg() {
    STATE.meta = STATE.meta || {};
    const c = STATE.meta.acabConfig && typeof STATE.meta.acabConfig === 'object' ? STATE.meta.acabConfig : {};
    return {
      inicio: c.inicio || '',
      inicioPor: c.inicioPor || '',
      osDispensadas: Array.isArray(c.osDispensadas) ? c.osDispensadas : []
    };
  }
  async function gravarCfg(mudar) {
    STATE.meta = STATE.meta || {};
    STATE.meta.acabConfig = Object.assign({}, STATE.meta.acabConfig || {}, mudar);
    await saveState('meta');
  }

  /* O INÍCIO DO ESTOQUE. "Começar do zero" quer dizer que as OS que já tinham
     entrado no estoque antes daqui não lançam entrada — essas peças o Render já
     contou, e muitas já foram vendidas. Sem início, nenhuma OS lança: é o lado
     seguro.
     O BOTÃO DÁ O INÍCIO (22/10/2026, Junior: "insira botão para dar início na
     contagem do estoque, pois o abate pelas OS devem ter início
     sincronizado"). Antes o início era carimbado sozinho na primeira vez em
     que alguém abria a tela — a hora era um acaso. Agora é o botão "Iniciar
     contagem do estoque", no Estoque atual, e ele vale para os DOIS lados ao
     mesmo tempo: as OS que entram em Estoque e os lançamentos daqui (entradas,
     saídas, ajustes, ERP) só contam do início em diante. Assim a entrada à mão
     e a OS que ela abate (planoAbateOS) começam juntas. O lançamento de antes
     não é apagado: fica guardado e volta a contar se o início for recuado em
     Produtos. Sem início, nada pode ser lançado (exigirInicio). */
  function exigirInicio() {
    if (cfg().inicio) return true;
    toast('Inicie a contagem do estoque primeiro (botão no Estoque atual)', 'err');
    return false;
  }
  const fmtDH = iso => new Date(iso).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
  // Os lançamentos que contam: os do início em diante (ver exigirInicio).
  function movsValidos() {
    const ini = cfg().inicio;
    return ini ? movs().filter(m => !m.criadoEm || String(m.criadoEm) >= ini) : movs();
  }

  const produtoPorId = id => produtos().find(p => p.id === id);
  const produtoPorSku = sku => { const s = norm(sku); return s ? produtos().find(p => norm(p.sku) === s) : null; };
  const produtoPorCodigo = c => (Number(c) > 0 ? produtos().find(p => Number(p.codigo) === Number(c)) : null);
  const rotuloProduto = p => (p ? (p.codigo ? String(p.codigo).padStart(3, '0') + ' · ' : '') + (p.sku || '?') : '(SKU apagado)');

  /* ------------------------------------------- as entradas que vêm das OS */
  /* A etapa terminal "Estoque" marcada na OS é a peça chegando à prateleira —
     o mesmo gatilho que o Render usava. A entrada é uma por tamanho, com a
     quantidade da coluna "Total por tamanho" da folha (totaisPorTamanhoTomOS),
     e o SKU é o da OS (skusDaOS: Linha do modelo/desenho + Sigla da cor).
     A data é a do carimbo da etapa; OS marcada antes do início fica de fora.
     O que não dá para lançar sozinho (sem SKU, mais de uma cor, SKU não
     cadastrado) vai para a lista "OS esperando SKU" da tela Estoque atual. */
  function _entradasDasOS() {
    const c = cfg();
    const inicio = c.inicio ? Date.parse(c.inicio) : NaN;
    const entradas = [], pendencias = [];
    if (isNaN(inicio)) return { entradas, pendencias, semInicio: true };
    const dispensadas = new Set(c.osDispensadas);
    (STATE.ordens || []).forEach(o => {
      if (typeof _statusOS === 'function' && _statusOS(o) === 'cancelado') return;
      const checks = (o.progresso && o.progresso.etapasCheck) || {};
      const seqs = (o.progresso && o.progresso.etapasSeq) || {};
      const etapa = (o.etapas || []).find(n => ETAPA_ESTOQUE_RE.test(n) && !!checks[n]);
      if (!etapa) return;
      const quando = Number(seqs[etapa]);
      if (!(quando >= inicio)) return;
      if (dispensadas.has(o.id)) return;
      const data = new Date(quando);
      const dataIso = data.getFullYear() + '-' + String(data.getMonth() + 1).padStart(2, '0') + '-' + String(data.getDate()).padStart(2, '0');
      const TT = totaisPorTamanhoTomOS(o);
      const porTam = [];
      Object.keys(TAM_DA_GRADE).forEach(k => {
        const q = Math.round(TT.colTotal(k) || 0);
        if (q > 0) porTam.push([TAM_DA_GRADE[k], q]);
      });
      const pecas = porTam.reduce((s, x) => s + x[1], 0);
      if (!pecas) return;
      const skus = typeof skusDaOS === 'function' ? skusDaOS(o) : [];
      const base = { osId: o.id, os: o.os || '—', modelo: o.modeloNome || '', pecas, data: dataIso, em: data.toISOString(), skus };
      if (skus.length !== 1) {
        pendencias.push(Object.assign(base, { motivo: skus.length ? 'mais de uma cor (' + skus.join(', ') + ') — lance à mão' : 'a OS não tem SKU (Linha SKU do modelo/desenho e Sigla SKU da cor)' }));
        return;
      }
      const p = produtoPorSku(skus[0]);
      if (!p) {
        pendencias.push(Object.assign(base, { motivo: 'o SKU ' + skus[0] + ' não está cadastrado', cadastrar: skus[0], tams: porTam.map(x => x[0]) }));
        return;
      }
      porTam.forEach(([tam, q]) => entradas.push({
        id: 'os:' + o.id + ':' + tam, tipo: 'entrada', origem: 'os', produtoId: p.id, tamanho: tam,
        data: dataIso, qtd: q, custo: 0, osId: o.id, os: o.os || '', sku: p.sku, criadoEm: data.toISOString()
      }));
    });
    // O que as entradas à mão já abateram (ver planoAbateOS) sai da OS.
    const abatido = new Map();
    movsValidos().forEach(m => (m.tipo === 'entrada' && Array.isArray(m.abate) ? m.abate : []).forEach(a => {
      const k = a.osId + '|' + m.produtoId + '|' + m.tamanho;
      abatido.set(k, (abatido.get(k) || 0) + (Number(a.qtd) || 0));
    }));
    entradas.forEach(e => {
      const ab = Math.min(e.qtd, abatido.get(e.osId + '|' + e.produtoId + '|' + e.tamanho) || 0);
      if (ab > 0) { e.qtdOS = e.qtd; e.abatido = ab; e.qtd -= ab; }
    });
    return { entradas, pendencias, semInicio: false };
  }

  /* A ENTRADA À MÃO ABATE AS OS EM ESTOQUE (22/10/2026, Junior: "as entradas
     manuais em estoque confecção abaterem as quantidades produzidas por OS que
     estão com status Estoque, preferencialmente pela OS com data mais recente";
     a data é a de quando a OS entrou em Estoque). Sem isso, a peça que a OS já
     lançou e alguém lança de novo à mão contava duas vezes, e o único remédio
     era dispensar a OS inteira.
     Ao gravar, a entrada procura as OS do mesmo SKU e tamanho com saldo, da que
     entrou em Estoque por último para a mais antiga, e guarda em `abate`
     quanto tirou de cada uma ([{osId, os, qtd}]). O plano fica GRAVADO, não é
     refeito a cada tela: refeito, a OS que chegasse depois "roubaria" o
     abatimento das antigas. O que passar do saldo das OS entra como peça nova;
     apagar a entrada devolve as peças às OS. */
  function planoAbateOS(produtoId, tamanho, qtd, dasOS) {
    const livres = (dasOS || _entradasDasOS()).entradas
      .filter(e => e.produtoId === produtoId && e.tamanho === tamanho && e.qtd > 0)
      .sort((a, b) => String(b.criadoEm).localeCompare(String(a.criadoEm)));
    const plano = [];
    let falta = qtd;
    for (const e of livres) {
      if (falta <= 0) break;
      const q = Math.min(falta, e.qtd);
      plano.push({ osId: e.osId, os: e.os, qtd: q });
      falta -= q;
    }
    return plano;
  }
  const textoAbate = plano => plano.map(a => `OS ${a.os} (${fmtN(a.qtd)})`).join(', ');
  // A caixa "abater das OS em Estoque" vem marcada; só desmarcada não abate.
  const abaterMarcado = () => { const el = document.getElementById('acab-ent-abate'); return !(el && el.checked === false); };

  // Todos os movimentos: os lançados (STATE.acabMov) e os das OS.
  function todosMovs(dasOS) {
    return movsValidos().concat((dasOS || _entradasDasOS()).entradas);
  }

  /* -------------------------------------------------------- as contas */
  function saldos(ms) {
    const mapa = new Map();
    const pega = (pid, t) => {
      const k = pid + '|' + t;
      let e = mapa.get(k);
      if (!e) { e = { ent: 0, entOS: 0, sai: 0, aj: 0, ult: null }; mapa.set(k, e); }
      return e;
    };
    ms.forEach(m => {
      const e = pega(m.produtoId, m.tamanho);
      if (m.tipo === 'entrada') { e.ent += Number(m.qtd) || 0; if (m.origem === 'os') e.entOS += Number(m.qtd) || 0; }
      else if (m.tipo === 'saida') e.sai += Number(m.qtd) || 0;
      else if (m.tipo === 'ajuste') {
        e.aj += Number(m.diferenca) || 0;
        if (!e.ult || String(m.data) > String(e.ult.data) || (m.data === e.ult.data && String(m.criadoEm) > String(e.ult.criadoEm))) e.ult = m;
      }
    });
    const de = (pid, t) => { const e = mapa.get(pid + '|' + t) || { ent: 0, entOS: 0, sai: 0, aj: 0, ult: null }; return Object.assign({ total: e.ent - e.sai + e.aj }, e); };
    return { de, mapa };
  }

  /* CUSTO MÉDIO PONDERADO, como no original: a cada entrada a média é refeita
     entre o que havia e o que entrou; a saída sai pela média e não a muda. O
     ajuste de inventário não entra na conta de custo (também como no original).
     Sem nenhuma entrada com custo, vale o preço do cadastro. */
  function cmp(ms, produtoId, tamanho) {
    const casa = m => m.produtoId === produtoId && (tamanho == null || m.tamanho === tamanho);
    const evs = ms.filter(m => casa(m) && (m.tipo === 'entrada' || m.tipo === 'saida'))
      .sort((a, b) => String(a.data).localeCompare(String(b.data)) || (a.tipo === b.tipo ? 0 : a.tipo === 'entrada' ? -1 : 1)
        || String(a.criadoEm || '').localeCompare(String(b.criadoEm || '')));
    let saldo = 0, medio = 0, cmv = 0;
    const linhas = [];
    evs.forEach(m => {
      const q = Number(m.qtd) || 0;
      if (m.tipo === 'entrada') {
        const custo = Number(m.custo) || 0;
        const novo = saldo + q;
        medio = novo > 0 ? (saldo * medio + q * custo) / novo : 0;
        saldo = novo;
        linhas.push({ m, qtdEnt: q, custoEnt: custo, totalEnt: q * custo, saldo, medio });
      } else {
        const sai = Math.min(q, Math.max(0, saldo));
        cmv += sai * medio;
        saldo -= sai;
        linhas.push({ m, qtdSai: q, custoSai: medio, totalSai: sai * medio, saldo, medio });
      }
    });
    const p = produtoPorId(produtoId);
    const temCusto = evs.some(m => m.tipo === 'entrada' && Number(m.custo) > 0);
    const custoMedio = temCusto ? medio : (Number(p && p.preco) || 0);
    return { linhas, saldo, medio: custoMedio, valor: saldo * custoMedio, cmv };
  }

  // As linhas SKU × tamanho: os tamanhos do cadastro e os que já têm movimento.
  function linhasSkuTam(ms) {
    const extra = new Map();
    ms.forEach(m => { if (!extra.has(m.produtoId)) extra.set(m.produtoId, new Set()); extra.get(m.produtoId).add(m.tamanho); });
    const out = [];
    produtos().slice().sort((a, b) => String(a.sku).localeCompare(String(b.sku), 'pt-BR')).forEach(p => {
      ordenarTams((p.tamanhos || []).concat([...(extra.get(p.id) || [])])).forEach(t => out.push({ p, tamanho: t }));
    });
    return out;
  }

  /* O MÍNIMO E O MÁXIMO DE CADA LINHA (28/10/2026, Jones: "Os estoques mínimos
     e máximos devem ser configurados para cada linha do Estoque atual, de
     acordo com o sku, cor e tamanho"). Antes era um par só, para o estoque
     inteiro (meta.acabConfig.estoqueMin/estoqueMax, 150/1000) — o mesmo
     mínimo para a camiseta preta G, que gira muito, e para a G3 tricolor.
     Agora cada SKU (que já traz a cor) e tamanho tem o seu, editado nas
     colunas Mín. e Máx. do Estoque atual e guardado no cadastro do SKU:
     p.limites = { G: { min: 150, max: 1000 }, ... }. O SKU lido do programa
     ganha registro próprio ao receber o limite (fixarProduto), como ao
     receber código ou preço. Linha sem limite fica "Sem limite": não entra
     no mínimo nem no excedente. O par geral antigo deixou de valer. */
  const _numLim = v => (v === '' || v == null || !Number.isFinite(Number(v)) ? null : Number(v));
  function limiteDe(p, tamanho) {
    const l = (p && p.limites && p.limites[tamanho]) || {};
    return { min: _numLim(l.min), max: _numLim(l.max) };
  }
  const statusDe = (total, lim) => {
    const l = lim || {};
    if (l.min == null && l.max == null) return { k: 'sem', rotulo: 'Sem limite', cor: '#eceff1' };
    if (l.min != null && total <= l.min) return { k: 'min', rotulo: 'Estoque mínimo', cor: '#f6dcda' };
    if (l.max != null && l.max > 0 && total > l.max) return { k: 'max', rotulo: 'Excedente', cor: '#fbeccb' };
    return { k: 'ok', rotulo: 'Estoque ok', cor: '#d6f0db' };
  };
  const podeLimites = () => (typeof temAcesso === 'function' ? temAcesso('estoque-confeccao') : true);
  const badgeStatus = st => `<span class="badge" style="background:${st.cor};">${esc(st.rotulo)}</span>`;
  // A célula editável do limite: número em branco = sem limite.
  const celLimite = (p, tamanho, campo, v) => `<td style="text-align:right;"><input type="number" min="0" step="1" inputmode="numeric" class="acab-lim"
    style="width:72px;text-align:right;" value="${v == null ? '' : v}" placeholder="—" ${podeLimites() ? '' : 'disabled'}
    title="${campo === 'min' ? 'Estoque mínimo deste SKU e tamanho: com esta quantidade ou menos, fica em Estoque mínimo' : 'Estoque máximo deste SKU e tamanho: acima disto, fica em Excedente'} (em branco = sem limite)"
    onchange="acabSalvarLimite('${esc(p.id)}','${esc(tamanho)}','${campo}',this)"></td>`;

  window.acabSalvarLimite = async function (produtoId, tamanho, campo, el) {
    if (!exigirEdicao('configurar o estoque de produtos')) return;
    let p = produtoPorId(produtoId);
    if (!p) return toast('SKU não encontrado', 'err');
    const bruto = String((el && el.value) || '').trim();
    const v = bruto === '' ? null : Math.round(Number(bruto));
    if (v != null && !(v >= 0)) return toast('Limite inválido', 'err');
    const atual = limiteDe(p, tamanho);
    const novo = Object.assign({}, atual, { [campo]: v });
    if (novo.min != null && novo.max != null && novo.max > 0 && novo.max < novo.min) {
      if (el) el.value = atual[campo] == null ? '' : atual[campo];
      return toast(`O máximo (${fmtN(novo.max)}) não pode ser menor que o mínimo (${fmtN(novo.min)})`, 'err');
    }
    p = fixarProduto(p) || p;
    p.limites = Object.assign({}, p.limites || {});
    if (novo.min == null && novo.max == null) delete p.limites[tamanho];
    else p.limites[tamanho] = { min: novo.min, max: novo.max };
    p.atualizadoEm = new Date().toISOString(); p.atualizadoPor = quemSou();
    esquecerProdutos();
    desfazerNomearAcao(`estoque ${campo === 'min' ? 'mínimo' : 'máximo'} de ${p.sku} ${tamanho}`);
    await saveState('acabProdutos');
    // Só o status da linha muda: refazer a tabela tiraria o foco do próximo campo.
    const tr = el && el.closest ? el.closest('tr') : null;
    const cel = tr && tr.querySelector('.acab-st');
    if (cel) cel.innerHTML = badgeStatus(statusDe(saldos(todosMovs()).de(p.id, tamanho).total, limiteDe(p, tamanho)));
  };

  /* Para a Contabilidade: por SKU, a quantidade em estoque — o mesmo formato
     que o Render publicava em produtos_acabados (só quantidades; a
     Contabilidade valora ao custo). Vai dentro do contabSnapshot. */
  function acabSnapshotContab() {
    const ms = todosMovs();
    const S = saldos(ms);
    const itens = [];
    let total = 0;
    produtos().forEach(p => {
      let q = 0;
      ordenarTams((p.tamanhos || []).concat(ms.filter(m => m.produtoId === p.id).map(m => m.tamanho)))
        .forEach(t => { q += S.de(p.id, t).total; });
      if (!q) return;
      total += q;
      itens.push({ codigo: p.codigo || null, sku: p.sku || '', descricao: p.descricao || p.sku || '', cor: p.cor || '', quantidade: q, preco: Number(p.preco) || 0 });
    });
    return { geradoEm: new Date().toISOString(), total, itens };
  }
  window.acabSnapshotContab = acabSnapshotContab;

  /* ------------------------------------------------------- os seletores */
  const opcoesProdutos = sel => '<option value="">Escolha o SKU…</option>' + produtos().slice()
    .sort((a, b) => String(a.sku).localeCompare(String(b.sku), 'pt-BR'))
    .map(p => `<option value="${esc(p.id)}"${p.id === sel ? ' selected' : ''}>${esc(rotuloProduto(p))}${p.descricao ? ' — ' + esc(p.descricao) : ''}</option>`).join('');
  const opcoesTams = (p, sel) => (p ? ordenarTams(p.tamanhos || []) : [])
    .map(t => `<option value="${esc(t)}"${t === sel ? ' selected' : ''}>${esc(t)}</option>`).join('') || '<option value="">—</option>';
  /* ENTRADA E SAÍDA À MÃO (21/10/2026, Junior: "insira entrada/saída manual
     no estoque confecção"). O estoque começou do zero e, sem SKU cadastrado, a
     tela de Entradas e saídas nem mostrava os formulários. Agora o SKU é
     DIGITADO (a lista sugere os cadastrados e os do catálogo) e o SKU que
     ainda não existe nasce na própria entrada, com os tamanhos da linha. A
     saída continua exigindo SKU cadastrado e estoque no tamanho. */
  const catalogo = () => ((typeof catalogoSkus !== 'undefined' && Array.isArray(catalogoSkus)) ? catalogoSkus : []);
  const doCatalogo = sku => { const s = norm(sku); return s ? catalogo().find(x => x && norm(x.item) === s) || null : null; };
  const listaSkus = () => `<datalist id="acab-skus">${[...new Set(produtos().map(p => norm(p.sku))
    .concat(catalogo().map(x => norm(x && x.item))).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'))
    .map(k => `<option value="${esc(k)}">`).join('')}</datalist>`;
  const produtoDoCampo = prefixo => { const v = val(prefixo + '-produto'); return produtoPorId(v) || produtoPorSku(v); };
  // O que a caixa de tamanho oferece: os do SKU cadastrado, ou os da linha do SKU novo.
  function tamsDoCampo(prefixo) {
    const p = produtoDoCampo(prefixo);
    if (p) return p;
    const sku = norm(val(prefixo + '-produto'));
    if (!sku) return null;
    const c = doCatalogo(sku);
    return { tamanhos: tamsDaLinha((c && c.tipo) || partesSku(sku).linha) };
  }
  function infoSku(prefixo) {
    const el = document.getElementById(prefixo + '-info');
    if (!el) return;
    const sku = norm(val(prefixo + '-produto'));
    const p = produtoDoCampo(prefixo);
    const t = asTam(val(prefixo + '-tam'));
    if (!sku) { el.textContent = ''; return; }
    if (!p) {
      const c = doCatalogo(sku);
      el.innerHTML = prefixo === 'acab-sai' ? '<span style="color:#c0392b;">SKU não cadastrado: não há estoque para sair.</span>'
        : `SKU novo${c ? ' (do catálogo' + (c.descricao ? ': ' + esc(c.descricao) : '') + ')' : ''}: é cadastrado junto com a entrada.`;
      return;
    }
    const disp = t ? saldos(todosMovs()).de(p.id, t).total : null;
    el.innerHTML = esc(p.descricao || '') + (disp != null ? `${p.descricao ? ' · ' : ''}em estoque no ${esc(t)}: <b>${fmtN(disp)}</b>` : '');
    // Na entrada, o que ela vai abater das OS (ver planoAbateOS).
    const qtd = parseInt(val('acab-ent-qtd'), 10);
    if (prefixo === 'acab-ent' && t && abaterMarcado()) {
      const plano = planoAbateOS(p.id, t, qtd > 0 ? qtd : Infinity);
      const soma = plano.reduce((s, a) => s + a.qtd, 0);
      if (plano.length) el.innerHTML += qtd > 0
        ? `<div>Vai abater das OS em Estoque: <b>${esc(textoAbate(plano))}</b>${qtd > soma ? ` · ${fmtN(qtd - soma)} entra(m) como peça nova` : ''}</div>`
        : `<div>OS em Estoque com saldo neste tamanho: ${esc(textoAbate(plano))}</div>`;
    }
  }
  window.acabInfoSku = infoSku;

  function trocarProduto(prefixo) {
    const p = prefixo === 'acab-aj' ? produtoPorId(val(prefixo + '-produto')) : tamsDoCampo(prefixo);
    const sel = document.getElementById(prefixo + '-tam');
    if (sel) sel.innerHTML = opcoesTams(p, sel.value);
    if (prefixo === 'acab-aj') infoAjuste();
    else infoSku(prefixo);
  }

  // O "+ entrada" / "− saída" de uma linha do Estoque atual: abre Entradas e
  // saídas com o SKU e o tamanho já escolhidos.
  let _preMov = null;
  window.acabLancar = function (tipo, produtoId, tamanho) {
    const p = produtoPorId(produtoId);
    _preMov = { tipo, sku: p ? p.sku : '', tamanho };
    goto('acab-mov');
  };
  window.acabTrocarProduto = trocarProduto;

  const semSkus = () => `<div class="info-box">Nenhum SKU ainda. Os SKUs vêm direto do programa: o <b>SKU</b> do desenho técnico ou do modelo, e o da OS (linha do modelo + <b>Sigla SKU</b> da cor) — preencha lá e ele aparece aqui na hora. O catálogo inteiro vem por <a href="#acab-produtos" onclick="goto('acab-produtos');return false;">Produtos (SKUs)</a>.</div>`;

  /* ================================================= TELA: ESTOQUE ATUAL */
  const filtro = { busca: '', linha: '', status: '' };
  function renderEstoqueAtual() {
    const box = document.getElementById('acab-estoque-painel');
    if (!box) return;
    const dasOS = _entradasDasOS();
    const ms = todosMovs(dasOS);
    const S = saldos(ms);
    const linhas = linhasSkuTam(ms);
    const busca = norm(filtro.busca);
    let tPecas = 0, tValor = 0, tMin = 0;
    const skusVistos = new Set();
    const rows = [];
    const imgs = imagensDosSkus();
    linhas.forEach(({ p, tamanho }) => {
      if (filtro.linha && norm(p.linha) !== norm(filtro.linha)) return;
      if (busca && !norm([p.sku, p.descricao, p.codigo, tamanho].join(' ')).includes(busca)) return;
      const e = S.de(p.id, tamanho);
      const lim = limiteDe(p, tamanho);
      const st = statusDe(e.total, lim);
      if (filtro.status && st.k !== filtro.status) return;
      const c = cmp(ms, p.id, tamanho);
      skusVistos.add(p.id);
      tPecas += e.total; tValor += Math.max(0, e.total) * c.medio;
      if (st.k === 'min') tMin++;
      rows.push(`<tr>
        ${celMiniatura(imgs, p)}
        <td style="font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${esc(rotuloProduto(p))}</td>
        <td>${esc(p.descricao || '')}</td>
        <td><b>${esc(tamanho)}</b></td>
        <td style="${mono}">${e.ult ? fmtN(e.ult.contado) : '—'}</td>
        <td class="muted" style="font-size:11px;white-space:nowrap;">${e.ult ? fmtD(e.ult.data) : '—'}</td>
        <td style="${mono}" title="${e.entOS ? fmtN(e.entOS) + ' das OS' : ''}">${fmtN(e.ent)}${e.entOS ? `<div class="muted" style="font-size:10px;">${fmtN(e.entOS)} das OS</div>` : ''}</td>
        <td style="${mono}">${fmtN(e.sai)}</td>
        <td style="${mono}font-weight:700;${e.total < 0 ? 'color:#c0392b;background:#fdecea;' : ''}">${fmtN(e.total)}</td>
        ${celLimite(p, tamanho, 'min', lim.min)}${celLimite(p, tamanho, 'max', lim.max)}
        <td style="${mono}">${fmtR(c.medio)}</td>
        <td class="acab-st">${badgeStatus(st)}</td>
        <td class="col-actions row-actions admin-only" style="white-space:nowrap;"><button onclick="acabLancar('ent','${esc(p.id)}','${esc(tamanho)}')" title="Entrada à mão deste SKU e tamanho">+ entrada</button><button onclick="acabLancar('sai','${esc(p.id)}','${esc(tamanho)}')" title="Saída à mão deste SKU e tamanho">− saída</button></td></tr>`);
    });
    const linhasCad = [...new Set(produtos().map(p => p.linha).filter(Boolean))].sort();
    const c = cfg();
    const pend = dasOS.pendencias;
    const pendHtml = !pend.length ? '' : `<div class="card" style="border-left:3px solid var(--alert, #c0392b);">
      <h2 style="margin-top:0;font-size:14px;">OS esperando SKU (${pend.length})</h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Estas OS marcaram a etapa <b>Estoque</b>, mas a entrada não pôde ser lançada sozinha. Cadastre o SKU (a entrada aparece na hora), acerte o SKU no modelo/desenho/cor da OS, ou lance à mão em Entradas e saídas e dispense a OS daqui.</div>
      <table class="table"><thead><tr><th>OS</th><th>Modelo</th><th style="text-align:right;">Peças</th><th>Entrou em</th><th>Por que não lançou</th><th class="col-actions admin-only">Ações</th></tr></thead><tbody>
      ${pend.map(x => `<tr><td><b>${esc(x.os)}</b></td><td>${esc(x.modelo)}</td><td style="${mono}">${fmtN(x.pecas)}</td><td>${fmtD(x.data)}</td><td>${esc(x.motivo)}</td>
        <td class="col-actions row-actions admin-only">${x.cadastrar ? `<button onclick="acabCadastrarSkuDaOS('${esc(x.cadastrar)}','${esc((x.tams || []).join(','))}')">cadastrar ${esc(x.cadastrar)}</button>` : ''}<button onclick="acabDispensarOS('${esc(x.osId)}')" title="Não lançar esta OS (por exemplo, já foi lançada à mão)">dispensar</button></td></tr>`).join('')}
      </tbody></table></div>`;
    box.innerHTML = `
      <div class="exp-resumo" style="margin-bottom:12px;">
        <div class="item"><div class="num">${fmtN(skusVistos.size)}</div><div class="lbl">SKUs</div></div>
        <div class="item"><div class="num">${fmtN(tPecas)}</div><div class="lbl">Peças em estoque</div></div>
        <div class="item${tMin ? ' alerta' : ''}"><div class="num">${fmtN(tMin)}</div><div class="lbl">No estoque mínimo</div></div>
        <div class="item"><div class="num">${fmtR(tValor)}</div><div class="lbl">Valor (custo médio)</div></div>
      </div>
      ${dasOS.semInicio
        ? `<div class="info-box" style="margin-bottom:12px;">A contagem do estoque ainda não começou: nenhuma OS lança entrada e nada pode ser lançado à mão. Dê o início quando a prateleira estiver pronta para contar — as OS que entrarem em Estoque e os lançamentos daqui contam a partir dele, juntos.
            <div style="margin-top:8px;"><button class="btn primary admin-only" onclick="acabIniciarContagem()">▶ Iniciar contagem do estoque</button></div></div>`
        : `<div class="muted" style="font-size:12px;margin-bottom:10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">Contagem ${c.inicioPor ? 'iniciada' : 'marcada sozinha'} em <b>${esc(fmtDH(c.inicio))}</b>${c.inicioPor ? ' por ' + esc(c.inicioPor) : ' (na primeira abertura da tela)'} — OS e lançamentos contam daí em diante.
            <button class="btn admin-only" onclick="acabIniciarContagem()" title="Recomeça a contagem agora: as OS que já estão em Estoque e os lançamentos feitos até aqui deixam de contar (ficam guardados)">${c.inicioPor ? '↻ Reiniciar contagem' : '▶ Iniciar contagem do estoque'}</button></div>`}
      ${pendHtml}
      <div class="card">
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;">
          <input type="search" placeholder="Buscar SKU, código, descrição ou tamanho…" value="${esc(filtro.busca)}" oninput="acabFiltro('busca',this.value)" style="flex:2;min-width:200px;">
          <select onchange="acabFiltro('linha',this.value)" style="flex:1;min-width:140px;"><option value="">Todas as linhas</option>${linhasCad.map(l => `<option value="${esc(l)}"${l === filtro.linha ? ' selected' : ''}>${esc(l)}</option>`).join('')}</select>
          <select onchange="acabFiltro('status',this.value)" style="flex:1;min-width:140px;">
            <option value="">Todos os status</option>
            <option value="min"${filtro.status === 'min' ? ' selected' : ''}>Só estoque mínimo</option>
            <option value="max"${filtro.status === 'max' ? ' selected' : ''}>Só excedente</option>
            <option value="ok"${filtro.status === 'ok' ? ' selected' : ''}>Só estoque ok</option>
            <option value="sem"${filtro.status === 'sem' ? ' selected' : ''}>Só sem limite</option>
          </select>
          <button class="btn" onclick="acabExportarCsv()" title="A posição do estoque numa planilha (abre no Excel)">Exportar CSV ↓</button>
          <button class="btn primary admin-only" onclick="acabLancar('ent','','')" title="Lançar entrada ou saída à mão">+ Entrada / − Saída</button>
        </div>
        ${!produtos().length ? semSkus() : `<div style="overflow:auto;max-height:70vh;"><table class="table"><thead><tr>
          <th title="O desenho técnico do produto">Desenho</th><th>SKU</th><th>Descrição</th><th>Tam.</th><th style="text-align:right;" title="A última contagem de inventário">Contagem</th><th>Data</th>
          <th style="text-align:right;">Entradas</th><th style="text-align:right;">Saídas</th><th style="text-align:right;">Estoque</th>
          <th style="text-align:right;" title="Estoque mínimo deste SKU, cor e tamanho">Mín.</th><th style="text-align:right;" title="Estoque máximo deste SKU, cor e tamanho">Máx.</th>
          <th style="text-align:right;" title="Custo médio ponderado">Custo médio</th><th>Status</th><th class="col-actions admin-only">Lançar à mão</th></tr></thead>
          <tbody>${rows.join('') || '<tr><td colspan="14" class="empty">Nada com estes filtros.</td></tr>'}</tbody></table></div>`}
      </div>`;
  }
  window.acabFiltro = (k, v) => {
    filtro[k] = v;
    renderEstoqueAtual();
    if (k === 'busca') { const i = document.querySelector('#acab-estoque-painel input[type=search]'); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } }
  };

  window.acabExportarCsv = function () {
    const ms = todosMovs();
    const S = saldos(ms);
    const cel = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const linhas = [['Código', 'SKU', 'Descrição', 'Linha', 'Tamanho', 'Entradas', 'Entradas das OS', 'Saídas', 'Ajustes', 'Estoque', 'Mínimo', 'Máximo', 'Custo médio', 'Valor', 'Status'].map(cel).join(';')];
    linhasSkuTam(ms).forEach(({ p, tamanho }) => {
      const e = S.de(p.id, tamanho);
      const c = cmp(ms, p.id, tamanho);
      const br = v => (Number(v) || 0).toFixed(2).replace('.', ',');
      const lim = limiteDe(p, tamanho);
      linhas.push([p.codigo || '', p.sku, p.descricao || '', p.linha || '', tamanho, e.ent, e.entOS, e.sai, e.aj, e.total,
        lim.min == null ? '' : lim.min, lim.max == null ? '' : lim.max, br(c.medio), br(Math.max(0, e.total) * c.medio), statusDe(e.total, lim).rotulo].map(cel).join(';'));
    });
    const blob = new Blob(['﻿' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'estoque-confeccao_' + hoje() + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  };

  window.acabIniciarContagem = async function () {
    if (!exigirEdicao('iniciar a contagem do estoque de produtos')) return;
    const c = cfg();
    const agora = new Date().toISOString();
    const antes = c.inicio ? todosMovs() : [];
    const nMovs = antes.filter(m => m.origem !== 'os').length;
    const nOS = new Set(antes.filter(m => m.origem === 'os').map(m => m.osId)).size;
    const msg = c.inicio && c.inicioPor
      ? `Reiniciar a contagem do estoque agora (${fmtDH(agora)})?

O estoque volta a zero: ${nOS} OS já em Estoque e ${nMovs} lançamento(s) feitos até aqui deixam de contar. Nada é apagado — recuar o início em Produtos (SKUs) traz de volta.`
      : `Iniciar a contagem do estoque agora (${fmtDH(agora)})?

Daqui em diante contam as OS que entrarem em Estoque e os lançamentos feitos aqui (que abatem essas OS).${nOS || nMovs ? `

${nOS} OS já em Estoque e ${nMovs} lançamento(s) de antes deixam de contar (ficam guardados).` : ''}`;
    if (!confirm(msg)) return;
    desfazerNomearAcao('início da contagem do Estoque de produtos');
    await gravarCfg({ inicio: agora, inicioPor: quemSou() || 'admin' });
    toast('Contagem do estoque iniciada', 'ok');
    renderAtual();
  };

  window.acabDispensarOS = async function (osId) {
    if (!exigirEdicao('configurar o estoque de produtos')) return;
    const o = (STATE.ordens || []).find(x => x.id === osId);
    if (!confirm(`Dispensar a OS ${o ? o.os : ''}?\n\nEla deixa de lançar entrada no Estoque de produtos. Use quando a entrada já foi feita à mão.`)) return;
    desfazerNomearAcao(`dispensa da OS ${o ? o.os : ''} no Estoque de produtos`);
    await gravarCfg({ osDispensadas: cfg().osDispensadas.concat([osId]) });
    toast('OS dispensada', 'ok');
    renderEstoqueAtual();
  };

  /* ============================================ TELA: ENTRADAS E SAÍDAS */
  function renderMovimentos() {
    const box = document.getElementById('acab-mov-painel');
    if (!box) return;
    const ms = todosMovs().slice().sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')));
    const ent = ms.filter(m => m.tipo === 'entrada').slice(0, 150);
    const sai = ms.filter(m => m.tipo === 'saida').slice(0, 150);
    const pre = _preMov; _preMov = null;
    const skuE = pre && pre.tipo === 'ent' ? pre.sku : val('acab-ent-produto');
    const skuS = pre && pre.tipo === 'sai' ? pre.sku : val('acab-sai-produto');
    const tamE = pre && pre.tipo === 'ent' ? pre.tamanho : val('acab-ent-tam');
    const tamS = pre && pre.tipo === 'sai' ? pre.tamanho : val('acab-sai-tam');
    const campoSku = (prefixo, v) => `<input type="text" id="${prefixo}-produto" list="acab-skus" value="${esc(norm(v))}" autocomplete="off"
      placeholder="Digite ou escolha o SKU (ex.: CM.LISA-PRE)" oninput="acabTrocarProduto('${prefixo}')" style="text-transform:uppercase;">
      <div class="field-hint" id="${prefixo}-info"></div>`;
    const origem = m => m.origem === 'os' ? `<span class="badge" style="background:#e3ecf7;" title="Etapa Estoque marcada na OS — desmarque na OS para estornar">OS ${esc(m.os)}</span>`
      : m.origem === 'erp' ? '<span class="badge" style="background:#efe7f6;">ERP</span>'
      : m.origem === 'dre' ? `<span class="badge" style="background:#e0f2ec;" title="Vendas das plataformas (Painel DRE): ${esc((m.loja || '') + ' · pedido ' + (m.pedidoNumero || ''))}">DRE</span>` : '';
    const apagar = m => m.origem === 'os' ? '' : `<button onclick="acabApagarMov('${esc(m.id)}')">apagar</button>`;
    box.innerHTML = `${listaSkus()}
      ${produtos().length ? '' : `<div class="info-box" style="margin-bottom:12px;">Nenhum SKU ainda. Os SKUs vêm direto do programa: o <b>SKU</b> do desenho técnico ou do modelo, e o da OS (linha do modelo + <b>Sigla SKU</b> da cor) — preencha lá e ele aparece aqui na hora. A entrada também cadastra o SKU digitado; para trazer o catálogo inteiro, use <a href="#acab-produtos" onclick="goto('acab-produtos');return false;">Produtos (SKUs)</a>.</div>`}
      <div class="grid-2" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px;">
        <div class="card admin-only">
          <h2 style="margin-top:0;font-size:15px;">Registrar entrada</h2>
          <div class="muted" style="font-size:12px;margin-bottom:8px;">Compra de fornecedor ou peça que chegou fora de uma OS. A OS que marca a etapa <b>Estoque</b> já entra sozinha.</div>
          <div class="form-grid cols-2">
            <div class="field full"><label>SKU *</label>${campoSku('acab-ent', skuE)}</div>
            <div class="field"><label>Tamanho *</label><select id="acab-ent-tam" onchange="acabInfoSku('acab-ent')"></select></div>
            <div class="field"><label>Data *</label><input type="date" id="acab-ent-data" value="${hoje()}"></div>
            <div class="field"><label>Quantidade *</label><input type="number" id="acab-ent-qtd" min="1" step="1" oninput="acabInfoSku('acab-ent')"></div>
            <div class="field"><label>Custo unitário (R$)</label><input type="text" inputmode="decimal" id="acab-ent-custo" placeholder="0,00"></div>
            <div class="field full"><label>Observação</label><input type="text" id="acab-ent-obs" placeholder="Ex.: NF 1234 / fornecedor"></div>
            <div class="field full"><label style="font-weight:normal;cursor:pointer;" title="Tira estas peças das OS do mesmo SKU e tamanho que já entraram em Estoque, da mais recente para a mais antiga — para não contar duas vezes. Desmarque para peça que não veio de OS (compra, devolução).">
              <input type="checkbox" id="acab-ent-abate" checked onchange="acabInfoSku('acab-ent')"> abater das OS em Estoque (da mais recente para a mais antiga)</label></div>
          </div>
          <button class="btn primary" onclick="acabRegistrarEntrada()">+ Registrar entrada</button>
        </div>
        <div class="card admin-only">
          <h2 style="margin-top:0;font-size:15px;">Registrar saída</h2>
          <div class="muted" style="font-size:12px;margin-bottom:8px;">Não sai mais do que há no estoque do tamanho. As vendas do ERP entram em lote pela tela Saídas do ERP.</div>
          <div class="form-grid cols-2">
            <div class="field full"><label>SKU *</label>${campoSku('acab-sai', skuS)}</div>
            <div class="field"><label>Tamanho *</label><select id="acab-sai-tam" onchange="acabInfoSku('acab-sai')"></select></div>
            <div class="field"><label>Destino *</label><select id="acab-sai-destino">${DESTINOS.map(d => `<option value="${d.k}">${esc(d.rotulo)}</option>`).join('')}</select></div>
            <div class="field"><label>Data *</label><input type="date" id="acab-sai-data" value="${hoje()}"></div>
            <div class="field"><label>Quantidade *</label><input type="number" id="acab-sai-qtd" min="1" step="1"></div>
            <div class="field"><label>Valor unitário (R$)</label><input type="text" inputmode="decimal" id="acab-sai-valor" placeholder="0,00"></div>
          </div>
          <button class="btn primary" onclick="acabRegistrarSaida()">− Registrar saída</button>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px;">
        <div class="card"><h2 style="margin-top:0;font-size:15px;">Últimas entradas</h2>
          <div style="overflow:auto;max-height:440px;"><table class="table"><thead><tr><th>Data</th><th>SKU</th><th>Tam.</th><th style="text-align:right;">Qtd</th><th style="text-align:right;">Custo</th><th>Origem</th><th class="col-actions admin-only"></th></tr></thead><tbody>
          ${ent.map(m => `<tr><td>${fmtD(m.data)}</td><td>${esc(rotuloProduto(produtoPorId(m.produtoId)))}</td><td>${esc(m.tamanho)}</td><td style="${mono}">${fmtN(m.qtd)}${m.abatido ? `<div class="muted" style="font-size:10px;" title="A OS lançou ${fmtN(m.qtdOS)}; ${fmtN(m.abatido)} já tinham entrado à mão">de ${fmtN(m.qtdOS)} · −${fmtN(m.abatido)} à mão</div>` : ''}</td><td style="${mono}">${m.origem === 'os' ? '—' : fmtR(m.custo)}</td><td>${origem(m)}${m.obs ? `<div class="muted" style="font-size:10px;">${esc(m.obs)}</div>` : ''}${Array.isArray(m.abate) && m.abate.length ? `<div class="muted" style="font-size:10px;">abateu ${esc(textoAbate(m.abate))}</div>` : ''}</td><td class="col-actions row-actions admin-only">${apagar(m)}</td></tr>`).join('') || '<tr><td colspan="7" class="empty">Nenhuma entrada.</td></tr>'}
          </tbody></table></div></div>
        <div class="card"><h2 style="margin-top:0;font-size:15px;">Últimas saídas</h2>
          <div style="overflow:auto;max-height:440px;"><table class="table"><thead><tr><th>Data</th><th>SKU</th><th>Tam.</th><th>Destino</th><th style="text-align:right;">Qtd</th><th style="text-align:right;">Valor</th><th class="col-actions admin-only"></th></tr></thead><tbody>
          ${sai.map(m => `<tr><td>${fmtD(m.data)}</td><td>${esc(rotuloProduto(produtoPorId(m.produtoId)))}</td><td>${esc(m.tamanho)}</td><td>${esc((DESTINOS.find(d => d.k === m.destino) || { rotulo: m.destino || '' }).rotulo.replace('Estoque → ', ''))} ${origem(m)}</td><td style="${mono}">${fmtN(m.qtd)}</td><td style="${mono}">${fmtR(m.valor)}</td><td class="col-actions row-actions admin-only">${apagar(m)}</td></tr>`).join('') || '<tr><td colspan="7" class="empty">Nenhuma saída.</td></tr>'}
          </tbody></table></div></div>
      </div>
      <!-- O INVENTÁRIO MORA AQUI, NA PARTE DE BAIXO (28/10/2026, Jones: "Migre o
           conteúdo do campo Inventário para dentro de Entradas e saídas, na
           parte de baixo"). Era uma tela própria do menu; o ajuste é um
           lançamento como a entrada e a saída. Desenhado por renderInventario. -->
      <h2 style="font-size:16px;margin:22px 0 10px;">Inventário</h2>
      <div id="acab-inventario-painel"></div>`;
    renderInventario();
    [['acab-ent', tamE], ['acab-sai', tamS]].forEach(([prefixo, t]) => {
      const sel = document.getElementById(prefixo + '-tam');
      if (sel) sel.innerHTML = opcoesTams(tamsDoCampo(prefixo), asTam(t));
      infoSku(prefixo);
    });
    if (pre) {
      const q = document.getElementById(pre.tipo === 'sai' ? 'acab-sai-qtd' : (pre.sku ? 'acab-ent-qtd' : 'acab-ent-produto'));
      if (q) { q.scrollIntoView({ block: 'center' }); q.focus(); }
    }
  }

  function novoMov(campos) {
    return Object.assign({ id: uid(), origem: 'manual', criadoEm: new Date().toISOString(), criadoPor: quemSou() }, campos);
  }

  window.acabRegistrarEntrada = async function () {
    if (!exigirEdicao('lançar no estoque de produtos') || !exigirInicio()) return;
    const sku = norm(val('acab-ent-produto'));
    let p = produtoDoCampo('acab-ent');
    const tamanho = asTam(val('acab-ent-tam'));
    const qtd = parseInt(val('acab-ent-qtd'), 10);
    const custo = val('acab-ent-custo').trim() ? num(val('acab-ent-custo')) : 0;
    if (!p && !sku) return toast('Informe o SKU', 'err');
    if (!tamanho) return toast('Escolha o tamanho', 'err');
    if (!(qtd > 0)) return toast('Informe a quantidade', 'err');
    if (!(custo >= 0)) return toast('Custo inválido', 'err');
    // SKU que ainda não está cadastrado nasce aqui (o do catálogo traz linha,
    // cor e descrição; o digitado fora do catálogo pede confirmação).
    let novo = null;
    if (!p) {
      if (!/^[A-Z0-9][A-Z0-9.]*(-[A-Z0-9.]+)*$/.test(sku)) return toast('SKU inválido: use linha + traço + sigla da cor (ex.: CM.LISA-PRE)', 'err');
      const c = doCatalogo(sku);
      if (!c && !confirm(`O SKU ${sku} não está cadastrado nem no catálogo.\n\nCadastrar agora e lançar a entrada?`)) return;
      const pt = partesSku(sku);
      novo = { id: uid(), codigo: null, sku, linha: norm(c && c.tipo) || pt.linha, cor: norm(c && c.cor) || pt.cor,
        descricao: (c && c.descricao) || '', tamanhos: tamsDaLinha((c && c.tipo) || pt.linha), preco: 0,
        criadoEm: new Date().toISOString(), criadoPor: quemSou(), origem: c ? 'catalogo' : 'entrada' };
      if (!novo.tamanhos.includes(tamanho)) return toast(`O tamanho ${tamanho} não é da linha ${novo.linha}`, 'err');
      cadastro().push(novo);
      esquecerProdutos();
      p = novo;
    } else if (fixarProduto(p)) novo = p;
    const plano = abaterMarcado() ? planoAbateOS(p.id, tamanho, qtd) : [];
    const mov = novoMov({ tipo: 'entrada', produtoId: p.id, tamanho, data: val('acab-ent-data') || hoje(), qtd, custo, obs: val('acab-ent-obs').trim() });
    if (plano.length) mov.abate = plano;
    movs().push(mov);
    desfazerNomearAcao(`entrada de ${qtd} ${p.sku} ${tamanho}`);
    if (novo) await saveState('acabProdutos');
    await saveState('acabMov');
    toast(`Entrada: ${qtd} × ${p.sku} ${tamanho}${novo ? ' (SKU cadastrado)' : ''}${plano.length ? ' · abateu ' + textoAbate(plano) : ''}`, 'ok');
    const q = document.getElementById('acab-ent-qtd'); if (q) q.value = '';
    renderMovimentos();
  };

  window.acabRegistrarSaida = async function () {
    if (!exigirEdicao('lançar no estoque de produtos') || !exigirInicio()) return;
    const sku = norm(val('acab-sai-produto'));
    const p = produtoDoCampo('acab-sai');
    const tamanho = asTam(val('acab-sai-tam'));
    const qtd = parseInt(val('acab-sai-qtd'), 10);
    const valor = val('acab-sai-valor').trim() ? num(val('acab-sai-valor')) : 0;
    const destino = DESTINOS.some(d => d.k === val('acab-sai-destino')) ? val('acab-sai-destino') : 'cliente';
    if (!p) return toast(sku ? `O SKU ${sku} não está cadastrado: não há estoque para sair` : 'Informe o SKU', 'err');
    if (!tamanho) return toast('Escolha o tamanho', 'err');
    if (!(qtd > 0)) return toast('Informe a quantidade', 'err');
    if (!(valor >= 0)) return toast('Valor inválido', 'err');
    const disp = saldos(todosMovs()).de(p.id, tamanho).total;
    if (qtd > disp) return toast(`Estoque insuficiente: há ${fmtN(disp)} de ${p.sku} ${tamanho}`, 'err');
    const fixou = fixarProduto(p);
    movs().push(novoMov({ tipo: 'saida', produtoId: p.id, tamanho, data: val('acab-sai-data') || hoje(), qtd, valor, destino }));
    desfazerNomearAcao(`saída de ${qtd} ${p.sku} ${tamanho}`);
    if (fixou) await saveState('acabProdutos');
    await saveState('acabMov');
    toast(`Saída: ${qtd} × ${p.sku} ${tamanho}`, 'ok');
    const q = document.getElementById('acab-sai-qtd'); if (q) q.value = '';
    renderMovimentos();
  };

  window.acabApagarMov = async function (id) {
    if (!exigirEdicao('apagar lançamento do estoque de produtos')) return;
    const m = movs().find(x => x.id === id);
    if (!m) return;
    const p = produtoPorId(m.produtoId);
    const nome = m.tipo === 'entrada' ? 'a entrada' : m.tipo === 'saida' ? 'a saída' : 'o ajuste';
    const devolve = Array.isArray(m.abate) && m.abate.length ? `\n\nAs peças que ela abateu voltam para: ${textoAbate(m.abate)}.` : '';
    // A saída do Painel DRE apagada à mão DISPENSA o pedido inteiro: sem isso
    // a próxima sincronização a lançaria de novo (ver vendas-dre.js).
    const doDre = m.origem === 'dre' && m.pedidoChave && typeof window.vendasDreDispensar === 'function';
    const dispensa = doDre ? `\n\nÉ uma venda do Painel DRE (${m.loja || ''}, pedido ${m.pedidoNumero || ''}): o pedido inteiro deixa de ser lançado pela sincronização.` : '';
    if (!confirm(`Apagar ${nome} de ${p ? p.sku : '?'} ${m.tamanho} (${fmtD(m.data)})?${devolve}${dispensa}`)) return;
    STATE.acabMov = movs().filter(x => x.id !== id);
    if (doDre) window.vendasDreDispensar(m.pedidoChave);
    desfazerNomearAcao(`${nome} de ${p ? p.sku : '?'} ${m.tamanho} apagada no Estoque de produtos`);
    await saveState('acabMov');
    if (doDre) await saveState('meta');
    toast('Lançamento apagado', 'ok');
    renderAtual();
  };

  /* ================================================== TELA: INVENTÁRIO */
  function infoAjuste() {
    const el = document.getElementById('acab-aj-info');
    if (!el) return;
    const p = produtoPorId(val('acab-aj-produto'));
    const t = asTam(val('acab-aj-tam'));
    if (!p || !t) { el.innerHTML = 'Escolha o SKU e o tamanho para ver o estoque atual.'; return; }
    const atual = saldos(todosMovs()).de(p.id, t).total;
    const contado = parseInt(val('acab-aj-contado'), 10);
    const dif = Number.isFinite(contado) ? contado - atual : null;
    el.innerHTML = `Estoque no programa: <b>${fmtN(atual)}</b>` + (dif == null ? '' : ` · diferença <b style="color:${dif > 0 ? '#2e7d4f' : dif < 0 ? '#c0392b' : 'inherit'};">${dif > 0 ? '+' : ''}${fmtN(dif)}</b>${dif === 0 ? ' (igual: nada a ajustar)' : dif > 0 ? ' (sobrou)' : ' (faltou)'}`);
  }
  window.acabInfoAjuste = infoAjuste;

  function renderInventario() {
    const box = document.getElementById('acab-inventario-painel');
    if (!box) return;
    if (!produtos().length) { box.innerHTML = `<div class="card">${semSkus()}</div>`; return; }
    const pA = produtoPorId(val('acab-aj-produto'));
    const ajs = movs().filter(m => m.tipo === 'ajuste').sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.criadoEm || '').localeCompare(String(a.criadoEm || ''))).slice(0, 200);
    box.innerHTML = `
      <div class="card admin-only">
        <h2 style="margin-top:0;font-size:15px;">Ajuste de inventário (contagem física)</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">Conte as peças na prateleira e informe o número contado: o programa lança a diferença e o saldo passa a ser o contado. É assim que o estoque começa — uma contagem por SKU e tamanho.</div>
        <div class="form-grid cols-2">
          <div class="field full"><label>SKU *</label><select id="acab-aj-produto" onchange="acabTrocarProduto('acab-aj')">${opcoesProdutos(pA && pA.id)}</select></div>
          <div class="field"><label>Tamanho *</label><select id="acab-aj-tam" onchange="acabInfoAjuste()">${opcoesTams(pA)}</select></div>
          <div class="field"><label>Data *</label><input type="date" id="acab-aj-data" value="${hoje()}"></div>
          <div class="field"><label>Quantidade contada *</label><input type="number" id="acab-aj-contado" min="0" step="1" oninput="acabInfoAjuste()"></div>
          <div class="field"><label>Motivo</label><input type="text" id="acab-aj-motivo" placeholder="Perda, dano, recontagem, contagem inicial…"></div>
        </div>
        <div class="info-box" id="acab-aj-info" style="margin-bottom:10px;"></div>
        <button class="btn primary" onclick="acabRegistrarAjuste()">Registrar ajuste</button>
      </div>
      <div class="card"><h2 style="margin-top:0;font-size:15px;">Últimos ajustes</h2>
        <div style="overflow:auto;max-height:480px;"><table class="table"><thead><tr><th>Data</th><th>SKU</th><th>Tam.</th><th style="text-align:right;">Anterior</th><th style="text-align:right;">Contado</th><th style="text-align:right;">Diferença</th><th>Motivo</th><th class="col-actions admin-only"></th></tr></thead><tbody>
        ${ajs.map(m => `<tr><td>${fmtD(m.data)}</td><td>${esc(rotuloProduto(produtoPorId(m.produtoId)))}</td><td>${esc(m.tamanho)}</td><td style="${mono}">${fmtN(m.anterior)}</td><td style="${mono}">${fmtN(m.contado)}</td>
          <td style="${mono}font-weight:700;color:${m.diferenca > 0 ? '#2e7d4f' : '#c0392b'};">${m.diferenca > 0 ? '+' : ''}${fmtN(m.diferenca)}</td><td>${esc(m.motivo || '')}<div class="muted" style="font-size:10px;">${esc(m.criadoPor || '')}</div></td>
          <td class="col-actions row-actions admin-only"><button onclick="acabApagarMov('${esc(m.id)}')">apagar</button></td></tr>`).join('') || '<tr><td colspan="8" class="empty">Nenhum ajuste.</td></tr>'}
        </tbody></table></div></div>`;
    infoAjuste();
  }

  window.acabRegistrarAjuste = async function () {
    if (!exigirEdicao('lançar no estoque de produtos') || !exigirInicio()) return;
    const p = produtoPorId(val('acab-aj-produto'));
    const tamanho = asTam(val('acab-aj-tam'));
    const contado = parseInt(val('acab-aj-contado'), 10);
    if (!p) return toast('Escolha o SKU', 'err');
    if (!tamanho) return toast('Escolha o tamanho', 'err');
    if (!(contado >= 0)) return toast('Informe a quantidade contada', 'err');
    const anterior = saldos(todosMovs()).de(p.id, tamanho).total;
    const diferenca = contado - anterior;
    if (diferenca === 0) return toast('A contagem é igual ao estoque atual — nada a ajustar', 'err');
    const fixou = fixarProduto(p);
    movs().push(novoMov({ tipo: 'ajuste', produtoId: p.id, tamanho, data: val('acab-aj-data') || hoje(), contado, anterior, diferenca, motivo: val('acab-aj-motivo').trim() }));
    desfazerNomearAcao(`contagem de ${p.sku} ${tamanho}`);
    if (fixou) await saveState('acabProdutos');
    await saveState('acabMov');
    toast(`${p.sku} ${tamanho}: ${diferenca > 0 ? '+' : ''}${fmtN(diferenca)} (agora ${fmtN(contado)})`, 'ok');
    const campo = document.getElementById('acab-aj-contado');
    renderInventario();
    const c2 = document.getElementById('acab-aj-contado');
    if (c2 && campo) c2.focus();
  };

  /* ============================================== TELA: PRODUTOS (SKUs) */
  let _buscaProd = '';
  let _editProdId = null;
  function renderProdutos() {
    const box = document.getElementById('acab-produtos-painel');
    if (!box) return;
    const c = cfg();
    const ms = movs();
    const b = norm(_buscaProd);
    const ps = produtos().slice().sort((a, b2) => String(a.sku).localeCompare(String(b2.sku), 'pt-BR'))
      .filter(p => !b || norm([p.sku, p.descricao, p.codigo, p.linha, p.cor].join(' ')).includes(b));
    const imgs = imagensDosSkus();
    const inicioLocal = c.inicio ? new Date(c.inicio) : null;
    const pad = n => String(n).padStart(2, '0');
    const inicioInput = inicioLocal ? `${inicioLocal.getFullYear()}-${pad(inicioLocal.getMonth() + 1)}-${pad(inicioLocal.getDate())}T${pad(inicioLocal.getHours())}:${pad(inicioLocal.getMinutes())}` : '';
    box.innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;">
        <div class="card admin-only">
          <h2 style="margin-top:0;font-size:15px;">Início do estoque</h2>
          <div class="form-grid cols-2">
            <div class="field full"><label>Início do estoque</label><input type="datetime-local" id="acab-cfg-inicio" value="${inicioInput}">
              <div class="field-hint">OS que entraram na etapa Estoque e lançamentos feitos antes disto não contam (o estoque começou do zero aqui). Para começar agora, use o botão no Estoque atual.</div></div>
          </div>
          <button class="btn primary" onclick="acabSalvarCfg()">Salvar</button>
        </div>
        <!-- O botão "+ Cadastrar SKU" saiu em 28/10/2026 (Jones): os SKUs vêm
             do cadastro do programa (modelo, desenho, OS). -->
      </div>
      <div class="card">
        <input type="search" placeholder="Buscar SKU, código, linha, cor ou descrição…" value="${esc(_buscaProd)}" oninput="acabBuscaProd(this.value)" style="margin-bottom:10px;width:100%;">
        <div style="overflow:auto;max-height:65vh;"><table class="table"><thead><tr><th class="col-actions admin-only">Ações</th><th title="O desenho técnico do produto">Desenho</th><th>Código</th><th>SKU</th><th>Descrição</th><th>Linha</th><th>Cor</th><th>Tamanhos</th><th style="text-align:right;">Preço base</th></tr></thead><tbody>
        ${ps.map(p => `<tr><td class="col-actions row-actions admin-only"><button onclick="acabAbrirProduto('${esc(p.id)}')">editar</button>${p.doPrograma ? '' : `<button onclick="acabExcluirProduto('${esc(p.id)}')">apagar</button>`}</td>
          ${celMiniatura(imgs, p)}
          <td style="font-family:'IBM Plex Mono',monospace;">${p.codigo ? esc(String(p.codigo).padStart(3, '0')) : '<span class="muted">—</span>'}</td><td><b>${esc(p.sku)}</b></td><td>${esc(p.descricao || '')}</td><td>${esc(p.linha || '')}</td><td>${esc(p.cor || '')}</td>
          <td style="font-size:11px;">${esc(ordenarTams(p.tamanhos || []).join(' '))}</td><td style="${mono}">${fmtR(p.preco)}</td></tr>`).join('') || `<tr><td colspan="9" class="empty">${b ? 'Nada na busca.' : 'Nenhum SKU cadastrado.'}</td></tr>`}
        </tbody></table></div>
        <div class="muted" style="font-size:11px;margin-top:6px;">${fmtN(produtos().length)} SKUs · ${fmtN(ms.length)} lançamentos</div>
      </div>`;
  }
  window.acabBuscaProd = v => {
    _buscaProd = v;
    renderProdutos();
    const i = document.querySelector('#acab-produtos-painel input[type=search]');
    if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); }
  };

  window.acabSalvarCfg = async function () {
    if (!exigirEdicao('configurar o estoque de produtos')) return;
    const ini = val('acab-cfg-inicio');
    const inicio = ini ? new Date(ini).toISOString() : '';
    if (!inicio) return toast('Informe o início do estoque', 'err');
    desfazerNomearAcao('início do Estoque de produtos');
    await gravarCfg({ inicio });
    toast('Estoque de produtos: configuração salva', 'ok');
    renderProdutos();
  };

  // Os tamanhos de uma linha nova: o infantil tem a numeração própria.
  const tamsDaLinha = linha => (/INF|KIDS/i.test(String(linha || '')) ? TAMS_INFANTIL.slice() : TAMS_PADRAO.slice());
  // "CM.LISA-PRE" → linha CM.LISA, cor PRE (a mesma regra do traço de skusDaOS).
  const partesSku = sku => { const s = norm(sku); const i = s.indexOf('-'); return i < 0 ? { linha: s, cor: '' } : { linha: s.slice(0, i), cor: s.slice(i + 1) }; };

  window.acabImportarCatalogo = async function () {
    if (!exigirEdicao('cadastrar SKU no estoque de produtos')) return;
    const cat = (typeof catalogoSkus !== 'undefined' && Array.isArray(catalogoSkus)) ? catalogoSkus : [];
    const novos = cat.filter(s => s && s.item && !produtoPorSku(s.item));
    if (!novos.length) return toast('Nenhum SKU novo no catálogo', 'ok');
    if (!confirm(`Cadastrar ${novos.length} SKUs do catálogo?`)) return;
    novos.forEach(s => {
      const pt = partesSku(s.item);
      cadastro().push({ id: uid(), codigo: null, sku: norm(s.item), linha: norm(s.tipo) || pt.linha, cor: norm(s.cor) || pt.cor,
        descricao: s.descricao || '', tamanhos: tamsDaLinha(s.tipo || pt.linha), preco: 0,
        criadoEm: new Date().toISOString(), criadoPor: quemSou(), origem: 'catalogo' });
    });
    esquecerProdutos();
    desfazerNomearAcao(`${novos.length} SKUs trazidos do catálogo`);
    await saveState('acabProdutos');
    toast(`${novos.length} SKUs cadastrados`, 'ok');
    renderProdutos();
  };

  window.acabCadastrarSkuDaOS = function (sku, tams) {
    window.acabAbrirProduto(null, { sku, tamanhos: String(tams || '').split(',').filter(Boolean) });
  };

  window.acabAbrirProduto = function (id, pre) {
    if (!exigirEdicao('cadastrar SKU no estoque de produtos')) return;
    const p = id ? produtoPorId(id) : null;
    _editProdId = p ? p.id : null;
    const base = p || Object.assign({ codigo: '', sku: '', linha: '', cor: '', descricao: '', tamanhos: TAMS_PADRAO, preco: '' }, pre || {});
    if (!p && pre && pre.sku) { const pt = partesSku(pre.sku); base.linha = pt.linha; base.cor = pt.cor; }
    const sel = new Set(base.tamanhos || []);
    document.getElementById('modal-acab-produto-title').textContent = p ? 'Editar SKU · ' + p.sku : 'Cadastrar SKU';
    document.getElementById('modal-acab-produto-fields').innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>SKU *</label><input type="text" id="acab-p-sku" value="${esc(base.sku)}" placeholder="Ex.: CM.LISA-PRE" oninput="acabSkuDigitado()">
          <div class="field-hint">Linha + traço + sigla da cor — o mesmo SKU que a OS calcula (Linha SKU do modelo + Sigla SKU da cor).</div></div>
        <div class="field"><label>Linha</label><input type="text" id="acab-p-linha" value="${esc(base.linha)}" placeholder="CM.LISA" oninput="this.dataset.mexido=1"></div>
        <div class="field"><label>Cor</label><input type="text" id="acab-p-cor" value="${esc(base.cor)}" placeholder="PRE" oninput="this.dataset.mexido=1"></div>
        <div class="field"><label>Código do ERP</label><input type="number" id="acab-p-codigo" min="1" step="1" value="${esc(base.codigo || '')}" placeholder="Ex.: 12"></div>
        <div class="field"><label>Preço base (R$)</label><input type="text" inputmode="decimal" id="acab-p-preco" value="${base.preco !== '' && base.preco != null ? esc(String(base.preco).replace('.', ',')) : ''}" placeholder="0,00"></div>
        <div class="field full"><label>Descrição</label><input type="text" id="acab-p-desc" value="${esc(base.descricao || '')}" placeholder="Ex.: CAMISETA BÁSICA PRETO"></div>
        <div class="field full"><label>Tamanhos *</label><div style="display:flex;flex-wrap:wrap;gap:6px 12px;">${TAMS.map(t =>
          `<label style="display:inline-flex;gap:4px;align-items:center;font-weight:400;"><input type="checkbox" class="acab-p-tam" value="${t}"${sel.has(t) ? ' checked' : ''}>${t}</label>`).join('')}</div></div>
      </div>`;
    openModal('modal-acab-produto');
    setTimeout(() => { const e = document.getElementById('acab-p-sku'); if (e) e.focus(); }, 50);
  };
  window.acabSkuDigitado = function () {
    const pt = partesSku(val('acab-p-sku'));
    const l = document.getElementById('acab-p-linha'), c = document.getElementById('acab-p-cor');
    if (l && !l.dataset.mexido) l.value = pt.linha;
    if (c && !c.dataset.mexido) c.value = pt.cor;
  };

  window.acabSalvarProduto = async function () {
    if (!exigirEdicao('cadastrar SKU no estoque de produtos')) return;
    const sku = norm(val('acab-p-sku'));
    if (!sku) return toast('Informe o SKU', 'err');
    const outro = produtoPorSku(sku);
    if (outro && outro.id !== _editProdId) return toast('Já existe o SKU ' + sku, 'err');
    const codigo = val('acab-p-codigo') ? parseInt(val('acab-p-codigo'), 10) : null;
    if (codigo != null && !(codigo > 0)) return toast('Código inválido', 'err');
    const comCod = codigo ? produtoPorCodigo(codigo) : null;
    if (comCod && comCod.id !== _editProdId) return toast(`O código ${codigo} já é do SKU ${comCod.sku}`, 'err');
    const preco = val('acab-p-preco').trim() ? num(val('acab-p-preco')) : 0;
    if (!(preco >= 0)) return toast('Preço inválido', 'err');
    const tamanhos = ordenarTams([...document.querySelectorAll('.acab-p-tam:checked')].map(x => x.value));
    if (!tamanhos.length) return toast('Marque ao menos um tamanho', 'err');
    const campos = { codigo, sku, linha: norm(val('acab-p-linha')), cor: norm(val('acab-p-cor')), descricao: val('acab-p-desc').trim(), tamanhos, preco,
      atualizadoEm: new Date().toISOString(), atualizadoPor: quemSou() };
    let p = _editProdId ? produtoPorId(_editProdId) : null;
    if (p) p = fixarProduto(p) || p;
    if (p) Object.assign(p, campos);
    else cadastro().push(Object.assign({ id: uid(), criadoEm: campos.atualizadoEm, criadoPor: campos.atualizadoPor }, campos));
    esquecerProdutos();
    desfazerNomearAcao(`${p ? 'edição' : 'cadastro'} do SKU ${sku}`);
    await saveState('acabProdutos');
    closeModal('modal-acab-produto');
    toast(`SKU ${sku} ${p ? 'salvo' : 'cadastrado'}`, 'ok');
    _editProdId = null;
    renderAtual();
  };

  window.acabExcluirProduto = async function (id) {
    if (!exigirEdicao('cadastrar SKU no estoque de produtos')) return;
    const p = produtoPorId(id);
    if (!p) return;
    const nMov = todosMovs().filter(m => m.produtoId === id).length;
    if (nMov) return toast(`${p.sku} tem ${nMov} lançamento(s): apague-os antes (ou deixe o SKU sem uso)`, 'err');
    if (p.doPrograma) return toast(`${p.sku} vem do programa (desenho técnico, modelo ou OS): para tirá-lo daqui, mude o SKU lá`, 'err');
    if (!confirm(`Apagar o SKU ${p.sku}?`)) return;
    STATE.acabProdutos = cadastro().filter(x => x.id !== id);
    esquecerProdutos();
    desfazerNomearAcao(`SKU ${p.sku} apagado`);
    await saveState('acabProdutos');
    toast('SKU apagado', 'ok');
    renderProdutos();
  };

  /* ================================================ TELA: SAÍDAS DO ERP */
  /* O relatório de vendas do ERP vira saídas, como no original: código do
     produto, variação (tamanho) e quantidade. Lê PDF (pdf.js, que o programa
     já carrega para os riscos) e CSV; a planilha do Excel entra salva como CSV.
     O que muda do original: produto que não existe NÃO nasce com 300 peças de
     contagem (era o semeio do Render) — só é cadastrado se a caixa pedir, e o
     estoque dele começa pela contagem de inventário. E a venda não é barrada
     por falta de saldo: ela já aconteceu; a linha fica negativa, em vermelho,
     até a contagem acertar. */
  let _erp = null;   // { rows, dataEmissao, arquivo }

  function extrairTipoCor(desc) {
    const s = String(desc || '').trim();
    const m = s.match(/\s+(PP|P|M|G|GG|G1|G2|G3|\d{1,2})-([A-Z][A-Z0-9.]*?)-([A-Z][A-Z0-9]*)\s*$/);
    if (m) return { tipo: m[2].toUpperCase(), cor: m[3].toUpperCase(), descricao: s.slice(0, m.index).trim() };
    return { tipo: '', cor: '', descricao: s };
  }
  function dataBR(v) {
    const s = String(v || '').trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (!m) return '';
    let y = m[3]; if (y.length === 2) y = (parseInt(y, 10) >= 70 ? '19' : '20') + y;
    return y + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  }

  async function lerPdfErp(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('o leitor de PDF não carregou');
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const linhas = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const tc = await (await pdf.getPage(i)).getTextContent();
      const itens = tc.items.map(it => ({ x: it.transform[4], y: -it.transform[5], s: it.str }))
        .filter(o => o.s && o.s.trim()).sort((a, b) => (a.y - b.y) || (a.x - b.x));
      const grupos = [];
      let atual = null;
      itens.forEach(it => {
        if (!atual || Math.abs(it.y - atual.y) > 3) { atual = { y: it.y, partes: [] }; grupos.push(atual); }
        atual.partes.push(it);
      });
      grupos.forEach(g => { const l = g.partes.sort((a, b) => a.x - b.x).map(o => o.s).join(' ').replace(/\s+/g, ' ').trim(); if (l) linhas.push(l); });
    }
    let dataEmissao = '';
    for (let i = 0; i < linhas.length && !dataEmissao; i++) {
      const m = linhas[i].match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
      if (m && /EMISS/i.test((linhas[i - 1] || '') + ' ' + linhas[i])) dataEmissao = dataBR(m[1]);
    }
    if (!dataEmissao) for (const l of linhas) { const m = l.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/); if (m && (dataEmissao = dataBR(m[1]))) break; }
    // cód + descrição + qtd + tamanho + custo [+ total] — às vezes quebrado em duas linhas.
    const re = /^\s*(\d{1,4})\s+(.+?)\s+([\d.]+,\d+)\s+(PP|P|M|G|GG|G1|G2|G3|\d{1,2})\s+([\d.]+,\d+)(?:\s+([\d.]+,\d+))?\s*$/;
    const rows = [];
    for (let i = 0; i < linhas.length; i++) {
      let m = linhas[i].match(re);
      if (!m && i + 1 < linhas.length) { m = (linhas[i] + ' ' + linhas[i + 1]).match(re); if (m) i++; }
      if (!m) continue;
      const tc = extrairTipoCor(m[2]);
      rows.push({ codigo: parseInt(m[1], 10), tamanho: m[4], quantidade: parseInt(m[3].split(',')[0].replace(/\./g, ''), 10),
        tipo: tc.tipo, cor: tc.cor, descricao: tc.descricao || m[2], preco: num(m[5]) || 0 });
    }
    return { rows, dataEmissao };
  }

  function lerCsvErp(texto) {
    const linhas = String(texto).replace(/\r/g, '').split('\n').filter(l => l.trim());
    if (!linhas.length) return { rows: [], dataEmissao: '' };
    const sep = (linhas[0].match(/;/g) || []).length > (linhas[0].match(/,/g) || []).length ? ';' : ',';
    const corta = l => { const out = []; let cur = '', q = false; for (let i = 0; i < l.length; i++) { const c = l[i]; if (c === '"') { if (q && l[i + 1] === '"') { cur += '"'; i++; } else q = !q; } else if (c === sep && !q) { out.push(cur); cur = ''; } else cur += c; } out.push(cur); return out; };
    const cab = corta(linhas[0]).map(s => norm(s).toLowerCase());
    const idx = (...ns) => { for (const n of ns) { const i = cab.findIndex(h => h.includes(n)); if (i >= 0) return i; } return -1; };
    const iCod = idx('cod.produto', 'cod produto', 'codigo', 'cod'), iSku = idx('sku'), iDesc = idx('descri', 'item', 'produto');
    const iQtd = idx('quant', 'qtd', 'qtde'), iVar = idx('varia', 'tamanho', 'tam'), iCusto = idx('custo', 'valor unit');
    if ((iCod < 0 && iSku < 0) || iQtd < 0 || iVar < 0) throw new Error('cabeçalho não reconhecido: precisa de código (ou SKU), quantidade e variação/tamanho');
    const rows = [];
    linhas.slice(1).forEach(l => {
      const c = corta(l);
      const codigo = iCod >= 0 ? parseInt(String(c[iCod] || '').replace(/\D/g, ''), 10) : NaN;
      const desc = iDesc >= 0 ? String(c[iDesc] || '').trim() : '';
      const tc = extrairTipoCor(desc);
      const skuCol = iSku >= 0 ? norm(c[iSku]) : '';
      const pt = skuCol ? partesSku(skuCol) : null;
      if (!(codigo > 0) && !skuCol) return;
      rows.push({ codigo: codigo > 0 ? codigo : null, tamanho: String(c[iVar] || '').trim().toUpperCase(),
        quantidade: parseInt(String(c[iQtd] || '').split(',')[0].replace(/\D/g, ''), 10) || 0,
        tipo: pt ? pt.linha : tc.tipo, cor: pt ? pt.cor : tc.cor, descricao: tc.descricao || desc,
        preco: iCusto >= 0 ? (num(c[iCusto]) || 0) : 0 });
    });
    return { rows, dataEmissao: '' };
  }

  // A prévia: casa cada linha com um SKU (código do ERP, senão linha-cor) e
  // soma as linhas repetidas do mesmo SKU e tamanho.
  function planoErp(rows, cadastrarFaltantes) {
    const r = { saidas: [], novos: [], codigosNovos: [], tamanhosNovos: [], problemas: [], agregadas: 0, zeradas: 0 };
    const mapa = new Map();
    const novosPorSku = new Map();
    rows.forEach((row, i) => {
      const linha = i + 1;
      const tamanho = asTam(row.tamanho);
      const q = parseInt(row.quantidade, 10);
      if (!tamanho) { r.problemas.push(`Linha ${linha}: tamanho inválido (${row.tamanho})`); return; }
      if (!(q >= 0)) { r.problemas.push(`Linha ${linha}: quantidade inválida`); return; }
      if (q === 0) { r.zeradas++; return; }
      const skuTxt = row.tipo && row.cor ? norm(row.tipo + '-' + row.cor) : '';
      let p = produtoPorCodigo(row.codigo) || (skuTxt ? produtoPorSku(skuTxt) : null);
      if (!p && skuTxt && novosPorSku.has(skuTxt)) p = novosPorSku.get(skuTxt);
      if (!p) {
        if (!skuTxt) { r.problemas.push(`Linha ${linha}: código ${row.codigo || '?'} sem SKU cadastrado e a descrição não diz a linha e a cor ("${row.descricao}")`); return; }
        if (!cadastrarFaltantes) { r.problemas.push(`Linha ${linha}: SKU ${skuTxt} (código ${row.codigo || '?'}) não está cadastrado`); return; }
        p = { id: uid(), codigo: row.codigo || null, sku: skuTxt, linha: norm(row.tipo), cor: norm(row.cor), descricao: row.descricao || '',
          tamanhos: tamsDaLinha(row.tipo), preco: Number(row.preco) || 0, criadoEm: new Date().toISOString(), criadoPor: quemSou(), origem: 'erp' };
        novosPorSku.set(skuTxt, p);
        r.novos.push(p);
      }
      if (row.codigo && !p.codigo && !r.novos.includes(p) && !r.codigosNovos.some(x => x.p === p)) r.codigosNovos.push({ p, codigo: row.codigo });
      if (!(p.tamanhos || []).includes(tamanho) && !r.tamanhosNovos.some(x => x.p === p && x.t === tamanho)) r.tamanhosNovos.push({ p, t: tamanho });
      const k = p.id + '|' + tamanho;
      if (mapa.has(k)) { mapa.get(k).qtd += q; r.agregadas++; }
      else mapa.set(k, { p, tamanho, qtd: q });
    });
    r.saidas = [...mapa.values()];
    return r;
  }

  function renderErp() {
    const box = document.getElementById('acab-erp-painel');
    if (!box) return;
    const hist = movs().filter(m => m.origem === 'erp').sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.criadoEm || '').localeCompare(String(a.criadoEm || ''))).slice(0, 150);
    let previa = '';
    if (_erp) {
      const cad = !!(document.getElementById('acab-erp-cadastrar') || {}).checked || !!_erp.cadastrar;
      const pl = planoErp(_erp.rows, cad);
      const data = _erp.data || _erp.dataEmissao || hoje();
      const jaTem = movs().some(m => m.origem === 'erp' && m.data === data);
      const S = saldos(todosMovs());
      const negativos = pl.saidas.filter(s => S.de(s.p.id, s.tamanho).total - s.qtd < 0).length;
      previa = `<div class="card">
        <h2 style="margin-top:0;font-size:15px;">Prévia · ${esc(_erp.arquivo)}</h2>
        <div class="form-grid cols-2">
          <div class="field"><label>Data das saídas *</label><input type="date" id="acab-erp-data" value="${esc(data)}" onchange="acabErpData(this.value)"></div>
          <div class="field"><label>Emissão do relatório</label><input type="text" value="${_erp.dataEmissao ? fmtD(_erp.dataEmissao) : '—'}" disabled></div>
          <div class="field full"><label style="font-weight:400;display:flex;gap:6px;align-items:center;"><input type="checkbox" id="acab-erp-cadastrar" ${cad ? 'checked' : ''} onchange="acabErpCadastrar(this.checked)"> Cadastrar os SKUs que ainda não existem (pela linha e cor da descrição)</label></div>
        </div>
        <div class="exp-resumo" style="margin:8px 0;">
          <div class="item"><div class="num">${fmtN(pl.saidas.length)}</div><div class="lbl">Saídas a lançar</div></div>
          <div class="item"><div class="num">${fmtN(pl.saidas.reduce((s, x) => s + x.qtd, 0))}</div><div class="lbl">Peças</div></div>
          <div class="item"><div class="num">${fmtN(pl.novos.length)}</div><div class="lbl">SKUs novos</div></div>
          <div class="item${pl.problemas.length ? ' alerta' : ''}"><div class="num">${fmtN(pl.problemas.length)}</div><div class="lbl">Linhas com problema</div></div>
          <div class="item${negativos ? ' alerta' : ''}"><div class="num">${fmtN(negativos)}</div><div class="lbl">Ficam negativos</div></div>
        </div>
        ${jaTem ? `<div class="info-box" style="margin-bottom:8px;background:#fbeccb;">Já existem saídas do ERP com a data ${fmtD(data)}. Confira se este relatório não é o mesmo de antes.</div>` : ''}
        ${pl.agregadas ? `<div class="muted" style="font-size:12px;">${fmtN(pl.agregadas)} linha(s) do mesmo SKU e tamanho foram somadas.</div>` : ''}
        ${pl.problemas.length ? `<details style="margin:8px 0;"><summary>Ver as ${fmtN(pl.problemas.length)} linhas com problema</summary><div style="max-height:220px;overflow:auto;font-size:12px;">${pl.problemas.slice(0, 300).map(x => `<div style="padding:3px 0;border-bottom:1px solid var(--line-2);">${esc(x)}</div>`).join('')}</div></details>` : ''}
        <div style="overflow:auto;max-height:320px;margin:8px 0;"><table class="table"><thead><tr><th>SKU</th><th>Tam.</th><th style="text-align:right;">Qtd</th><th style="text-align:right;">Estoque depois</th></tr></thead><tbody>
          ${pl.saidas.map(s => { const dep = S.de(s.p.id, s.tamanho).total - s.qtd; return `<tr><td>${esc(rotuloProduto(s.p))}${pl.novos.includes(s.p) ? ' <span class="badge" style="background:#d6f0db;">novo</span>' : ''}</td><td>${esc(s.tamanho)}</td><td style="${mono}">${fmtN(s.qtd)}</td><td style="${mono}${dep < 0 ? 'color:#c0392b;font-weight:700;' : ''}">${fmtN(dep)}</td></tr>`; }).join('') || '<tr><td colspan="4" class="empty">Nada a lançar.</td></tr>'}
        </tbody></table></div>
        <button class="btn primary" onclick="acabErpConfirmar()" ${pl.saidas.length ? '' : 'disabled'}>Lançar ${fmtN(pl.saidas.length)} saídas</button>
        <button class="btn" onclick="acabErpCancelar()">Cancelar</button>
      </div>`;
    }
    box.innerHTML = `
      <div class="card admin-only">
        <h2 style="margin-top:0;font-size:15px;">Importar o relatório de vendas</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">O relatório do ERP em <b>PDF</b> ou <b>CSV</b> (do Excel: Salvar como → CSV). O programa lê o <b>código do produto</b>, a <b>variação</b> (tamanho) e a <b>quantidade</b>, e mostra a prévia antes de lançar. Cada linha vira uma saída "Estoque → Cliente" com valor R$ 0,00 — o custo sai pelo custo médio.</div>
        <input type="file" id="acab-erp-arquivo" accept=".pdf,.csv,.txt" onchange="acabErpArquivo(this.files[0])">
      </div>
      ${previa}
      ${(() => {
        // A loja do Mercado Livre, da Shopee ou do TikTok ligada em Vendas das
        // plataformas já dá baixa sozinha: importar o relatório dela de novo
        // baixaria a mesma venda duas vezes.
        const ls = typeof window.vendasDreLojasDoErp === 'function' ? window.vendasDreLojasDoErp() : [];
        return ls.length ? `<div class="info-box" style="margin-bottom:12px;background:#fbeccb;"><b>Atenção:</b> estas lojas já dão baixa sozinhas em <a href="#acab-dre" onclick="goto('acab-dre');return false;">Vendas das plataformas</a> — não importe as vendas delas aqui a partir da data de corte: ${ls.map(l => `<b>${esc(l.nome)}</b> (desde ${fmtD(l.corte)})`).join(', ')}.</div>` : '';
      })()}
      <div class="card"><h2 style="margin-top:0;font-size:15px;">Saídas que vieram do ERP</h2>
        <div style="overflow:auto;max-height:440px;"><table class="table"><thead><tr><th>Data</th><th>SKU</th><th>Tam.</th><th style="text-align:right;">Qtd</th><th>Relatório</th><th class="col-actions admin-only"></th></tr></thead><tbody>
        ${hist.map(m => `<tr><td>${fmtD(m.data)}</td><td>${esc(rotuloProduto(produtoPorId(m.produtoId)))}</td><td>${esc(m.tamanho)}</td><td style="${mono}">${fmtN(m.qtd)}</td><td class="muted" style="font-size:11px;">${m.erpEmissao ? 'emitido ' + fmtD(m.erpEmissao) : ''} ${esc(m.erpArquivo || '')}</td><td class="col-actions row-actions admin-only"><button onclick="acabApagarMov('${esc(m.id)}')">apagar</button></td></tr>`).join('') || '<tr><td colspan="6" class="empty">Nenhuma saída importada.</td></tr>'}
        </tbody></table></div></div>`;
  }

  window.acabErpArquivo = async function (file) {
    if (!file) return;
    if (!exigirEdicao('importar saídas do ERP')) return;
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    try {
      const r = ext === 'pdf' ? await lerPdfErp(file) : lerCsvErp(await file.text());
      if (!r.rows.length) return toast('Nenhuma linha de produto encontrada no arquivo', 'err');
      _erp = { rows: r.rows, dataEmissao: r.dataEmissao || '', data: r.dataEmissao || hoje(), arquivo: file.name, cadastrar: false };
      toast(`${r.rows.length} linhas lidas — confira a prévia`, 'ok');
      renderErp();
    } catch (e) {
      console.warn('acab erp', e);
      toast('Não deu para ler o arquivo: ' + (e.message || e), 'err');
    }
  };
  window.acabErpData = v => { if (_erp) _erp.data = v; renderErp(); };
  window.acabErpCadastrar = v => { if (_erp) _erp.cadastrar = !!v; renderErp(); };
  window.acabErpCancelar = () => { _erp = null; renderErp(); };
  window.acabErpConfirmar = async function () {
    if (!_erp || !exigirEdicao('importar saídas do ERP') || !exigirInicio()) return;
    const data = _erp.data;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data || '')) return toast('Informe a data das saídas', 'err');
    const pl = planoErp(_erp.rows, !!_erp.cadastrar);
    if (!pl.saidas.length) return toast('Nada a lançar', 'err');
    if (!confirm(`Lançar ${pl.saidas.length} saídas (${pl.saidas.reduce((s, x) => s + x.qtd, 0)} peças) com a data ${fmtD(data)}?`)) return;
    const nCadastro = cadastro().length + pl.novos.length;
    pl.novos.forEach(p => cadastro().push(p));
    // O SKU do programa que ganha código, tamanho ou saída do ERP fica gravado aqui.
    const fixos = new Map();
    const fixo = p => { if (!fixos.has(p.id)) fixos.set(p.id, fixarProduto(p) || p); return fixos.get(p.id); };
    pl.codigosNovos.forEach(({ p, codigo }) => { if (!produtoPorCodigo(codigo)) fixo(p).codigo = codigo; });
    pl.tamanhosNovos.forEach(({ p, t }) => { const q = fixo(p); q.tamanhos = ordenarTams((q.tamanhos || []).concat([t])); });
    pl.saidas.forEach(s => fixo(s.p));
    esquecerProdutos();
    pl.saidas.forEach(s => movs().push(novoMov({ tipo: 'saida', origem: 'erp', produtoId: s.p.id, tamanho: s.tamanho, data, qtd: s.qtd, valor: 0,
      destino: 'cliente', erpEmissao: _erp.dataEmissao || null, erpArquivo: _erp.arquivo })));
    desfazerNomearAcao(`importação de ${pl.saidas.length} saídas do ERP`);
    if (pl.novos.length || pl.codigosNovos.length || pl.tamanhosNovos.length || cadastro().length !== nCadastro) await saveState('acabProdutos');
    await saveState('acabMov');
    toast(`${pl.saidas.length} saídas do ERP lançadas`, 'ok');
    _erp = null;
    renderErp();
  };

  /* ================================================ TELA: RELATÓRIO CMP */
  const _cmpSel = { produto: '', tam: '' };
  function renderCmp() {
    const box = document.getElementById('acab-cmp-painel');
    if (!box) return;
    if (!produtos().length) { box.innerHTML = `<div class="card">${semSkus()}</div>`; return; }
    const p = produtoPorId(_cmpSel.produto);
    const ms = todosMovs();
    const r = p ? cmp(ms, p.id, _cmpSel.tam || null) : null;
    box.innerHTML = `<div class="card">
      <div class="muted" style="font-size:12px;margin-bottom:8px;">A cada entrada o custo médio é refeito como a média ponderada entre o estoque e o que entrou; a saída sai pelo custo médio do momento e não o muda. As entradas das OS entram com custo R$ 0,00 — o custo de produção fica na Contabilidade.</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;">
        <select onchange="acabCmpSel('produto',this.value)" style="flex:2;min-width:220px;">${opcoesProdutos(p && p.id)}</select>
        <select onchange="acabCmpSel('tam',this.value)" style="flex:1;min-width:120px;"><option value="">Todos os tamanhos</option>${p ? ordenarTams(p.tamanhos || []).map(t => `<option value="${t}"${t === _cmpSel.tam ? ' selected' : ''}>${t}</option>`).join('') : ''}</select>
      </div>
      ${!r ? '<div class="info-box">Escolha um SKU.</div>' : `
      <div class="exp-resumo" style="margin-bottom:10px;">
        <div class="item"><div class="num">${fmtN(r.saldo)}</div><div class="lbl">Saldo (entradas − saídas)</div></div>
        <div class="item"><div class="num">${fmtR(r.medio)}</div><div class="lbl">Custo médio</div></div>
        <div class="item"><div class="num">${fmtR(r.valor)}</div><div class="lbl">Valor do saldo</div></div>
        <div class="item"><div class="num">${fmtR(r.cmv)}</div><div class="lbl">CMV acumulado</div></div>
      </div>
      <div style="overflow:auto;max-height:60vh;"><table class="table"><thead><tr><th>Data</th><th>Movimento</th><th style="text-align:right;">Qtd ent.</th><th style="text-align:right;">Custo ent.</th><th style="text-align:right;">Total ent.</th><th style="text-align:right;">Qtd saí.</th><th style="text-align:right;">Custo saí.</th><th style="text-align:right;">Total saí.</th><th style="text-align:right;">Saldo</th><th style="text-align:right;">Custo médio</th></tr></thead><tbody>
      ${r.linhas.map(l => `<tr><td>${fmtD(l.m.data)}</td><td>${l.qtdEnt != null ? 'Entrada' + (l.m.origem === 'os' ? ' · OS ' + esc(l.m.os) : '') : 'Saída → ' + esc(((DESTINOS.find(d => d.k === l.m.destino) || {}).rotulo || '').replace('Estoque → ', ''))} [${esc(l.m.tamanho)}]</td>
        <td style="${mono}">${l.qtdEnt != null ? fmtN(l.qtdEnt) : '—'}</td><td style="${mono}">${l.qtdEnt != null ? fmtR(l.custoEnt) : '—'}</td><td style="${mono}">${l.qtdEnt != null ? fmtR(l.totalEnt) : '—'}</td>
        <td style="${mono}">${l.qtdSai != null ? fmtN(l.qtdSai) : '—'}</td><td style="${mono}">${l.qtdSai != null ? fmtR(l.custoSai) : '—'}</td><td style="${mono}">${l.qtdSai != null ? fmtR(l.totalSai) : '—'}</td>
        <td style="${mono}font-weight:700;">${fmtN(l.saldo)}</td><td style="${mono}">${fmtR(l.medio)}</td></tr>`).join('') || '<tr><td colspan="10" class="empty">Sem entradas nem saídas.</td></tr>'}
      </tbody></table></div>`}
    </div>`;
  }
  window.acabCmpSel = (k, v) => { _cmpSel[k] = v; if (k === 'produto') _cmpSel.tam = ''; renderCmp(); };

  /* ------------------------------------------------------- a navegação */
  const TELAS = { 'acab-estoque': renderEstoqueAtual, 'acab-mov': renderMovimentos, 'acab-inventario': renderInventario,
    'acab-produtos': renderProdutos, 'acab-erp': renderErp, 'acab-cmp': renderCmp,
    // Vendas das plataformas: a tela mora no vendas-dre.js.
    'acab-dre': () => { if (typeof window.vendasDreRender === 'function') window.vendasDreRender(); } };
  function renderAtual() {
    const s = document.querySelector('section.page:not(.hidden)');
    const page = s && s.dataset && s.dataset.page;
    if (TELAS[page]) TELAS[page]();
  }
  window.renderEstoqueConfeccao = function (page) {
    const fn = TELAS[page];
    if (!fn) return;
    try { fn(); } catch (e) { console.warn('Estoque de produtos', page, e); }
  };

  /* OS SKUs QUE O PROGRAMA DEFINE (22/10/2026; desde 23/10/2026 lidos na
     hora, ver produtos() no topo — não há mais cópia nem botão de sincronizar):
       · o SKU completo do desenho técnico e do modelo (skuLinha com traço,
         "CM.LISA-PRE");
       · o de cada OS (skusDaOS: linha do modelo/desenho + sigla da cor), que é
         o mesmo que a OS usa para lançar a entrada quando chega ao Estoque.
     Só a combinação que existe entra — a linha sozinha (CJM) não é cruzada
     com as 50 cores. A descrição vem do catálogo quando ele tem o SKU; senão,
     do desenho, do modelo ou da OS.
     O DESENHO SEM SKU COMPLETO (28/10/2026, Jones: "atualização automática do
     Estoque de produtos, para sincronizar novos SKU quando forem cadastrados"):
     o desenho novo nasce com o SKU vazio ou só com a linha ("CJM"), e o
     estoque só o via quando saía a primeira OS dele. Quem completava era
     revalidarSkusDesenhos, que depende do catálogo skus_catalogo — tabela que
     o servidor da fábrica não tem. Agora o desenho compõe o SKU como a OS
     compõe: a linha (a do desenho, senão a do modelo) mais a Sigla SKU da cor
     principal. Cadastrou o desenho (ou deu a sigla à cor), o SKU aparece.
     O SKU DO MODELO, DIRETO DA FONTE (28/10/2026, Jones: "a atualização
     automática dos SKU novos em Estoque de produtos não funciona. O estoque
     de produtos deve ler os cadastros de SKU direto da fonte"). Os modelos
     novos (Short tactel SM.LISA, Short feminino SF.BERM.SEM, Bermuda BERM,
     Polo Tricolor PM.TRI) têm só a linha, sem desenho nem OS com cor — e a
     regra acima ("a linha sozinha não entra") os deixava fora. Agora o SKU
     cadastrado no modelo (ou no desenho que não compõe cor) entra como está,
     enquanto nenhuma variante com cor dessa linha existir. Quando o desenho
     ou a OS trouxer o "SM.LISA-PRE", a linha sozinha sai da lista (a menos
     que já tenha lançamento, código ou preço próprios — aí está no cadastro).
     Continua sem cruzar a linha com as cores. */
  function skuDoDesenho(d) {
    const proprio = norm(d && d.skuLinha);
    if (proprio.includes('-')) return proprio;
    const m = (STATE.modelos || []).find(x => x.id === (d && d.modeloId));
    const linha = proprio || norm(m && m.skuLinha);
    if (!linha || linha.includes('-')) return '';
    const c = (STATE.cores || []).find(x => x.id === (d && d.corPrincipalId));
    const sigla = norm(c && c.siglaSku);
    return sigla ? linha + '-' + sigla : '';
  }
  /* A MINIATURA DO DESENHO TÉCNICO (28/10/2026, Jones: "Insira no estoque de
     produtos coluna com miniatura do desenho técnico dos produtos"). A imagem
     é a do desenho cujo SKU é o do produto (skuDoDesenho); sem desenho com
     esse SKU, a do desenho da OS que compõe o SKU; para a linha sozinha (o
     modelo novo), a de um desenho dessa linha. Uma conta por tela, não por
     linha da tabela. */
  function imagensDosSkus() {
    const mapa = new Map();
    const poe = (sku, d) => { const s = norm(sku); if (s && d && d.img && !mapa.has(s)) mapa.set(s, { img: d.img, desc: d.desc || '' }); };
    (STATE.desenhos || []).forEach(d => poe(skuDoDesenho(d), d));
    (STATE.ordens || []).forEach(o => {
      const d = (STATE.desenhos || []).find(x => x.id === o.desenhoId) || o.desenhoFoto;
      if (d && d.img) (typeof skusDaOS === 'function' ? skusDaOS(o) : []).forEach(s => poe(s, d));
    });
    (STATE.desenhos || []).forEach(d => { const s = skuDoDesenho(d); if (s) poe(s.split('-')[0], d); });
    return mapa;
  }
  const _urlImg = v => (typeof urlDesenho === 'function' ? urlDesenho(v) : String(v || ''));
  function celMiniatura(imgs, p) {
    const s = norm(p && p.sku);
    const m = imgs.get(s) || imgs.get(s.split('-')[0]);
    return `<td style="width:56px;padding:3px 6px;"><div style="width:48px;height:36px;background:#f5f2ea;display:flex;align-items:center;justify-content:center;border:1px solid var(--line);overflow:hidden;">${m
      ? `<img src="${esc(_urlImg(m.img))}" loading="lazy" alt="" title="${esc(m.desc)}" style="max-width:100%;max-height:100%;object-fit:contain;">`
      : '<span class="muted" style="font-size:10px;">—</span>'}</div></td>`;
  }
  function skusDoPrograma() {
    const achados = new Map();
    const poe = (sku, descricao) => {
      const s = norm(sku);
      if (!/^[A-Z0-9][A-Z0-9.]*-[A-Z0-9.]+$/.test(s) || achados.has(s)) return;
      achados.set(s, descricao || '');
    };
    (STATE.desenhos || []).forEach(d => { const s = skuDoDesenho(d); if (s) poe(s, d.desc); });
    (STATE.modelos || []).forEach(m => { if (String(m.skuLinha || '').includes('-')) poe(m.skuLinha, m.nome); });
    (STATE.ordens || []).forEach(o => {
      if (typeof _statusOS === 'function' && _statusOS(o) === 'cancelado') return;
      (typeof skusDaOS === 'function' ? skusDaOS(o) : []).forEach(s => poe(s, o.modeloNome));
    });
    // A linha sem cor (o modelo novo): entra até haver variante com cor dela.
    const linhas = new Map();
    const poeLinha = (sku, descricao) => {
      const s = norm(sku);
      if (/^[A-Z0-9][A-Z0-9.]*$/.test(s) && !linhas.has(s)) linhas.set(s, descricao || '');
    };
    (STATE.modelos || []).forEach(m => poeLinha(m.skuLinha, m.nome));
    (STATE.desenhos || []).forEach(d => { if (!skuDoDesenho(d)) poeLinha(d.skuLinha, d.desc); });
    const comCor = new Set([...achados.keys()].concat(cadastro().map(p => norm(p.sku))).map(s => s.split('-')[0]));
    linhas.forEach((desc, linha) => { if (!comCor.has(linha) && !achados.has(linha)) achados.set(linha, desc); });
    return achados;
  }

  /* OS AVISOS DO SINO (21/10/2026): a coluna Avisar da área "Estoque
     confecção" em Configurações. Três notícias, montadas aqui e desenhadas pelo
     painel de avisos do app.js (_avisosEventos):
       · acab-os   — a OS entrou no estoque (a entrada foi lançada sozinha);
       · acab-pend — a OS entrou no estoque mas espera SKU: fica no sino
                     enquanto estiver pendente, como o planejamento à espera;
       · acab-erp  — saídas importadas do ERP, uma notícia por importação.
     'quem' das OS é vazio: o carimbo da etapa não guarda quem marcou. */
  function acabAvisos(limiteMs) {
    const ev = [];
    const r = _entradasDasOS();
    const porOS = new Map();
    r.entradas.forEach(e => {
      const g = porOS.get(e.osId) || { e, pecas: 0 };
      g.pecas += Number(e.qtd) || 0;
      porOS.set(e.osId, g);
    });
    porOS.forEach(({ e, pecas }) => {
      if (Date.parse(e.criadoEm) < limiteMs) return;
      ev.push({ tipo: 'acab-os', osId: e.osId, os: e.os, quem: '', em: e.criadoEm, texto: e.sku + ' · ' + fmtN(pecas) + ' peças' });
    });
    r.pendencias.forEach(p => ev.push({ tipo: 'acab-pend', osId: p.osId, os: p.os, quem: '', em: p.em, texto: p.motivo }));
    const lotes = new Map();
    movs().filter(m => m.origem === 'erp' && m.criadoEm && Date.parse(m.criadoEm) >= limiteMs).forEach(m => {
      const k = (m.erpArquivo || '') + '|' + String(m.criadoEm).slice(0, 16);
      const g = lotes.get(k) || { m, n: 0, pecas: 0 };
      g.n++; g.pecas += Number(m.qtd) || 0;
      lotes.set(k, g);
    });
    lotes.forEach(({ m, n, pecas }) => ev.push({ tipo: 'acab-erp', osId: '', os: '', quem: m.criadoPor || '', em: m.criadoEm,
      dataOe: m.data, n, texto: fmtN(n) + ' saídas · ' + fmtN(pecas) + ' peças' + (m.erpArquivo ? ' · ' + m.erpArquivo : '') }));
    return ev;
  }
  window.acabAvisos = acabAvisos;

  // Para os testes (node): as contas, sem a tela. todosMovs e skuDoDesenho
  // também servem ao planejamento-automatico.js (o saldo de cada SKU e
  // tamanho, e o desenho de cada SKU).
  window._acabInterno = { acabAvisos, _entradasDasOS, saldos, cmp, planoErp, lerCsvErp, extrairTipoCor, acabSnapshotContab, cfg, produtos, skusDoPrograma, limiteDe, statusDe, imagensDosSkus, celMiniatura, todosMovs, skuDoDesenho,
    // e estes, ao vendas-dre.js (a baixa das vendas das plataformas).
    produtoPorSku, fixarProduto, esquecerProdutos, TAMS };
})();
