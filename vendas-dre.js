/* =====================================================================
   VENDAS DAS PLATAFORMAS — a baixa do Estoque de produtos pelo Painel DRE
   (30/10/2026, Jones: prompt-gerador-os-vendas-dre.md)
   =====================================================================
   O Painel DRE (painel-dre.onrender.com) já está ligado às lojas da Shopify,
   do Mercado Livre, da Shopee e do TikTok Shop e publica uma rota só de
   leitura para o Gerador-OS. Nenhuma chave de plataforma mora aqui: só a
   CHAVE DE LIGAÇÃO com o DRE, digitada pelo admin na tela e guardada nos dados
   (STATE.meta.vendasDre.chave) — o repositório é público, e ela não pode ir
   para o código.

   O QUE ACONTECE A CADA RODADA (5 em 5 minutos, num computador só):
     1. pede ao DRE os pedidos ALTERADOS desde o cursor, página por página,
        até `maisPaginas` acabar (o cursor é guardado como veio — ele tem
        microssegundos e não pode passar por Date);
     2. fica só com o pedido de LOJA LIGADA e com `data` (dia de São Paulo) na
        data de corte da loja ou depois. O DRE regrava pedidos antigos o tempo
        todo; o resto é ignorado sem lançar nada;
     3. para cada pedido, a quantidade que DEVE estar fora do estoque, por SKU
        e tamanho (faturado = soma dos itens; não faturado = 0), comparada com
        o que JÁ FOI lançado para aquele pedido. Só a diferença é lançada:
        positiva vira saída "Estoque → Cliente"; negativa (cancelado ou
        devolvido depois) é estorno.

   O QUE JÁ FOI LANÇADO NÃO TEM CADERNO PRÓPRIO: é a soma dos movimentos do
   próprio pedido em STATE.acabMov (origem 'dre', pedidoChave). O prompt
   sugeria um caderno em meta.vendasDre.pedidos — não foi feito assim por dois
   motivos. (1) Caderno e movimentos são duas chaves que podiam divergir (uma
   gravada e a outra não), e aí o pedido baixava duas vezes ou nunca; lendo do
   próprio movimento, o que está lançado é, por definição, o que o caderno
   diria. (2) O `meta` sobe inteiro a cada configuração miúda de qualquer
   pessoa (acessos, fila de produção), e um caderno com milhares de pedidos
   ia junto em cada uma.
   Por isso o ESTORNO é tirar do movimento do pedido (e apagá-lo quando zera):
   a venda que não aconteceu some do Relatório CMP em vez de virar uma
   entrada que não é peça voltando.
   E apagar à mão um movimento do DRE DISPENSA o pedido (meta.vendasDre.
   dispensados): sem isso a rodada seguinte o lançaria de novo.

   NÃO SAI MAIS DO QUE HÁ no tamanho: lança o que der e o resto fica em
   Pendências (falta de estoque). O pedido com falta, com SKU não reconhecido
   ou sem tamanho fica guardado em meta.vendasDre.pendentes e é refeito a cada
   rodada — a falta se resolve sozinha quando a peça entra, e o SKU quando
   alguém cria a regra "De → Para".

   UM COMPUTADOR POR VEZ: a trava meta.vendasDre.trava = { quem, aba, ate }.
   Pegar a trava é gravar nela; e só se trabalha depois de uma leitura do
   servidor confirmar que ela continua nossa (ESPERA_MS) — se dois pegaram ao
   mesmo tempo, o último a gravar fica com ela e o outro desiste ao ler. Ela
   vale TRAVA_MS (7 min) e é renovada a cada rodada (5 min): assim não vence no
   intervalo entre duas rodadas, e se o computador fechar, outro assume em até
   7 minutos. Roda quem é admin ou tem a área Estoque de produtos, com a aba
   ativa (o polling para quando a aba fica escondida).

   As gravações da rodada não entram no Desfazer nem no diário: o Ctrl+Z da
   pessoa tem de voltar o que ELA fez, e não a venda que o robô baixou.
   ===================================================================== */
(function () {
  'use strict';

  const BASE = 'https://painel-dre.onrender.com';
  const INTERVALO_MS = 5 * 60 * 1000;
  const TRAVA_MS = 7 * 60 * 1000;
  const ESPERA_MS = 40 * 1000;             // duas leituras do polling (15 s cada)
  const PRIMEIRA_MS = 30 * 1000;
  const TIMEOUT_MS = 90 * 1000;            // o Render gratuito leva ~1 min para acordar
  const LIMITE = 1000;
  const MAX_PAGINAS = 20;                  // por rodada; o resto fica para a próxima
  const MAX_PENDENTES = 3000;
  const DIAS_ESTAT = 15;
  const PLAT_COM_ERP = ['mercadolivre', 'shopee', 'tiktok'];
  const ROTULO_PLAT = { shopify: 'Shopify', mercadolivre: 'Mercado Livre', shopee: 'Shopee', tiktok: 'TikTok Shop' };
  const TAMS_LETRA = ['PP', 'P', 'M', 'G', 'GG', 'G1', 'G2', 'G3'];
  // A linha pela sigla do tipo (o meio do SKU da plataforma). A ordem importa:
  // o específico (CM.TRI) antes do genérico (CM.).
  const LINHAS = [['CM.TRI', 'CM.TRI'], ['CM.REC', 'CM.REC'], ['CM.', 'CM.LISA'],
    ['PM.TRI', 'PM.TRI'], ['PM.', 'PM.LISA'], ['BM.TRI', 'BM.TRI'], ['BM.', 'BM.LISA']];

  const ABA = 'aba-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  const _abertoEm = Date.now();
  let _ultima = 0;
  let _rodando = false;
  let _pegouEm = 0;
  let _integracoes = null;      // a última lista de /integracoes (memória da aba)
  let _integracoesErro = '';
  let _regraForm = null;        // a regra sendo criada ou editada na tela

  /* ---------------------------------------------------------------- base */
  const A = () => window._acabInterno;
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase();
  const quemSou = () => (typeof _obsQuemSou === 'function' ? _obsQuemSou() : '');
  const fmtN = v => Math.round(Number(v) || 0).toLocaleString('pt-BR');
  const fmtDH = iso => { const d = new Date(iso); return isNaN(d) ? '—' : d.toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' }); };
  const fmtD = dia => (dia ? (typeof formatDate === 'function' ? formatDate(String(dia).slice(0, 10)) : dia) : '—');
  const val = id => ((document.getElementById(id) || {}).value || '');
  const mono = "text-align:right;font-family:'IBM Plex Mono',monospace;white-space:nowrap;";

  // O dia (AAAA-MM-DD) de um instante no fuso de São Paulo.
  function diaSP(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return '';
    return d.toLocaleDateString('en-CA', { timeZone: 'America/Sao_Paulo' });
  }
  const hojeSP = () => diaSP(new Date().toISOString());
  // O começo de um dia de São Paulo, em ISO (o cursor que NÓS montamos ao
  // ligar uma loja; o que vem do DRE é guardado como veio).
  const inicioDoDia = dia => new Date(dia + 'T00:00:00-03:00').toISOString();

  function dre() {
    STATE.meta = STATE.meta || {};
    const c = STATE.meta.vendasDre && typeof STATE.meta.vendasDre === 'object' ? STATE.meta.vendasDre : (STATE.meta.vendasDre = {});
    if (!c.lojas || typeof c.lojas !== 'object') c.lojas = {};
    if (!Array.isArray(c.mapa)) c.mapa = [];
    if (!c.pendentes || typeof c.pendentes !== 'object') c.pendentes = {};
    if (!c.dispensados || typeof c.dispensados !== 'object') c.dispensados = {};
    if (!c.estat || typeof c.estat !== 'object') c.estat = {};
    if (!c.status || typeof c.status !== 'object') c.status = {};
    return c;
  }

  // Grava sem entrar no Desfazer nem no diário (ver o cabeçalho).
  async function gravar(chaves) {
    let antes = false;
    try { antes = _desfazendo; _desfazendo = true; } catch (e) { }
    try { for (const k of chaves) await saveState(k); }
    finally { try { _desfazendo = antes; } catch (e) { } }
  }

  async function api(caminho, chave) {
    const ctl = typeof AbortController === 'function' ? new AbortController() : null;
    const t = ctl ? setTimeout(() => ctl.abort(), TIMEOUT_MS) : null;
    try {
      const r = await fetch(BASE + caminho, { headers: { 'x-chave': chave }, signal: ctl ? ctl.signal : undefined });
      let corpo = null;
      try { corpo = await r.json(); } catch (e) { }
      if (r.status === 401) throw new Error('chave de ligação recusada pelo Painel DRE' + (corpo && corpo.erro ? ' (' + corpo.erro + ')' : ''));
      if (!r.ok) throw new Error('o Painel DRE respondeu ' + r.status + (corpo && corpo.erro ? ': ' + corpo.erro : ''));
      if (!corpo) throw new Error('resposta do Painel DRE fora do formato');
      return corpo;
    } catch (e) {
      if (e && e.name === 'AbortError') throw new Error('o Painel DRE não respondeu em ' + Math.round(TIMEOUT_MS / 1000) + ' s');
      if (/failed to fetch|networkerror|load failed/i.test(String(e && e.message))) throw new Error('sem conexão com o Painel DRE (internet da fábrica?)');
      throw e;
    } finally { if (t) clearTimeout(t); }
  }

  /* ------------------------------------- o SKU da plataforma → o daqui */
  const linhaExiste = meio => A().produtos().some(p => norm(p.sku).startsWith(meio + '-') || norm(p.linha) === meio);
  function linhaDe(meio) {
    if (!meio) return '';
    if (linhaExiste(meio)) return meio;            // o meio já é um SKU-base cadastrado
    for (const [pre, linha] of LINHAS) {
      const generico = pre.endsWith('.');
      if (generico ? meio.startsWith(pre) : (meio === pre || meio.startsWith(pre + '.'))) return linha;
    }
    return '';
  }
  // Tamanho escrito na descrição (o item sem SKU): a última palavra que é um
  // tamanho de letra; número só depois de "TAM"/"TAMANHO" ("KIT 2" não é 2).
  function tamDaDescricao(desc) {
    const ps = norm(desc).split(/[^A-Z0-9]+/).filter(Boolean);
    for (let i = ps.length - 1; i >= 0; i--) {
      if (TAMS_LETRA.includes(ps[i])) return ps[i];
      if (/^\d{1,2}$/.test(ps[i]) && /^TAM/.test(ps[i - 1] || '') && A().TAMS.includes(ps[i])) return ps[i];
    }
    return '';
  }
  // A regra "De → Para" que casa com o item: por SKU (sem o tamanho, ou pelo
  // começo) e, para o item sem SKU, pela descrição. A mais longa ganha.
  function acharRegra(mapa, skuP, resto, desc) {
    let melhor = null;
    (mapa || []).forEach(r => {
      const de = norm(r && r.de);
      if (!de) return;
      let casa = false;
      if (r.tipo === 'descricao') casa = !!desc && desc.includes(de);
      else {
        // O tamanho que o estoque não tem ("U-BRINDE", tamanho único) também
        // sai da frente: a regra casa com o que vem depois do primeiro traço.
        const semPrimeiro = skuP.includes('-') ? skuP.slice(skuP.indexOf('-') + 1) : '';
        casa = !!skuP && [resto, skuP, semPrimeiro].some(x => x && (x === de || x.startsWith(de)));
      }
      if (!casa) return;
      const peso = (r.tipo === 'descricao' ? 0 : 10000) + de.length;   // SKU vence descrição
      if (!melhor || peso > melhor.peso) melhor = { r, peso };
    });
    return melhor ? melhor.r : null;
  }
  /* Um item do pedido → [{ produto, tamanho, qtd }] por unidade vendida, ou
     { ignorar } ou { erro }. */
  function converter(item, mapa) {
    const skuP = norm(item && item.sku);
    const desc = norm(item && item.descricao);
    let tam = '', resto = skuP;
    const i = skuP.indexOf('-');
    if (i > 0 && A().TAMS.includes(skuP.slice(0, i))) { tam = skuP.slice(0, i); resto = skuP.slice(i + 1); }
    const regra = acharRegra(mapa, skuP, resto, desc);
    if (regra) {
      if (regra.ignorar) return { ignorar: true, regra };
      const t = tam || norm(regra.tam) || tamDaDescricao(desc);
      if (!t) return { erro: 'sem tamanho (a regra "' + regra.de + '" não diz o tamanho e o item não traz)', regra };
      const partes = [];
      for (const x of regra.para || []) {
        const p = A().produtoPorSku(x.sku);
        if (!p) return { erro: 'a regra "' + regra.de + '" aponta para ' + norm(x.sku) + ', que não está cadastrado', regra };
        partes.push({ produto: p, tamanho: t, qtd: Math.max(1, Math.round(Number(x.qtd) || 1)) });
      }
      if (!partes.length) return { erro: 'a regra "' + regra.de + '" não tem destino', regra };
      return { partes, regra };
    }
    if (!skuP) return { erro: 'item sem SKU (crie a regra pela descrição)' };
    if (!tam) return { erro: 'o SKU não começa por um tamanho' };
    const j = resto.lastIndexOf('-');
    if (j <= 0) return { erro: 'SKU fora do padrão TAMANHO-TIPO-COR' };
    const meio = resto.slice(0, j), cor = resto.slice(j + 1);
    const linha = linhaDe(meio);
    if (!linha) return { erro: 'o tipo ' + meio + ' não é de nenhuma linha conhecida' };
    const p = A().produtoPorSku(linha + '-' + cor);
    if (!p) return { erro: linha + '-' + cor + ' não está cadastrado' };
    return { partes: [{ produto: p, tamanho: tam, qtd: 1 }] };
  }

  /* --------------------------------------------------------- a baixa */
  const compacto = p => ({ chave: p.chave, id: p.id, integracao_id: p.integracao_id, plataforma: p.plataforma, loja: p.loja,
    numero: p.numero, status: p.status, faturado: !!p.faturado, data: p.data,
    itens: (p.itens || []).map(x => ({ sku: x.sku || '', descricao: x.descricao || '', quantidade: Number(x.quantidade) || 0 })) });

  function relevante(p, c) {
    const l = c.lojas[String(p && p.integracao_id)];
    return !!(l && l.ligada && l.corte && p.chave && diaSP(p.data) >= l.corte);
  }

  /* Lança a diferença de cada pedido. Síncrona de propósito: lê e escreve o
     STATE de agora, sem nenhum await no meio (o polling pode refazer o STATE
     enquanto se espera a rede). Devolve o resumo da rodada. */
  function processar(pedidos, c) {
    const r = { saidas: 0, pecas: 0, estornos: 0, faltas: 0, desconhecidos: 0, fixados: false, resolvidos: 0 };
    const acab = A();
    if (!Array.isArray(STATE.acabMov)) STATE.acabMov = [];
    const S = acab.saldos(acab.todosMovs());
    const delta = new Map();
    const disp = (pid, t) => S.de(pid, t).total + (delta.get(pid + '|' + t) || 0);
    const doPedido = new Map();
    STATE.acabMov.forEach(m => {
      if (m.origem !== 'dre' || !m.pedidoChave) return;
      if (!doPedido.has(m.pedidoChave)) doPedido.set(m.pedidoChave, []);
      doPedido.get(m.pedidoChave).push(m);
    });
    const apagar = new Set();
    const agora = new Date().toISOString();
    pedidos.forEach(p => {
      if (c.dispensados[p.chave]) { delete c.pendentes[p.chave]; return; }
      if (!relevante(p, c)) { delete c.pendentes[p.chave]; return; }
      const desejado = new Map();
      const problemas = [];
      if (p.faturado) (p.itens || []).forEach(item => {
        const q = Math.round(Number(item.quantidade) || 0);
        if (q <= 0) return;
        const cv = converter(item, c.mapa);
        if (cv.ignorar) return;
        if (cv.erro) { problemas.push({ tipo: 'sku', sku: item.sku || '', descricao: item.descricao || '', qtd: q, motivo: cv.erro }); r.desconhecidos++; return; }
        cv.partes.forEach(x => {
          const k = x.produto.id + '|' + x.tamanho;
          const d = desejado.get(k) || { produto: x.produto, tamanho: x.tamanho, qtd: 0 };
          d.qtd += x.qtd * q;
          desejado.set(k, d);
        });
      });
      const lancados = (doPedido.get(p.chave) || []).filter(m => !apagar.has(m.id));
      const porK = new Map();
      lancados.forEach(m => { const k = m.produtoId + '|' + m.tamanho; if (!porK.has(k)) porK.set(k, []); porK.get(k).push(m); });
      new Set([...desejado.keys(), ...porK.keys()]).forEach(k => {
        const quer = (desejado.get(k) || {}).qtd || 0;
        const ms = porK.get(k) || [];
        const tem = ms.reduce((s, m) => s + (Number(m.qtd) || 0), 0);
        const [pid, tam] = [k.slice(0, k.lastIndexOf('|')), k.slice(k.lastIndexOf('|') + 1)];
        if (quer > tem) {
          const falta = quer - tem;
          const q = Math.min(falta, Math.max(0, disp(pid, tam)));
          if (q > 0) {
            let prod = desejado.get(k).produto;
            const fixo = acab.fixarProduto(prod);
            if (fixo) { prod = fixo; r.fixados = true; }
            const m0 = ms[0];
            if (m0) { m0.qtd = (Number(m0.qtd) || 0) + q; m0.atualizadoEm = agora; }
            else {
              const mov = { id: uid(), tipo: 'saida', origem: 'dre', produtoId: prod.id, tamanho: tam, data: diaSP(p.data) || hojeSP(),
                qtd: q, valor: 0, destino: 'cliente', pedidoChave: p.chave, pedidoNumero: p.numero == null ? '' : String(p.numero),
                loja: p.loja || '', plataforma: p.plataforma || '', criadoEm: agora, criadoPor: quemSou() };
              STATE.acabMov.push(mov);
              if (!doPedido.has(p.chave)) doPedido.set(p.chave, []);
              doPedido.get(p.chave).push(mov);
            }
            delta.set(k, (delta.get(k) || 0) - q);
            r.saidas++; r.pecas += q;
          }
          if (q < falta) {
            const prod = desejado.get(k).produto;
            problemas.push({ tipo: 'falta', sku: prod.sku, tamanho: tam, qtd: falta - q, motivo: 'falta estoque: há ' + fmtN(Math.max(0, disp(pid, tam) + q)) + ' e o pedido pede mais ' + fmtN(falta) });
            r.faltas++;
          }
        } else if (quer < tem) {
          let sobra = tem - quer;
          ms.slice().sort((a, b) => String(b.criadoEm || '').localeCompare(String(a.criadoEm || ''))).forEach(m => {
            if (sobra <= 0) return;
            const tira = Math.min(sobra, Number(m.qtd) || 0);
            m.qtd = (Number(m.qtd) || 0) - tira;
            m.atualizadoEm = agora;
            if (m.qtd <= 0) apagar.add(m.id);
            sobra -= tira;
            r.estornos += tira;
            delta.set(k, (delta.get(k) || 0) + tira);
          });
        }
      });
      if (problemas.length) c.pendentes[p.chave] = Object.assign(compacto(p), { problemas, em: agora });
      else if (c.pendentes[p.chave]) { delete c.pendentes[p.chave]; r.resolvidos++; }
    });
    if (apagar.size) STATE.acabMov = STATE.acabMov.filter(m => !apagar.has(m.id));
    // Teto das pendências: as mais antigas saem (o aviso fica no status).
    const chaves = Object.keys(c.pendentes);
    if (chaves.length > MAX_PENDENTES) {
      chaves.sort((a, b) => String(c.pendentes[a].data).localeCompare(String(c.pendentes[b].data)))
        .slice(0, chaves.length - MAX_PENDENTES).forEach(k => delete c.pendentes[k]);
      r.cortadas = chaves.length - MAX_PENDENTES;
    }
    if (r.fixados && acab.esquecerProdutos) acab.esquecerProdutos();
    return r;
  }

  function anotarEstat(c, r, lidos, relevantes) {
    const dia = hojeSP();
    const e = c.estat[dia] || { lidos: 0, relevantes: 0, pecas: 0, estornos: 0 };
    e.lidos += lidos; e.relevantes += relevantes; e.pecas += r.pecas; e.estornos += r.estornos;
    c.estat[dia] = e;
    Object.keys(c.estat).sort().slice(0, -DIAS_ESTAT).forEach(k => delete c.estat[k]);
  }

  const lojasLigadas = c => Object.keys(c.lojas).filter(id => c.lojas[id].ligada && c.lojas[id].corte);

  /* Uma rodada inteira: rede primeiro (todas as páginas), conta depois, numa
     tacada só (ver processar), e grava. */
  async function rodar() {
    if (_rodando) return;
    _rodando = true;
    _ultima = Date.now();
    let c = dre();
    try {
      c.trava = { quem: quemSou(), aba: ABA, ate: new Date(Date.now() + TRAVA_MS).toISOString() };
      if (!A().cfg().inicio) throw new Error('inicie a contagem do estoque primeiro (botão no Estoque atual)');
      const ligadas = lojasLigadas(c);
      if (!ligadas.length) {
        c.status = Object.assign({}, c.status, { ultimaEm: new Date().toISOString(), erro: '', aviso: 'nenhuma loja ligada' });
        await gravar(['meta']);
        return;
      }
      const chave = c.chave;
      let cursor = c.cursor && c.cursor.desde ? Object.assign({}, c.cursor)
        : { desde: inicioDoDia(ligadas.map(id => c.lojas[id].corte).sort()[0]), apos: '' };
      const chegaram = new Map();
      let lidos = 0, paginas = 0, mais = false;
      do {
        const q = '?desde=' + encodeURIComponent(cursor.desde) + (cursor.apos ? '&apos=' + encodeURIComponent(cursor.apos) : '') + '&limite=' + LIMITE;
        const resp = await api('/api/gerador-os/pedidos' + q, chave);
        const lista = Array.isArray(resp.pedidos) ? resp.pedidos : [];
        lidos += lista.length;
        lista.forEach(p => { if (p && p.chave) chegaram.set(p.chave, p); });
        mais = !!resp.maisPaginas;
        const prox = resp.cursor && resp.cursor.desde ? { desde: resp.cursor.desde, apos: resp.cursor.apos == null ? '' : resp.cursor.apos } : null;
        if (!prox) break;
        if (mais && prox.desde === cursor.desde && String(prox.apos) === String(cursor.apos)) throw new Error('o cursor do Painel DRE não andou (' + cursor.desde + ') — parei para não ficar em laço');
        cursor = prox;
        paginas++;
      } while (mais && paginas < MAX_PAGINAS);
      // A rede acabou: daqui até gravar, nada de await (ver processar).
      c = dre();
      const ligadoAinda = c.trava && c.trava.aba === ABA;
      if (!ligadoAinda) throw new Error('outro computador assumiu a sincronização no meio da rodada — nada foi lançado');
      const novos = [...chegaram.values()].filter(p => relevante(p, c)).map(compacto);
      const fila = new Map(Object.values(c.pendentes).map(p => [p.chave, p]));
      novos.forEach(p => fila.set(p.chave, p));
      const r = processar([...fila.values()], c);
      c.cursor = cursor;
      anotarEstat(c, r, lidos, novos.length);
      c.status = { ultimaEm: new Date().toISOString(), quem: quemSou(), erro: '', aviso: mais ? 'ainda há páginas: continua na próxima rodada' : '',
        lidos, relevantes: novos.length, pecas: r.pecas, estornos: r.estornos, pendentes: Object.keys(c.pendentes).length,
        cortadas: r.cortadas || 0 };
      c.trava = { quem: quemSou(), aba: ABA, ate: new Date(Date.now() + TRAVA_MS).toISOString() };
      const chaves = [];
      if (r.fixados) chaves.push('acabProdutos');
      if (r.saidas || r.estornos) chaves.push('acabMov');
      chaves.push('meta');
      await gravar(chaves);
      if (r.pecas || r.estornos) toast(`Vendas das plataformas: ${fmtN(r.pecas)} peça(s) baixada(s)${r.estornos ? ', ' + fmtN(r.estornos) + ' estornada(s)' : ''}`, 'ok');
    } catch (e) {
      console.warn('vendas DRE', e);
      c = dre();
      c.status = Object.assign({}, c.status, { ultimaEm: new Date().toISOString(), quem: quemSou(), erro: String(e && e.message || e) });
      try { await gravar(['meta']); } catch (e2) { }
    } finally {
      _rodando = false;
      redesenharSeAberta();
    }
  }

  /* ------------------------------------------------- quem roda, e quando */
  function podeRodar() {
    if (typeof currentUser === 'undefined' || !currentUser) return false;
    if (typeof podeGravar === 'function' && !podeGravar()) return false;
    if (!(STATE.meta && STATE.meta.vendasDre && STATE.meta.vendasDre.chave)) return false;
    if (!A() || !A().produtoPorSku) return false;
    if (currentRole === 'admin') return true;
    return typeof contaTemAcesso === 'function' && contaTemAcesso('estoque-confeccao', quemSou());
  }
  const travaDeOutro = c => !!(c.trava && c.trava.aba !== ABA && Date.parse(c.trava.ate) > Date.now());

  // Chamada pelo polling (15 s). Pega a trava, espera uma leitura confirmar e
  // roda de 5 em 5 minutos.
  window.vendasDreRodar = async function (forcar) {
    if (_rodando || !podeRodar()) return;
    if (!forcar && Date.now() - _abertoEm < PRIMEIRA_MS) return;
    const c = dre();
    if (travaDeOutro(c)) return;
    const minha = c.trava && c.trava.aba === ABA && Date.parse(c.trava.ate) > Date.now();
    if (!minha) {
      c.trava = { quem: quemSou(), aba: ABA, ate: new Date(Date.now() + TRAVA_MS).toISOString() };
      _pegouEm = Date.now();
      await gravar(['meta']);
      if (!forcar) return;
    }
    if (!forcar && Date.now() - _pegouEm < ESPERA_MS) return;
    if (!forcar && Date.now() - _ultima < INTERVALO_MS) return;
    await rodar();
  };

  window.vendasDreSincronizar = async function () {
    if (!exigirEdicao('lançar no estoque de produtos')) return;
    const c = dre();
    if (!c.chave) return toast('Falta a chave de ligação com o Painel DRE', 'err');
    if (travaDeOutro(c)) return toast(`Quem sincroniza agora é ${c.trava.quem || 'outro computador'} (até ${fmtDH(c.trava.ate)})`, 'err');
    if (!podeRodar()) return toast('Esta conta não sincroniza (precisa ser admin ou ter a área Estoque de produtos)', 'err');
    toast('Sincronizando com o Painel DRE…', 'ok');
    await window.vendasDreRodar(true);
  };

  /* --------------------------------------------------------- a tela */
  function redesenharSeAberta() {
    const s = document.querySelector && document.querySelector('section.page:not(.hidden)');
    if (s && s.dataset && s.dataset.page === 'acab-dre') render();
  }

  const mascarar = k => (k ? '•'.repeat(Math.max(4, Math.min(12, k.length - 4))) + String(k).slice(-4) : '');
  const ehAdmin = () => typeof currentRole !== 'undefined' && currentRole === 'admin';
  const podeEditar = () => ehAdmin() || (typeof contaTemAcesso === 'function' && contaTemAcesso('estoque-confeccao', quemSou()));

  async function carregarIntegracoes() {
    const c = dre();
    if (!c.chave) return;
    try {
      const r = await api('/api/gerador-os/integracoes', c.chave);
      _integracoes = Array.isArray(r.integracoes) ? r.integracoes : [];
      _integracoesErro = '';
    } catch (e) { _integracoesErro = String(e.message || e); }
    redesenharSeAberta();
  }

  function render() {
    const box = document.getElementById('acab-dre-painel');
    if (!box) return;
    const c = dre();
    const st = c.status || {};
    if (c.chave && !_integracoes && !_integracoesErro) { _integracoes = []; carregarIntegracoes(); }
    const lojasApi = _integracoes || [];
    const ids = new Set(lojasApi.map(l => String(l.id)).concat(Object.keys(c.lojas)));
    const lojas = [...ids].map(id => {
      const a = lojasApi.find(l => String(l.id) === id) || {};
      const g = c.lojas[id] || {};
      return { id, nome: a.nome || g.nome || id, plataforma: a.plataforma || g.plataforma || '', empresa: a.empresa || g.empresa || '',
        status: a.status || '', ultima_sync: a.ultima_sync || '', ligada: !!g.ligada, corte: g.corte || '' };
    }).sort((a, b) => String(a.empresa + a.nome).localeCompare(String(b.empresa + b.nome), 'pt-BR'));
    const ed = podeEditar();
    const trava = c.trava && Date.parse(c.trava.ate) > Date.now() ? c.trava : null;

    // Resumo: as saídas do DRE por dia do pedido (hoje e 7 dias).
    const hoje = hojeSP();
    const semana = diaSP(new Date(Date.now() - 6 * 86400000).toISOString());
    const dres = (STATE.acabMov || []).filter(m => m.origem === 'dre');
    const somar = desde => {
      const mapa = new Map();
      let pecas = 0;
      const pedidos = new Set();
      dres.filter(m => String(m.data) >= desde).forEach(m => {
        const k = m.produtoId + '|' + m.tamanho;
        mapa.set(k, (mapa.get(k) || 0) + (Number(m.qtd) || 0));
        pecas += Number(m.qtd) || 0;
        pedidos.add(m.pedidoChave);
      });
      return { mapa, pecas, pedidos: pedidos.size };
    };
    const rH = somar(hoje), rS = somar(semana);
    const est = (desde) => Object.keys(c.estat).filter(d => d >= desde).reduce((s, d) => ({ lidos: s.lidos + (c.estat[d].lidos || 0), estornos: s.estornos + (c.estat[d].estornos || 0) }), { lidos: 0, estornos: 0 });
    const eH = est(hoje), eS = est(semana);
    const rot = pid => { const p = A().produtos().find(x => x.id === pid); return p ? p.sku : '(SKU apagado)'; };
    const tabelaSkus = mapa => [...mapa.entries()].sort((a, b) => b[1] - a[1]).slice(0, 60).map(([k, q]) => {
      const i = k.lastIndexOf('|');
      return `<tr><td>${esc(rot(k.slice(0, i)))}</td><td>${esc(k.slice(i + 1))}</td><td style="${mono}">${fmtN(q)}</td></tr>`;
    }).join('') || '<tr><td colspan="3" class="empty">Nenhuma peça baixada.</td></tr>';

    // Pendências.
    const pend = Object.values(c.pendentes);
    const skusDesc = new Map();
    const faltas = [];
    pend.forEach(p => (p.problemas || []).forEach(x => {
      if (x.tipo === 'falta') { faltas.push({ p, x }); return; }
      const de = x.sku ? (norm(x.sku).replace(/^(PP|P|M|G|GG|G1|G2|G3|\d{1,2})-/, '')) : norm(x.descricao);
      const k = (x.sku ? 'sku:' : 'desc:') + de;
      const g = skusDesc.get(k) || { tipo: x.sku ? 'sku' : 'descricao', de, exemplo: x.sku || '', descricao: x.descricao, qtd: 0, lojas: new Set(), motivo: x.motivo, pedidos: 0 };
      g.qtd += Number(x.qtd) || 0; g.lojas.add(p.loja || ''); g.pedidos++;
      skusDesc.set(k, g);
    }));

    const f = _regraForm;
    box.innerHTML = `
      <div class="card">
        <h2 style="margin-top:0;font-size:15px;">Ligação com o Painel DRE</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">As vendas da Shopify, do Mercado Livre, da Shopee e do TikTok Shop chegam pelo Painel DRE, que já está ligado às lojas. Aqui só mora a <b>chave de ligação</b> com ele (nenhuma chave de plataforma). A sincronização roda sozinha a cada 5 minutos, num computador por vez (admin ou área Estoque de produtos, com o programa aberto).</div>
        <div class="form-grid cols-2">
          <div class="field"><label>Chave de ligação</label>${ehAdmin()
            ? `<input type="password" id="dre-chave" autocomplete="off" placeholder="${c.chave ? esc(mascarar(c.chave)) + ' (gravada — digite para trocar)' : 'Cole a chave de ligação'}">`
            : `<input type="text" value="${c.chave ? esc(mascarar(c.chave)) : 'não configurada (só o admin)'}" disabled>`}</div>
          <div class="field" style="display:flex;gap:6px;align-items:flex-end;flex-wrap:wrap;">
            ${ehAdmin() ? `<button class="btn primary" onclick="vendasDreSalvarChave()">Gravar chave</button>` : ''}
            <button class="btn" onclick="vendasDreTestar()" ${c.chave ? '' : 'disabled'}>Testar ligação</button>
            ${ed ? `<button class="btn" onclick="vendasDreSincronizar()" ${c.chave ? '' : 'disabled'}>Sincronizar agora</button>` : ''}
          </div>
        </div>
        <div class="info-box" style="margin-top:8px;font-size:12px;${st.erro ? 'background:#f6dcda;' : ''}">
          ${_rodando ? '<b>Sincronizando agora…</b> ' : ''}
          Última rodada: <b>${st.ultimaEm ? fmtDH(st.ultimaEm) : 'nenhuma'}</b>${st.quem ? ' por ' + esc(st.quem) : ''}
          ${st.erro ? `<div style="color:#c0392b;"><b>Erro:</b> ${esc(st.erro)}</div>` : ''}
          ${st.aviso ? `<div>${esc(st.aviso)}</div>` : ''}
          ${st.ultimaEm && !st.erro ? `<div>${fmtN(st.lidos)} pedido(s) lido(s), ${fmtN(st.relevantes)} de loja ligada · ${fmtN(st.pecas)} peça(s) baixada(s) · ${fmtN(st.estornos)} estornada(s) · ${fmtN(st.pendentes)} pedido(s) com pendência${st.cortadas ? ` · <b>${fmtN(st.cortadas)} pendência(s) antiga(s) descartada(s) pelo teto de ${fmtN(MAX_PENDENTES)}</b>` : ''}</div>` : ''}
          <div class="muted">Cursor: ${c.cursor && c.cursor.desde ? esc(c.cursor.desde) + (c.cursor.apos ? ' · ' + esc(c.cursor.apos) : '') : '—'} · ${trava ? 'quem sincroniza: <b>' + esc(trava.quem || '?') + '</b>' + (trava.aba === ABA ? ' (este computador)' : '') : 'ninguém sincronizando agora'}</div>
        </div>
      </div>

      <div class="card">
        <h2 style="margin-top:0;font-size:15px;">Lojas</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">Só a loja <b>ligada</b> dá baixa, e só dos pedidos com data na <b>data de corte</b> ou depois. Todas começam desligadas.</div>
        <div class="info-box" style="margin-bottom:8px;font-size:12px;background:#fbeccb;"><b>Atenção:</b> as vendas do Mercado Livre e da Shopee também entram pela tela <b>Saídas do ERP</b> (relatório do SoftUp). Ao ligar uma loja do Mercado Livre, da Shopee ou do TikTok aqui, <b>pare de importar o relatório do ERP dessa loja</b> a partir da data de corte — senão a venda sai duas vezes. A Shopify não passa pelo SoftUp.</div>
        ${_integracoesErro ? `<div class="info-box" style="background:#f6dcda;margin-bottom:8px;">Não deu para ler as lojas: ${esc(_integracoesErro)}</div>` : ''}
        <div style="overflow:auto;"><table class="table"><thead><tr><th>Loja</th><th>Plataforma</th><th>Empresa</th><th>No DRE</th><th>Ligada</th><th>Data de corte</th></tr></thead><tbody>
        ${lojas.map(l => `<tr><td><b>${esc(l.nome)}</b></td><td>${esc(ROTULO_PLAT[l.plataforma] || l.plataforma)}</td><td>${esc(l.empresa)}</td>
          <td class="muted" style="font-size:11px;">${esc(l.status)}${l.ultima_sync ? '<br>sync ' + esc(fmtDH(l.ultima_sync)) : ''}</td>
          <td><label style="cursor:pointer;"><input type="checkbox" ${l.ligada ? 'checked' : ''} ${ed ? '' : 'disabled'} onchange="vendasDreLigar('${esc(l.id)}',this.checked)"> ${l.ligada ? 'ligada' : 'desligada'}</label></td>
          <td><input type="date" value="${esc(l.corte)}" ${ed && l.ligada ? '' : 'disabled'} onchange="vendasDreCorte('${esc(l.id)}',this.value)"></td></tr>`).join('')
          || `<tr><td colspan="6" class="empty">${c.chave ? 'Carregando as lojas do Painel DRE…' : 'Grave a chave de ligação para ver as lojas.'}</td></tr>`}
        </tbody></table></div>
      </div>

      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px;">
        <div class="card"><h2 style="margin-top:0;font-size:15px;">Hoje</h2>
          <div class="exp-resumo" style="margin-bottom:8px;">
            <div class="item"><div class="num">${fmtN(eH.lidos)}</div><div class="lbl">Pedidos lidos</div></div>
            <div class="item"><div class="num">${fmtN(rH.pedidos)}</div><div class="lbl">Pedidos baixados</div></div>
            <div class="item"><div class="num">${fmtN(rH.pecas)}</div><div class="lbl">Peças baixadas</div></div>
            <div class="item"><div class="num">${fmtN(eH.estornos)}</div><div class="lbl">Estornos</div></div>
          </div>
          <div style="overflow:auto;max-height:300px;"><table class="table"><thead><tr><th>SKU</th><th>Tam.</th><th style="text-align:right;">Peças</th></tr></thead><tbody>${tabelaSkus(rH.mapa)}</tbody></table></div>
        </div>
        <div class="card"><h2 style="margin-top:0;font-size:15px;">Últimos 7 dias</h2>
          <div class="exp-resumo" style="margin-bottom:8px;">
            <div class="item"><div class="num">${fmtN(eS.lidos)}</div><div class="lbl">Pedidos lidos</div></div>
            <div class="item"><div class="num">${fmtN(rS.pedidos)}</div><div class="lbl">Pedidos baixados</div></div>
            <div class="item"><div class="num">${fmtN(rS.pecas)}</div><div class="lbl">Peças baixadas</div></div>
            <div class="item"><div class="num">${fmtN(eS.estornos)}</div><div class="lbl">Estornos</div></div>
          </div>
          <div style="overflow:auto;max-height:300px;"><table class="table"><thead><tr><th>SKU</th><th>Tam.</th><th style="text-align:right;">Peças</th></tr></thead><tbody>${tabelaSkus(rS.mapa)}</tbody></table></div>
        </div>
      </div>

      <div class="card">
        <h2 style="margin-top:0;font-size:15px;">Pendências</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">O que não foi lançado. Os pedidos ficam guardados e são refeitos a cada rodada: a falta de estoque se resolve quando a peça entra; o SKU não reconhecido, com uma regra "De → Para".</div>
        <h3 style="font-size:13px;margin:10px 0 6px;">SKUs não reconhecidos (${fmtN(skusDesc.size)})</h3>
        <div style="overflow:auto;max-height:360px;"><table class="table"><thead><tr><th>SKU / descrição</th><th>Motivo</th><th>Lojas</th><th style="text-align:right;">Pedidos</th><th style="text-align:right;">Peças</th><th class="col-actions"></th></tr></thead><tbody>
        ${[...skusDesc.values()].sort((a, b) => b.qtd - a.qtd).map(g => `<tr><td>${g.exemplo ? `<b>${esc(norm(g.exemplo))}</b><div class="muted" style="font-size:11px;">${esc(g.descricao)}</div>` : `<i>sem SKU</i> · ${esc(g.descricao)}`}</td>
          <td style="font-size:12px;">${esc(g.motivo)}</td><td style="font-size:12px;">${esc([...g.lojas].join(', '))}</td><td style="${mono}">${fmtN(g.pedidos)}</td><td style="${mono}">${fmtN(g.qtd)}</td>
          <td class="col-actions row-actions">${ed ? `<button onclick="vendasDreNovaRegra('${g.tipo}','${esc(encodeURIComponent(g.de))}')">criar regra</button>` : ''}</td></tr>`).join('') || '<tr><td colspan="6" class="empty">Nenhum.</td></tr>'}
        </tbody></table></div>
        <h3 style="font-size:13px;margin:14px 0 6px;">Faltas de estoque (${fmtN(faltas.length)})</h3>
        <div style="overflow:auto;max-height:300px;"><table class="table"><thead><tr><th>Pedido</th><th>Loja</th><th>Data</th><th>SKU</th><th>Tam.</th><th style="text-align:right;">Falta</th></tr></thead><tbody>
        ${faltas.sort((a, b) => String(b.p.data).localeCompare(String(a.p.data))).slice(0, 300).map(({ p, x }) => `<tr><td>${esc(p.numero)}</td><td>${esc(p.loja)}</td><td>${fmtD(diaSP(p.data))}</td><td>${esc(x.sku)}</td><td>${esc(x.tamanho)}</td><td style="${mono}">${fmtN(x.qtd)}</td></tr>`).join('') || '<tr><td colspan="6" class="empty">Nenhuma.</td></tr>'}
        </tbody></table></div>
      </div>

      <div class="card">
        <h2 style="margin-top:0;font-size:15px;">Tabela "De → Para"</h2>
        <div class="muted" style="font-size:12px;margin-bottom:8px;">Vence a regra automática (TAMANHO-TIPO.ESTAMPA-COR → linha-cor). <b>Por SKU</b>: casa com o SKU sem o tamanho, ou pelo começo dele. <b>Por descrição</b>: para o item sem SKU (hoje, os da Shopify) — casa quando a descrição contém o texto. Destino: um ou mais SKUs do estoque (kit: <code>CM.LISA-PRE x1, BM.LISA-PRE x1</code>) ou <b>ignorar</b> (não é produto do estoque).</div>
        ${ed ? `<div class="form-grid cols-2" style="margin-bottom:8px;">
          <div class="field"><label>Casar por</label><select id="dre-r-tipo"><option value="sku"${f && f.tipo === 'descricao' ? '' : ' selected'}>SKU (sem o tamanho, ou o começo)</option><option value="descricao"${f && f.tipo === 'descricao' ? ' selected' : ''}>Descrição (contém)</option></select></div>
          <div class="field"><label>De *</label><input type="text" id="dre-r-de" value="${esc(f ? f.de : '')}" placeholder="Ex.: CM.TRI.LA.MANGA-PRE ou KIT PARIS" style="text-transform:uppercase;"></div>
          <div class="field"><label>Para *</label><input type="text" id="dre-r-para" list="dre-skus" value="${esc(f ? f.para : '')}" placeholder="CM.LISA-PRE x1, BM.LISA-PRE x1 — ou: ignorar" style="text-transform:uppercase;"></div>
          <div class="field"><label>Tamanho (só se o item não trouxer)</label><input type="text" id="dre-r-tam" value="${esc(f ? f.tam || '' : '')}" placeholder="—" style="text-transform:uppercase;width:90px;"></div>
        </div>
        <datalist id="dre-skus">${A().produtos().map(p => `<option value="${esc(norm(p.sku))}">`).join('')}</datalist>
        <button class="btn primary" onclick="vendasDreGravarRegra()">${f && f.id ? 'Gravar alteração' : '+ Adicionar regra'}</button>
        ${f ? `<button class="btn" onclick="vendasDreCancelarRegra()">Cancelar</button>` : ''}` : ''}
        <div style="overflow:auto;max-height:360px;margin-top:10px;"><table class="table"><thead><tr><th>Por</th><th>De</th><th>Para</th><th>Tam.</th><th class="col-actions"></th></tr></thead><tbody>
        ${c.mapa.map(r => `<tr><td>${r.tipo === 'descricao' ? 'descrição' : 'SKU'}</td><td><b>${esc(r.de)}</b></td><td>${r.ignorar ? '<i>ignorar</i>' : esc((r.para || []).map(x => x.sku + ' x' + x.qtd).join(', '))}</td><td>${esc(r.tam || '')}</td>
          <td class="col-actions row-actions">${ed ? `<button onclick="vendasDreEditarRegra('${esc(r.id)}')">editar</button><button onclick="vendasDreApagarRegra('${esc(r.id)}')">apagar</button>` : ''}</td></tr>`).join('') || '<tr><td colspan="5" class="empty">Nenhuma regra.</td></tr>'}
        </tbody></table></div>
      </div>`;
  }
  window.vendasDreRender = render;

  /* --------------------------------------------------------- as ações */
  window.vendasDreSalvarChave = async function () {
    if (!ehAdmin()) return toast('Só o admin grava a chave de ligação', 'err');
    if (typeof _recusarSomenteLeitura === 'function' && _recusarSomenteLeitura('gravar a chave do Painel DRE')) return;
    const k = val('dre-chave').trim();
    if (!k) return toast('Cole a chave de ligação', 'err');
    try { await api('/api/gerador-os/integracoes', k); }
    catch (e) { if (!confirm(`O Painel DRE não aceitou esta chave agora (${e.message}).\n\nGravar assim mesmo?`)) return; }
    const c = dre();
    c.chave = k;
    c.chaveEm = new Date().toISOString(); c.chavePor = quemSou();
    await gravar(['meta']);
    toast('Chave de ligação gravada', 'ok');
    _integracoes = null; _integracoesErro = '';
    render();
  };

  window.vendasDreTestar = async function () {
    const c = dre();
    if (!c.chave) return toast('Falta a chave de ligação', 'err');
    toast('Testando a ligação (o Painel DRE pode levar até 1 minuto para acordar)…', 'ok');
    try {
      const r = await api('/api/gerador-os/integracoes', c.chave);
      _integracoes = Array.isArray(r.integracoes) ? r.integracoes : [];
      _integracoesErro = '';
      toast(`Ligação ok: ${_integracoes.length} loja(s) no Painel DRE`, 'ok');
    } catch (e) {
      _integracoesErro = String(e.message || e);
      toast('Falhou: ' + _integracoesErro, 'err');
    }
    render();
  };

  // Ligar (ou mudar a data de corte para trás) volta o cursor para o corte:
  // os pedidos desde então que o cursor já tinha passado são lidos de novo —
  // sem risco, porque o que já foi lançado de cada pedido é descontado.
  function recuarCursor(c, corte) {
    const ini = inicioDoDia(corte);
    if (!c.cursor || !c.cursor.desde || ini < String(c.cursor.desde)) c.cursor = { desde: ini, apos: '' };
  }

  window.vendasDreLigar = async function (id, ligar) {
    if (!exigirEdicao('configurar o estoque de produtos')) { render(); return; }
    const c = dre();
    const a = (_integracoes || []).find(l => String(l.id) === String(id)) || {};
    const g = Object.assign({}, c.lojas[id] || {});
    const plat = a.plataforma || g.plataforma || '';
    const nome = a.nome || g.nome || id;
    if (ligar && PLAT_COM_ERP.includes(plat) && !confirm(`Ligar ${nome} (${ROTULO_PLAT[plat] || plat})?\n\nA partir da data de corte (hoje), as vendas desta loja saem do estoque por aqui.\n\nPARE DE IMPORTAR O RELATÓRIO DO ERP (Saídas do ERP) PARA ESTA LOJA a partir de hoje — senão a mesma venda sai duas vezes.`)) { render(); return; }
    g.ligada = !!ligar;
    g.nome = nome; g.plataforma = plat; g.empresa = a.empresa || g.empresa || '';
    if (ligar) { g.corte = hojeSP(); recuarCursor(c, g.corte); }
    g.alteradoEm = new Date().toISOString(); g.alteradoPor = quemSou();
    c.lojas[id] = g;
    await gravar(['meta']);
    toast(ligar ? `${nome} ligada (corte ${fmtD(g.corte)})` : `${nome} desligada — os lançamentos já feitos continuam`, 'ok');
    render();
  };

  window.vendasDreCorte = async function (id, dia) {
    if (!exigirEdicao('configurar o estoque de produtos')) { render(); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dia || '')) { toast('Data de corte inválida', 'err'); render(); return; }
    const c = dre();
    const g = c.lojas[id];
    if (!g) return;
    if (dia < g.corte && PLAT_COM_ERP.includes(g.plataforma) && !confirm(`Recuar o corte de ${g.nome} para ${fmtD(dia)}?\n\nAs vendas desde essa data vão sair do estoque por aqui. Se o relatório do ERP desses dias já foi importado, elas saem duas vezes.`)) { render(); return; }
    g.corte = dia;
    g.alteradoEm = new Date().toISOString(); g.alteradoPor = quemSou();
    recuarCursor(c, dia);
    await gravar(['meta']);
    toast(`Corte de ${g.nome}: ${fmtD(dia)}`, 'ok');
    render();
  };

  // "CM.LISA-PRE x1, BM.LISA-PRE x2" → [{ sku, qtd }]
  function lerPara(txt) {
    const t = norm(txt);
    if (t === 'IGNORAR') return { ignorar: true };
    const para = [];
    for (const parte of t.split(/[,;+]/).map(s => s.trim()).filter(Boolean)) {
      const m = parte.match(/^([A-Z0-9][A-Z0-9.]*-[A-Z0-9.]+)(?:\s*[X×*]\s*(\d+))?$/);
      if (!m) return { erro: `"${parte}" não é SKU (linha-cor, ex.: CM.LISA-PRE x1)` };
      if (!A().produtoPorSku(m[1])) return { erro: `${m[1]} não está cadastrado no Estoque de produtos` };
      para.push({ sku: m[1], qtd: m[2] ? parseInt(m[2], 10) : 1 });
    }
    return para.length ? { para } : { erro: 'informe o destino' };
  }

  window.vendasDreNovaRegra = function (tipo, deCod) {
    _regraForm = { tipo: tipo === 'descricao' ? 'descricao' : 'sku', de: decodeURIComponent(deCod || ''), para: '', tam: '' };
    render();
    const el = document.getElementById('dre-r-para');
    if (el && el.scrollIntoView) { el.scrollIntoView({ block: 'center' }); el.focus(); }
  };
  window.vendasDreEditarRegra = function (id) {
    const r = dre().mapa.find(x => x.id === id);
    if (!r) return;
    _regraForm = { id, tipo: r.tipo, de: r.de, para: r.ignorar ? 'ignorar' : (r.para || []).map(x => x.sku + ' x' + x.qtd).join(', '), tam: r.tam || '' };
    render();
  };
  window.vendasDreCancelarRegra = function () { _regraForm = null; render(); };

  window.vendasDreGravarRegra = async function () {
    if (!exigirEdicao('configurar o estoque de produtos')) return;
    const tipo = val('dre-r-tipo') === 'descricao' ? 'descricao' : 'sku';
    const de = norm(val('dre-r-de'));
    const tam = norm(val('dre-r-tam'));
    if (!de) return toast('Informe o "De"', 'err');
    if (tam && !A().TAMS.includes(tam)) return toast(`Tamanho ${tam} não existe no estoque`, 'err');
    const p = lerPara(val('dre-r-para'));
    if (p.erro) return toast(p.erro, 'err');
    const c = dre();
    const id = _regraForm && _regraForm.id;
    if (c.mapa.some(r => r.id !== id && r.tipo === tipo && norm(r.de) === de)) return toast('Já existe uma regra com este "De"', 'err');
    const regra = { id: id || uid(), tipo, de, tam, ignorar: !!p.ignorar, para: p.para || [], em: new Date().toISOString(), por: quemSou() };
    c.mapa = id ? c.mapa.map(r => (r.id === id ? regra : r)) : c.mapa.concat([regra]);
    _regraForm = null;
    desfazerNomearAcao(`regra De → Para ${de}`);
    await saveState('meta');
    toast('Regra gravada', 'ok');
    // Reprocessa as pendências que ela resolve — por quem sincroniza (na
    // próxima rodada, se for outro computador; ver a trava).
    if (!travaDeOutro(dre()) && podeRodar()) await window.vendasDreRodar(true);
    else toast('O computador que sincroniza aplica a regra na próxima rodada (até 5 min)', 'ok');
    render();
  };

  window.vendasDreApagarRegra = async function (id) {
    if (!exigirEdicao('configurar o estoque de produtos')) return;
    const c = dre();
    const r = c.mapa.find(x => x.id === id);
    if (!r || !confirm(`Apagar a regra "${r.de}"?\n\nOs lançamentos já feitos por ela continuam.`)) return;
    c.mapa = c.mapa.filter(x => x.id !== id);
    desfazerNomearAcao(`regra De → Para ${r.de} apagada`);
    await saveState('meta');
    render();
  };

  // Chamado pelo acabApagarMov (estoque-confeccao.js) ao apagar um movimento
  // do DRE: o pedido fica dispensado, senão a rodada seguinte o relançaria.
  window.vendasDreDispensar = function (chave) {
    if (!chave) return;
    const c = dre();
    c.dispensados[chave] = new Date().toISOString();
    delete c.pendentes[chave];
  };

  // As lojas do ERP (Mercado Livre, Shopee, TikTok) ligadas aqui — para a
  // tela Saídas do ERP avisar.
  window.vendasDreLojasDoErp = function () {
    const c = (STATE.meta && STATE.meta.vendasDre) || {};
    const lojas = c.lojas || {};
    return Object.keys(lojas).filter(id => lojas[id].ligada && PLAT_COM_ERP.includes(lojas[id].plataforma))
      .map(id => ({ nome: lojas[id].nome || id, plataforma: lojas[id].plataforma, corte: lojas[id].corte }));
  };

  // Para os testes (node).
  window._vendasDreInterno = { converter, linhaDe, tamDaDescricao, acharRegra, processar, relevante, diaSP, lerPara, rodar, dre, ABA,
    _set: o => { if ('ultima' in o) _ultima = o.ultima; if ('pegouEm' in o) _pegouEm = o.pegouEm; } };
})();
