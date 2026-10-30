/* =====================================================================
   PLANEJAMENTO AUTOMÁTICO (30/10/2026)
   =====================================================================
   (Jones: "Faça com que o programa faça planejamento de OS automaticamente.
   Estoque atinge o mínimo e dispara relatório de demanda. Programa aguarda
   demanda aglomerar em grades sugeridas para produção e gera a OS
   automaticamente no planejamento de OS, que fica aguardando aprovação por
   usuário com concessão de aprovação".)

   O CAMINHO, com os dois elos que eram à mão agora sozinhos:

     Estoque de produtos ──(1)──▶ Relatório de demanda ──(2)──▶ Planejamento
     (Mín./Máx. por SKU e       (demanda automática,         (planejamento
      tamanho)                   recalculada a cada passada)   automático)
                                                                   │
                                        quem tem "Aprovar planejamento"
                                                                   ▼
                                                         OS com número

   (1) A DEMANDA AUTOMÁTICA. Decidido com o Jones: quando o saldo de um SKU e
       tamanho chega ao mínimo, pede-se o que falta para o MÁXIMO. Linha sem
       mínimo ou sem máximo não entra. A conta:
         falta de verdade = máximo − saldo − em produção − planejado − demanda à mão
       Só existe demanda automática enquanto essa falta for maior que zero. A
       quantidade gravada é a BRUTA (sem descontar o que está em produção),
       porque o agrupamento (_demGrupos) já abate o trabalho em progresso — se
       ela viesse líquida, a mesma OS seria descontada duas vezes.
       Ela é do programa, não de alguém: tem id fixo por SKU e tamanho
       ("auto:CM.LISA-PRE:g"), é refeita a cada passada (sobe quando vende,
       some quando a produção cobre) e não se edita à mão.

   (2) O PLANEJAMENTO AUTOMÁTICO. Decidido com o Jones: o agrupamento vira
       planejamento ASSIM QUE A DEMANDA FECHAR UM ENFESTO CHEIO (o limite de
       camadas do tecido). Usa a "solução mais eficiente" que o quadro do
       agrupamento já mostra, mas só com grades CADASTRADAS e com a medida
       provada por PDF de risco — grade nova e grade sem PDF ficam para o
       planejamento à mão, com o motivo escrito no quadro. Só entra o
       agrupamento que tem demanda automática; o feito só à mão segue à mão.
       O planejamento nasce com UMA aprovação pela frente (decidido com o
       Jones): entra direto na lista, marcado "🤖 Automático", e o "✔ Aprovar"
       de sempre gera a OS com o número do servidor.

   MONTAR SEM O FORMULÁRIO. O planejamento à mão sai do formulário da Nova OS;
   este não pode tomar a tela de quem está trabalhando. Ele é montado como a
   conjugada (gerarConjugada): a PEÇA vem do desenho e da grade (fases, cor por
   fase, tecidos, variante, enfesto) e o CONTEXTO — coleção, marca, base,
   bloco, equipe, etapas — vem da última OS do mesmo desenho (ou do mesmo SKU).
   Sem nenhuma OS anterior para copiar, não há como saber esse contexto: o
   agrupamento fica para o planejamento à mão, com o motivo no quadro.
   Componentes e aviamentos vêm da OS-modelo quando ela é do mesmo desenho (o
   que o formulário já acertou para aquela peça), recalculados para a grade
   nova; senão, do cadastro do desenho.

   QUEM RODA. O navegador de quem pode planejar (área Planejamento de produção
   ou Aprovar planejamento), num computador que pode gravar, a cada 5 minutos
   enquanto o programa estiver aberto — pendurado no mesmo relógio de 15 s do
   polling (iniciarPolling). Dois computadores rodando juntos não duplicam: os
   ids são fixos (a demanda por SKU e tamanho, o planejamento pelas demandas
   que ele cobre), e um planejamento só nasce se as demandas ainda estiverem
   lá. Não roda enquanto alguém monta à mão um planejamento da demanda
   (_demAgrupando), para não planejar duas vezes o mesmo agrupamento.

   Liga e desliga no Planejamento de produção (STATE.meta.autoPlan.ligado,
   ligado quando não há nada gravado).
   ===================================================================== */
(function () {
  'use strict';

  const INTERVALO_MS = 5 * 60 * 1000;
  const PRIMEIRA_MS = 20 * 1000;          // a primeira passada, logo depois de entrar
  const PREFIXO_DEM = 'auto:';
  const PREFIXO_PLAN = 'auto-plan-';
  const QUEM = '🤖 Automático';
  const TAM_DO_ESTOQUE = { P: 'p', M: 'm', G: 'g', GG: 'gg', G1: 'g1', G2: 'g2', G3: 'g3',
    2: 't2', 4: 't4', 6: 't6', 8: 't8', 10: 't10', 12: 't12', 14: 't14', 16: 't16' };
  const TAMS = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3', 't2', 't4', 't6', 't8', 't10', 't12', 't14', 't16'];
  const INFANTIL = ['t2', 't4', 't6', 't8', 't10', 't12', 't14', 't16'];

  let _ultima = 0;
  let _rodando = false;
  const _abertoEm = Date.now();
  // O que a última passada concluiu de cada agrupamento (chave do _demGrupos →
  // texto), para o quadro do agrupamento dizer por que ainda não planejou.
  let _motivos = new Map();
  let _ultimaEm = null;

  function cfg() {
    const c = (typeof STATE !== 'undefined' && STATE.meta && STATE.meta.autoPlan) || {};
    return { ligado: c.ligado !== false };
  }

  const ehAuto = d => !!(d && String(d.id || '').startsWith(PREFIXO_DEM));
  const soma = q => Object.values(q || {}).reduce((s, v) => s + (Number(v) || 0), 0);

  function podeRodar() {
    if (!cfg().ligado) return false;
    if (typeof currentUser === 'undefined' || !currentUser) return false;
    if (typeof podeGravar === 'function' && !podeGravar()) return false;
    if (!(temAcesso('planejamento') || temAcesso('aprovar-plan'))) return false;
    if (!window._acabInterno || !window._acabInterno.todosMovs) return false;
    if (typeof _demAgrupando !== 'undefined' && _demAgrupando) return false;
    return true;
  }

  /* ------------------------------------------------ (1) a demanda */

  // O SKU de prateleira do desenho (o mesmo do Estoque de produtos).
  const skuDe = desenhoId => {
    const d = _demDesenho(desenhoId);
    return d ? window._acabInterno.skuDoDesenho(d) : '';
  };

  // Peças por "SKU|tamanho" de uma lista de OS ou planejamentos, pela mesma
  // conta da folha (totaisPorTamanhoTomOS) que o abate do agrupamento usa.
  function porSkuTam(lista, vezes) {
    const m = new Map();
    (lista || []).forEach(o => {
      const sku = skuDe(o && o.desenhoId);
      if (!sku) return;
      let t;
      try { t = totaisPorTamanhoTomOS(o); } catch (e) { return; }
      const n = vezes ? vezes(o) : 1;
      (t.tamanhos || []).forEach(k => {
        const q = Math.round(t.colTotal(k)) || 0;
        if (q > 0) m.set(sku + '|' + k, (m.get(sku + '|' + k) || 0) + q * n);
      });
    });
    return m;
  }

  // O desenho que representa o SKU na demanda: o da OS mais recente com ele,
  // senão o primeiro desenho com esse SKU.
  function desenhoDoSku(sku) {
    const candidatos = (STATE.desenhos || []).filter(d => window._acabInterno.skuDoDesenho(d) === sku);
    if (!candidatos.length) return null;
    const ids = new Set(candidatos.map(d => d.id));
    const ultima = (STATE.ordens || []).filter(o => ids.has(o.desenhoId))
      .sort((a, b) => String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')))[0];
    return ultima ? candidatos.find(d => d.id === ultima.desenhoId) : candidatos[0];
  }

  // As linhas do Estoque de produtos que têm mínimo E máximo, com tudo o que a
  // conta da demanda precisa.
  function linhasDoEstoque() {
    const ai = window._acabInterno;
    const { de } = ai.saldos(ai.todosMovs());
    const wip = porSkuTam((STATE.ordens || []).filter(o => !_DEM_FORA_DO_WIP.has(_statusOS(o))));
    const plan = porSkuTam(_planLista(), _planVezes);
    const manual = new Map();
    _demLista().filter(d => !ehAuto(d)).forEach(d => {
      const sku = skuDe(d.desenhoId);
      if (!sku) return;
      const q = _demQtds(d);
      Object.keys(q).forEach(k => manual.set(sku + '|' + k, (manual.get(sku + '|' + k) || 0) + (q[k] || 0)));
    });
    const linhas = [];
    ai.produtos().forEach(p => {
      const sku = String(p.sku || '');
      if (!sku.includes('-') || !p.limites) return;
      Object.keys(p.limites).forEach(tamEst => {
        const tam = TAM_DO_ESTOQUE[tamEst];
        if (!tam) return;                      // PP: a demanda não tem esse tamanho
        const lim = ai.limiteDe(p, tamEst);
        if (lim.min == null || lim.max == null) return;
        const des = desenhoDoSku(sku);
        const k = sku + '|' + tam;
        linhas.push({ sku, tam, tamEst, desenhoId: des ? des.id : '', saldo: de(p.id, tamEst).total,
          min: lim.min, max: lim.max, wip: wip.get(k) || 0, planejado: plan.get(k) || 0, manual: manual.get(k) || 0 });
      });
    });
    return linhas;
  }

  /* A conta, sem tela nem STATE (testada em testes/planejamento-automatico.js).
     `atuais`: Map id → demanda automática gravada. Devolve o que gravar e o
     que apagar. */
  function calcularDemandas(linhas, atuais, agora) {
    const gravar = [], manter = new Set();
    (linhas || []).forEach(l => {
      if (l.min == null || l.max == null || !l.desenhoId) return;
      const id = PREFIXO_DEM + l.sku + ':' + l.tam;
      const ja = atuais.get(id);
      // Dispara no mínimo; depois de disparada, segue até a falta acabar —
      // senão uma venda a menos a faria sumir e voltar a cada passada.
      if (!ja && !(l.saldo <= l.min)) return;
      if (!(l.max - l.saldo - l.wip - l.planejado - l.manual > 0)) return;
      const qtd = Math.round(l.max - l.saldo - l.planejado - l.manual);
      manter.add(id);
      if (ja && ja.qtd === qtd && ja.desenhoId === l.desenhoId) return;
      gravar.push(Object.assign({ id, desenhoId: l.desenhoId, tam: l.tam, qtd,
        auto: { sku: l.sku, saldo: l.saldo, min: l.min, max: l.max, emProducao: l.wip,
          planejado: l.planejado, calculadoEm: agora },
        criadoEm: ja ? ja.criadoEm : agora, criadoPor: QUEM },
        ja ? { editadoEm: agora, editadoPor: QUEM } : {}));
    });
    const apagar = Array.from(atuais.keys()).filter(id => !manter.has(id));
    return { gravar, apagar };
  }

  async function passoDemanda(agora) {
    const atuais = new Map(_demLista().filter(ehAuto).map(d => [d.id, d]));
    const { gravar, apagar } = calcularDemandas(linhasDoEstoque(), atuais, agora);
    if (!gravar.length && !apagar.length) return false;
    const fora = new Set(apagar.concat(gravar.map(d => d.id)));
    STATE.osDemandas = _demLista().filter(d => !fora.has(d.id)).concat(gravar);
    await saveState('osDemandas');
    return true;
  }

  /* ------------------------------------------- (2) o planejamento */

  // Fechou um enfesto cheio? A demanda a produzir enche pelo menos um enfesto
  // (o limite de camadas do tecido) da menor parte da solução. Solução com
  // grade nova, ou com tecido sem limite de camadas conhecido, não conta.
  function aglomerou(cb, q) {
    if (!cb || cb.novas || !(cb.partes || []).length) return false;
    if (cb.partes.some(p => !((p.cam || 1) > 1))) return false;
    const umEnfesto = Math.min(...cb.partes.map(p => soma(_demRende(p))));
    return soma(q) >= umEnfesto;
  }

  // A solução com grades cadastradas e medida provada por PDF, ou o motivo de
  // não haver uma.
  function escolherSolucao(gr) {
    const todas = _demCombinacoes(gr);
    const cadastradas = todas.filter(c => !c.novas);
    if (!cadastradas.length) {
      return { motivo: todas.length ? 'só fecha com grade nova — planeje à mão e cadastre a grade sugerida'
        : 'nenhuma grade deste SKU fecha a demanda' };
    }
    const comProva = cadastradas.find(c => c.partes.every(p => !fasesSemProvaDeMedida(p.gradeId).length));
    if (!comProva) return { motivo: 'a grade da solução tem fase sem PDF de risco — importe o risco ou planeje à mão' };
    return { cb: comProva };
  }

  // A OS que empresta o contexto: a última do mesmo desenho; senão a última
  // do mesmo SKU; senão um planejamento do mesmo desenho.
  function osModelo(desenhoId) {
    const recente = l => l.slice().sort((a, b) => String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')))[0] || null;
    const ordens = (STATE.ordens || []).filter(o => _statusOS(o) !== 'cancelado');
    const doDesenho = recente(ordens.filter(o => o.desenhoId === desenhoId));
    if (doDesenho) return doDesenho;
    const sku = skuDe(desenhoId);
    const doSku = sku ? recente(ordens.filter(o => skuDe(o.desenhoId) === sku)) : null;
    return doSku || recente(_planLista().filter(p => p.desenhoId === desenhoId && !p.auto));
  }

  const nomeDe = (lista, id, campo) => ((lista || []).find(x => x.id === id) || {})[campo || 'nome'] || '';

  // Quantidades por tamanho de um item (componente ou aviamento) para as peças
  // da grade, como o coletaOS faz.
  function porTamanho(pecas, qtdPorPeca) {
    const qtdPorTamanho = {};
    let qtdTotal = 0;
    TAMS.forEach(t => { const v = (pecas[t] || 0) * qtdPorPeca; qtdPorTamanho[t] = v; qtdTotal += v; });
    return { qtdPorTamanho, qtdTotal };
  }

  /* O planejamento de UMA parte da solução: a grade `g` com `cam` camadas
     (o enfesto cheio), repetida `vezes` vezes. Ver MONTAR SEM O FORMULÁRIO. */
  function montarPlano(modelo, d, g, cam, vezes, agora) {
    const tam = g.tamanhos || {};
    const grade = { descricao: g.nome || '' };
    ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].forEach(k => { grade[k] = parseInt(tam[k], 10) || 0; });
    INFANTIL.forEach(k => { const q = parseInt(tam[k], 10) || 0; if (q > 0) grade[k] = q; });
    grade.total = TAMS.reduce((s, k) => s + (grade[k] || 0), 0);

    // A cor de cada fase é a do desenho, na ordem (a mesma regra do formulário
    // e da conjugada): fase 1 → primária, 2 → secundária, 3 → terciária, e da
    // 4ª em diante repete a primária.
    const coresIds = ordenarCoresIdsPorDesc(
      [d.corPrincipalId, d.corSecundariaId, d.corTerciariaId].filter(Boolean), d);
    const nomeCor = id => nomeDe(STATE.cores, id);
    const corDaFase = n => coresIds[n - 1] || coresIds[0] || '';
    const fases = (g.fases || []).map((f, i) => {
      const corId = corDaFase(i + 1) || f.corId || '';
      return { ordem: f.ordem, nome: f.nome || '', tecidoId: f.tecidoId || '',
        tecidoNome: nomeDe(STATE.tecidos, f.tecidoId), corId, corNome: nomeCor(corId),
        comp: f.comp || '', larg: f.larg || '' };
    });
    const tecidos = fases.map(f => ({ tecidoId: f.tecidoId, tecidoNome: f.tecidoNome,
      corId: f.corId, corNome: f.corNome, c1: '', lote: '', loteTexto: '' }));
    const variantes = [{ num: 1, obs: '',
      cor1: coresIds[0] || '', cor1Nome: nomeCor(coresIds[0]) || '—',
      cor2: coresIds[1] || '', cor2Nome: nomeCor(coresIds[1]) || '—',
      cor3: coresIds[2] || '', cor3Nome: nomeCor(coresIds[2]) || '—' }];

    const mult = unidadesPorCamadaPrincipal(fases.map(f => (STATE.tecidos || []).find(t => t.id === f.tecidoId)));
    const pecas = {};
    TAMS.forEach(k => { pecas[k] = (grade[k] || 0) * cam * mult; });

    const mesmoDesenho = modelo && modelo.desenhoId === d.id;
    const componentes = mesmoDesenho
      ? (modelo.componentes || []).map(c => Object.assign({ nome: c.nome, material: c.material, materialNome: c.materialNome,
          cor: c.cor || '', corNome: c.corNome || '', qtdPorPeca: Number(c.qtdPorPeca) || 0 },
          porTamanho(pecas, Number(c.qtdPorPeca) || 0)))
      : (d.componentes || []).map(c => {
          const qtdPorPeca = c.qtdPorPeca != null ? Number(c.qtdPorPeca) || 0 : 1;
          return Object.assign({ nome: c.nome || nomeDe(STATE.componentes, c.componenteId),
            material: c.tecidoId ? 'T:' + c.tecidoId : '', materialNome: nomeDe(STATE.tecidos, c.tecidoId),
            cor: c.corId || '', corNome: nomeCor(c.corId), qtdPorPeca }, porTamanho(pecas, qtdPorPeca));
        });
    const aviamentos = mesmoDesenho
      ? (modelo.aviamentos || []).map(a => {
          const qtdPorPeca = Number(a.qtdPorPeca != null ? a.qtdPorPeca : a.qtd) || 0;
          return Object.assign({ material: a.material, materialNome: a.materialNome || '', app: a.app || '',
            qtd: qtdPorPeca, qtdPorPeca }, porTamanho(pecas, qtdPorPeca));
        })
      : (d.aviamentos || []).map(a => {
          const qtdPorPeca = parseFloat(a.qtdPorPeca) || 1;
          return Object.assign({ material: a.materialId, materialNome: nomeDe(STATE.materiais, a.materialId, 'desc'),
            app: a.aplicacao || '', qtd: qtdPorPeca, qtdPorPeca }, porTamanho(pecas, qtdPorPeca));
        });

    const blocos = fases.length
      ? fases.map((f, i) => ({ ordem: i + 1, nomeTecido: f.nome || f.tecidoNome, nomeCor: f.corNome,
          comp: parseFloat(f.comp) || 0, larg: parseFloat(f.larg) || 0,
          camadas: /vi[eé]s/i.test(f.nome || '') ? 1 : cam, bobinas: '' }))
      : [{ ordem: 1, nomeTecido: '', nomeCor: '', comp: 0, larg: 0, camadas: cam, bobinas: '' }];
    const enfesto = { comprimento: blocos[0].comp, largura: blocos[0].larg, camadas: cam, target: 0,
      blocos, totalPecas: grade.total * cam };

    const m = modelo || {};
    const plano = {
      id: '', os: '', codigo: d.codigo || '', data: agora.slice(0, 10),
      coordenadoId: m.coordenadoId || '', coordenadoNome: m.coordenadoNome || '',
      colecaoId: m.colecaoId || '', colecaoNome: m.colecaoNome || '',
      modeloId: d.modeloId || m.modeloId || '',
      modeloNome: nomeDe(STATE.modelos, d.modeloId) || m.modeloNome || '',
      blocoId: m.blocoId || '', blocoNome: m.blocoNome || '',
      linhaId: m.linhaId || '', linhaNome: m.linhaNome || '',
      griffeId: m.griffeId || '', griffeNome: m.griffeNome || '',
      baseId: m.baseId || '', baseNome: m.baseNome || '',
      designerId: m.designerId || '', designerNome: m.designerNome || '',
      ftecId: m.ftecId || '', ftecNome: m.ftecNome || '',
      desenhoId: d.id, gradeId: g.id, fases, tecidos, grade, enfesto,
      etapas: (m.etapas || []).slice(), variantes, componentes, aviamentos,
      // Reposição de estoque é produção puxada: quem pediu foi a prateleira.
      tipoProducao: 'puxada', repeticoes: Math.max(1, vezes || 1),
      obs: '', atencao: '', criadoEm: agora, criadoPor: QUEM
    };
    return _aplicarTonsPadrao(plano);
  }

  // Id fixo: o mesmo conjunto de demandas e a mesma parte dão sempre o mesmo
  // id — dois computadores na mesma passada geram o MESMO planejamento.
  function idDoPlano(ids, parte) {
    const txt = ids.slice().sort().join('+') + '#' + parte;
    let h = 5381;
    for (let i = 0; i < txt.length; i++) h = ((h * 33) ^ txt.charCodeAt(i)) >>> 0;
    return PREFIXO_PLAN + h.toString(36) + '-' + parte;
  }

  async function passoPlanejamento(agora) {
    const motivos = new Map();
    const novos = [];
    const consumidas = new Set();
    _demGrupos().forEach(gr => {
      if (!gr.ids.some(id => String(id).startsWith(PREFIXO_DEM))) return;   // só à mão: segue à mão
      if (gr.coberta) return;
      const d = _demDesenho(gr.desenhoId);
      if (!d) return;
      const { cb, motivo } = escolherSolucao(gr);
      if (!cb) { motivos.set(gr.chave, motivo); return; }
      if (!aglomerou(cb, gr.q)) {
        const um = Math.min(...cb.partes.map(p => soma(_demRende(p))));
        motivos.set(gr.chave, cb.partes.some(p => !((p.cam || 1) > 1))
          ? 'o tecido da grade não tem limite de camadas cadastrado — sem ele não há "enfesto cheio"; planeje à mão'
          : `acumulando: ${soma(gr.q).toLocaleString('pt-BR')} de ${um.toLocaleString('pt-BR')} peças para fechar 1 enfesto cheio`);
        return;
      }
      const modelo = osModelo(d.id);
      if (!modelo) {
        motivos.set(gr.chave, 'nenhuma OS anterior deste desenho para copiar coleção, marca e etapas — planeje à mão');
        return;
      }
      const combinacao = _demCombTexto(cb);
      cb.partes.forEach((p, i) => {
        const g = (STATE.grades || []).find(x => x.id === p.gradeId);
        if (!g) return;
        const plano = montarPlano(modelo, d, g, p.cam, p.c, agora);
        plano.id = idDoPlano(gr.ids, i + 1);
        plano.auto = { em: agora, regra: '1 enfesto cheio', modeloOS: modelo.os || '', rodouEm: _obsQuemSou() };
        plano.demanda = { ids: gr.ids.slice(), pedidas: Object.assign({}, gr.q), sugerida: null,
          lancadas: Object.assign({}, gr.pedido), abatidas: Object.assign({}, gr.wip),
          combinacao, parte: i + 1, partes: cb.partes.length, pendentes: [],
          planejadoEm: agora, planejadoPor: QUEM };
        novos.push(plano);
      });
      gr.ids.forEach(id => consumidas.add(id));
    });
    _motivos = motivos;
    if (!novos.length) return false;
    // Outro computador pode ter planejado agora há pouco: o mesmo id já na
    // lista não entra de novo.
    const tem = new Set(_planLista().map(p => p.id));
    const entram = novos.filter(p => !tem.has(p.id));
    STATE.osPlanejadas = _planLista().concat(entram);
    STATE.osDemandas = _demLista().filter(d => !consumidas.has(d.id));
    await saveState('osPlanejadas');
    await saveState('osDemandas');
    if (entram.length) {
      toast(`🤖 Planejamento automático: ${entram.length} ${entram.length === 1 ? 'planejamento novo aguarda' : 'planejamentos novos aguardam'} aprovação`, 'ok');
    }
    return true;
  }

  /* ------------------------------------------------- a passada */

  async function rodar(forcar) {
    if (_rodando || !podeRodar()) return;
    const agoraMs = Date.now();
    if (!forcar && (agoraMs - _abertoEm < PRIMEIRA_MS || agoraMs - _ultima < INTERVALO_MS)) return;
    _rodando = true;
    _ultima = agoraMs;
    try {
      const agora = new Date().toISOString();
      const mudouDem = await passoDemanda(agora);
      const mudouPlan = await passoPlanejamento(agora);
      _ultimaEm = agora;
      if (mudouDem || mudouPlan) atualizarContagensNav();
      const ativa = document.querySelector('section.page:not(.hidden)');
      const pagina = ativa && ativa.dataset && ativa.dataset.page;
      if (pagina === 'planejamento-producao') renderPlanejamentoProducao();
      else if (pagina === 'relatorio-demanda' && mudouDem) renderRelatorioDemanda();
    } catch (e) {
      console.warn('planejamento automático', e);
    } finally {
      _rodando = false;
    }
  }

  window.autoPlanRodar = () => rodar(false);

  window.autoPlanRodarAgora = async function () {
    if (!cfg().ligado) return toast('O planejamento automático está desligado', 'err');
    if (!podeRodar()) return toast('Este computador ou acesso não roda o planejamento automático', 'err');
    await rodar(true);
    toast('Planejamento automático conferido', 'ok');
  };

  window.autoPlanLigar = async function (ligado) {
    if (typeof podeAprovarPlanejamento === 'function' && !podeAprovarPlanejamento()) {
      return toast('Só quem pode aprovar o planejamento liga ou desliga o automático', 'err');
    }
    if (_recusarSomenteLeitura('ligar o planejamento automático')) return;
    STATE.meta = STATE.meta || {};
    STATE.meta.autoPlan = Object.assign({}, STATE.meta.autoPlan || {}, { ligado: !!ligado,
      alteradoEm: new Date().toISOString(), alteradoPor: _obsQuemSou() });
    await saveState('meta');
    toast(ligado ? 'Planejamento automático ligado' : 'Planejamento automático desligado', 'ok');
    renderPlanejamentoProducao();
    if (ligado) rodar(true);
  };

  // A faixa do topo do Planejamento de produção.
  window.autoPlanFaixaHtml = function () {
    const c = cfg();
    const hora = t => { const d = new Date(t); return isNaN(d) ? '' : String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };
    const nAuto = _planLista().filter(p => p.auto).length;
    const nDem = _demLista().filter(ehAuto).length;
    const pode = typeof podeAprovarPlanejamento === 'function' && podeAprovarPlanejamento();
    return `<div style="display:flex;flex-wrap:wrap;gap:8px 16px;align-items:center;font-size:12px;">
      <b>🤖 Planejamento automático: ${c.ligado ? '<span style="color:#2e9d5b;">ligado</span>' : '<span style="color:var(--alert);">desligado</span>'}</b>
      <span style="color:var(--ink-2);">Estoque no mínimo → demanda até o máximo → planejamento quando fecha 1 enfesto cheio → aguarda aprovação.</span>
      <span style="color:var(--ink-2);">${nDem} ${nDem === 1 ? 'demanda automática' : 'demandas automáticas'} · ${nAuto} ${nAuto === 1 ? 'planejamento automático' : 'planejamentos automáticos'} na lista${_ultimaEm ? ` · conferido às ${hora(_ultimaEm)}` : ''}</span>
      <span style="margin-left:auto;display:flex;gap:6px;">
        ${c.ligado ? `<button type="button" class="btn small" onclick="autoPlanRodarAgora()" title="Confere o estoque e a demanda agora, sem esperar os 5 minutos">Conferir agora</button>` : ''}
        ${pode ? `<button type="button" class="btn small ${c.ligado ? '' : 'primary'}" onclick="autoPlanLigar(${!c.ligado})">${c.ligado ? 'Desligar' : 'Ligar'}</button>` : ''}
      </span>
    </div>`;
  };

  // O motivo de o agrupamento ainda não ter virado planejamento automático.
  window.autoPlanMotivo = chave => _motivos.get(chave) || '';
  window.autoPlanEhDemandaAuto = ehAuto;

  // Para os testes (node).
  window._autoPlanInterno = { calcularDemandas, aglomerou, montarPlano, idDoPlano, escolherSolucao, osModelo, PREFIXO_DEM, QUEM };
})();
