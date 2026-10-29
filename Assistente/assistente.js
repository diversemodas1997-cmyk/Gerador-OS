/* =====================================================================
   ASSISTENTE CLAUDE — janela de conversa dentro do Gerador-OS (19/10/2026)
   =====================================================================
   (Junior: chat do Claude dentro do programa, para TODOS os usuários, SÓ
   CONSULTA.)

   Arquivo separado do app.js de propósito: o app.js tem 2,4 MB e mexer nele
   para pendurar uma janela seria arriscar o resto. Este arquivo só LÊ o que o
   app.js já deixou carregado (STATE, currentUser, currentRole, supa,
   _servidorUrlAtivo, _statusOS, produtosOS, STATUS_OS, contaVeArea) e nunca
   grava nada: nenhuma ferramenta daqui chama cloudFlush, scheduleCloudSave ou
   altera STATE.

   O CAMINHO DE UMA PERGUNTA
   1. A pergunta vai para /functions/v1/assistente (servidor da fábrica), que
      guarda a chave da Anthropic e fala com o Claude.
   2. Se o Claude pede uma ferramenta (buscar OS, resumo por status...), a
      resposta volta para cá, a ferramenta roda AQUI, com os dados desta tela e
      com a visão desta conta (contaVeArea), e o resultado volta ao Claude.
   3. Repete até o Claude responder em texto (no máximo RODADAS_MAX vezes).
   ===================================================================== */
(function () {
  'use strict';

  const RODADAS_MAX = 8;
  const TETO_RESULTADO = 14000;      // caracteres por resultado de ferramenta
  const HIST_MAX = 50;               // mensagens guardadas na conversa

  /* ---------------- quem vê o quê ----------------
     Cada base do STATE pertence a uma área de VISÃO (ver VISAO_PAGINAS no
     app.js). Base que não está aqui é só do admin. */
  const BASES_AREA = {
    ordens: 'os', osPlanejadas: 'planejamento', osDemandas: 'demanda',
    cores: 'cadastros', tecidos: 'cadastros', modelos: 'cadastros', grades: 'cadastros',
    desenhos: 'cadastros', etapas: 'cadastros', equipe: 'cadastros', bases: 'cadastros',
    blocos: 'cadastros', linhas: 'cadastros', marcas: 'cadastros', colecoes: 'cadastros',
    componentes: 'cadastros', materiais: 'cadastros', fornecedores: 'cadastros',
    funcoes: 'cadastros', tarefas: 'cadastros', gradeFolderLabels: 'cadastros',
    aviamentoTipos: 'cadastros',
    estoqueMov: 'estoque-tecidos', materiaisEstCad: 'estoque-tecidos', materiaisEstMov: 'estoque-tecidos',
    aviamentosMov: 'estoque-tecidos', pecasCad: 'estoque-tecidos', pecasMov: 'estoque-tecidos',
    corteMov: 'estoque-fases', fiosMov: 'estoque-fases', costurandoMov: 'estoque-fases',
    expedicaoMov: 'estoque-fases',
    expedicaoCargas: 'expedicao', expedicaoJanelas: 'expedicao', expedicaoExcecoes: 'expedicao', expedicaoObs: 'expedicao',
    operacoes: 'operacoes', costuraProd: 'costura', compraOCs: 'compra', compraPlano: 'compra'
  };
  const BASES_NUNCA = new Set(['_device']);

  /* A CONTA DO ASSISTENTE (19/10/2026, Junior: "configure uma conta de acesso
     para que o Assistente, por meio dessa conta, possa gerar folhas de OS, OE,
     ordens de compra...").

     As FOLHAS (OS, etiquetas, OE, OC) precisam de DUAS licenças: a visão da
     conta "assistente" (que o admin cria e ajusta em Configurações → Contas de
     acesso, como qualquer outra; sem marcação ela vê tudo) E a visão de quem
     está logado. A conta do assistente funciona como um limite a mais do que o
     chat pode fazer — nunca como atalho para alguém abrir pelo chat uma folha
     que a própria conta não abre. Nada aqui grava dados nem usa senha. */
  const CONTA_ASSISTENTE = 'assistente';
  function assistenteVe(area) {
    try { return typeof contaVeArea === 'function' ? contaVeArea(area, CONTA_ASSISTENTE, 'usuario') : false; }
    catch (e) { return false; }
  }

  const ehAdmin = () => typeof currentRole !== 'undefined' && currentRole === 'admin';
  const login = () => (typeof _obsQuemSou === 'function' ? _obsQuemSou() : '');

  function veArea(area) {
    if (ehAdmin()) return true;
    if (!area) return false;
    try { return typeof contaVeArea === 'function' ? contaVeArea(area, login()) : true; }
    catch (e) { return false; }
  }
  function podeBase(nome) {
    if (BASES_NUNCA.has(nome)) return false;
    if (ehAdmin()) return true;
    return veArea(BASES_AREA[nome]);
  }
  const SEM_ACESSO = area => ({ sem_acesso: true, mensagem: `Esta conta não tem visão da área "${area}". Peça ao admin.` });

  /* ---------------- utilidades ---------------- */
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

  // Cópia enxuta para mandar ao Claude: corta textos longos (imagens em base64,
  // PDFs), listas enormes e profundidade, que só custam e não informam.
  function enxugar(v, prof = 0, maxTxt = 400, maxLista = 40) {
    if (v == null || typeof v === 'number' || typeof v === 'boolean') return v;
    if (typeof v === 'string') {
      if (/^data:/.test(v)) return '[arquivo]';
      return v.length > maxTxt ? v.slice(0, maxTxt) + '…' : v;
    }
    if (prof >= 5) return Array.isArray(v) ? `[${v.length} itens]` : '[…]';
    if (Array.isArray(v)) {
      const r = v.slice(0, maxLista).map(x => enxugar(x, prof + 1, maxTxt, maxLista));
      if (v.length > maxLista) r.push(`… mais ${v.length - maxLista} itens`);
      return r;
    }
    if (typeof v === 'object') {
      const r = {};
      for (const k of Object.keys(v)) {
        if (/Id$|^id$/.test(k) && typeof v[k] === 'string' && /^id_/.test(v[k])) continue; // ids internos
        r[k] = enxugar(v[k], prof + 1, maxTxt, maxLista);
      }
      return r;
    }
    return String(v);
  }

  function limitar(obj) {
    let s = JSON.stringify(obj);
    if (s.length > TETO_RESULTADO) s = s.slice(0, TETO_RESULTADO) + ' …[resposta cortada: refine a busca ou use limite/pular]';
    return s;
  }

  function rotuloStatus(k) {
    const s = (typeof STATUS_OS !== 'undefined' ? STATUS_OS : []).find(x => x.k === k);
    return s ? s.rotulo : k;
  }
  function statusDe(o) { try { return _statusOS(o); } catch (e) { return ''; } }
  function produtosDe(o) { try { return produtosOS(o); } catch (e) { return null; } }

  function coresDaOS(o) {
    const nomes = [];
    (o.variantes || []).forEach(v => {
      if (!v || typeof v !== 'object') return;
      Object.keys(v).forEach(k => { if (/Nome$/.test(k) && v[k]) nomes.push(v[k]); });
    });
    return [...new Set(nomes)];
  }
  function tecidosDaOS(o) {
    return [...new Set((o.tecidos || []).map(t => [t && t.tecidoNome, t && t.corNome].filter(Boolean).join(' ')).filter(Boolean))];
  }
  function resumoOS(o) {
    const st = statusDe(o);
    let desde = null;
    try { const t = _statusDesdeOS(o); if (t) desde = new Date(t).toLocaleString('pt-BR'); } catch (e) { /* sem diário */ }
    return {
      os: o.os, codigo: o.codigo, data: o.data, modelo: o.modeloNome, colecao: o.colecaoNome,
      linha: o.linhaNome, base: o.baseNome, bloco: o.blocoNome,
      grade: o.grade && o.grade.descricao, cores: coresDaOS(o), tecidos: tecidosDaOS(o),
      status: rotuloStatus(st), status_desde: desde, produtos: produtosDe(o),
      obs: o.obs || undefined, atencao: o.atencao || undefined, criadoPor: o.criadoPor || undefined
    };
  }
  function textoBuscaOS(o) {
    return norm([o.os, o.codigo, o.modeloNome, o.colecaoNome, o.linhaNome, o.baseNome, o.blocoNome,
      o.grade && o.grade.descricao, coresDaOS(o).join(' '), tecidosDaOS(o).join(' '),
      o.obs, o.atencao, o.criadoPor, o.designerNome, o.coordenadoNome].join(' | '));
  }
  const numOS = s => String(s || '').replace(/\D/g, '').replace(/^0+/, '');

  /* ---------------- ferramentas (só leitura) ---------------- */
  const FERRAMENTAS = {
    resumo_status() {
      const linhas = (typeof _dashPorStatus === 'function') ? _dashPorStatus() : [];
      return {
        status: linhas.map(x => ({ chave: x.st.k, status: x.st.rotulo, os: x.os, produtos: x.produtos,
          fora_da_producao: !!x.st.foraDaProducao || x.st.k === 'estoque' })),
        total_os_cadastradas: (STATE.ordens || []).length
      };
    },

    buscar_os(p) {
      if (!podeBase('ordens')) return SEM_ACESSO('Ordens de Serviço');
      const t = norm(p.texto || '').trim();
      const soNum = /^\d+$/.test(t);
      const stq = norm(p.status || '').trim();
      const lim = Math.max(1, Math.min(80, parseInt(p.limite, 10) || 30));
      let lista = (STATE.ordens || []).filter(o => {
        if (t) {
          if (soNum && numOS(o.os) === numOS(t)) return true;
          if (!textoBuscaOS(o).includes(t)) return false;
        }
        if (stq) {
          const k = statusDe(o);
          if (norm(k) !== stq && !norm(rotuloStatus(k)).includes(stq)) return false;
        }
        if (p.data_de && String(o.data || '') < p.data_de) return false;
        if (p.data_ate && String(o.data || '') > p.data_ate) return false;
        return true;
      });
      lista.sort((a, b) => (parseInt(numOS(b.os), 10) || 0) - (parseInt(numOS(a.os), 10) || 0));
      const total = lista.length;
      const prod = lista.reduce((s, o) => s + (produtosDe(o) || 0), 0);
      return { encontradas: total, produtos_somados: prod, mostrando: Math.min(total, lim),
        os: lista.slice(0, lim).map(resumoOS) };
    },

    detalhe_os(p) {
      if (!podeBase('ordens')) return SEM_ACESSO('Ordens de Serviço');
      const n = numOS(p.os);
      const o = (STATE.ordens || []).find(x => numOS(x.os) === n);
      if (!o) return { erro: `Não achei a OS ${p.os}.` };
      return { resumo: resumoOS(o), completo: enxugar(o) };
    },

    listar_bases() {
      const out = [];
      Object.keys(STATE || {}).forEach(nome => {
        if (!podeBase(nome)) return;
        const v = STATE[nome];
        if (Array.isArray(v)) {
          const campos = new Set();
          v.slice(0, 40).forEach(x => x && typeof x === 'object' && Object.keys(x).forEach(k => campos.add(k)));
          out.push({ base: nome, registros: v.length, campos: [...campos].slice(0, 40) });
        } else if (v && typeof v === 'object') {
          out.push({ base: nome, registros: Object.keys(v).length, tipo: 'objeto' });
        }
      });
      return { bases: out };
    },

    consultar_base(p) {
      const nome = String(p.base || '');
      if (!(nome in (STATE || {}))) return { erro: `Base "${nome}" não existe. Use listar_bases.` };
      if (!podeBase(nome)) return SEM_ACESSO(BASES_AREA[nome] || 'administração');
      const v = STATE[nome];
      let regs = Array.isArray(v) ? v.slice()
        : (v && typeof v === 'object') ? Object.entries(v).map(([k, x]) => (x && typeof x === 'object' && !Array.isArray(x)) ? Object.assign({ _chave: k }, x) : { _chave: k, valor: x })
        : [{ valor: v }];
      if (p.mais_recentes_primeiro) regs.reverse();
      const t = norm(p.texto || '').trim();
      if (t) regs = regs.filter(r => norm(JSON.stringify(r)).includes(t));
      const total = regs.length;
      const pular = Math.max(0, parseInt(p.pular, 10) || 0);
      const lim = Math.max(1, Math.min(100, parseInt(p.limite, 10) || 40));
      const campos = Array.isArray(p.campos) && p.campos.length ? p.campos.map(String) : null;
      const fatia = regs.slice(pular, pular + lim).map(r => {
        if (!campos || !r || typeof r !== 'object') return enxugar(r, 1, 200, 15);
        const o = {};
        campos.forEach(c => { if (c in r) o[c] = enxugar(r[c], 2, 200, 15); });
        return o;
      });
      return { base: nome, encontrados: total, pulados: pular, mostrando: fatia.length, registros: fatia };
    },

    preparar_folha(p) {
      const tipo = String(p.tipo || '').toLowerCase();
      const imprimir = !!p.imprimir;
      const AREA = { os: 'os', etiquetas: 'os', oe: 'expedicao', oc: 'compra' };
      if (!AREA[tipo]) return { erro: 'tipo deve ser os, etiquetas, oe ou oc' };
      if (!assistenteVe(AREA[tipo])) {
        return { sem_acesso: true, mensagem: `A conta do assistente ("${CONTA_ASSISTENTE}") não tem visão da área "${AREA[tipo]}". O admin libera em Configurações → Contas de acesso.` };
      }
      if (!veArea(AREA[tipo])) return SEM_ACESSO(AREA[tipo]);
      const verbo = imprimir ? '🖨 Imprimir' : '📄 Abrir';

      if (tipo === 'os' || tipo === 'etiquetas') {
        const n = numOS(p.alvo);
        const o = (STATE.ordens || []).find(x => numOS(x.os) === n);
        if (!o) return { erro: `Não achei a OS ${p.alvo}.` };
        const nome = tipo === 'os' ? `folha da OS ${o.os}` : `etiquetas da OS ${o.os}`;
        oferecerAcao(`${verbo} ${nome}`, () => {
          if (tipo === 'etiquetas') { imprimirEtiquetasPdf(o.id); return; }
          fechar();
          verOS(o.id);
          if (imprimir) setTimeout(() => window.print(), 700);
        });
        return { botao_mostrado: `${verbo} ${nome}`, os: resumoOS(o),
          aviso: tipo === 'etiquetas' ? 'As etiquetas abrem como PDF numa aba nova; a impressão é feita por lá.' : undefined };
      }

      if (tipo === 'oe') {
        const data = dataIso(p.alvo);
        if (!data) return { erro: `Não entendi a data "${p.alvo}". Use AAAA-MM-DD, DD/MM, "hoje" ou "amanhã".` };
        const rot = data.split('-').reverse().join('/');
        oferecerAcao(`${verbo} folha de OE de ${rot}`, () => {
          fechar();
          expPlanoModo = 'dia';
          expPlanoAncora = data;
          try { sessionStorage.setItem('gos:exp:modo', 'dia'); sessionStorage.setItem('gos:exp:ancora', data); } catch (e) { /* sem storage */ }
          goto('print-expedicao');
          if (imprimir) setTimeout(() => window.print(), 700);
        });
        return { botao_mostrado: `${verbo} folha de OE de ${rot}`, data,
          aviso: 'A folha só traz as expedições com carga alocada; dia sem carga sai vazio.' };
      }

      // oc
      const alvo = String(p.alvo || '').replace(/\D/g, '').replace(/^0+/, '');
      const ocs = Array.isArray(STATE.compraOCs) ? STATE.compraOCs : [];
      const oc = ocs.find(x => String(x.numero || '').replace(/\D/g, '').replace(/^0+/, '') === alvo);
      if (!oc) return { erro: `Não achei a ordem de compra ${p.alvo}.`, existentes: ocs.slice(-15).map(x => x.numero) };
      // A janela da OC já abre o diálogo de impressão sozinha.
      oferecerAcao(`🖨 Abrir e imprimir ${oc.numero}`, () => imprimirOC(oc.id));
      return { botao_mostrado: `Abrir e imprimir ${oc.numero}`, oc: enxugar(oc, 1, 200, 20) };
    },

    tela_atual() {
      const sec = [...document.querySelectorAll('section.page')].find(s => !s.classList.contains('hidden'));
      if (!sec) return { erro: 'Não identifiquei a tela aberta.' };
      const texto = (sec.innerText || '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
      return { pagina: sec.getAttribute('data-page'), texto: texto.slice(0, 9000) + (texto.length > 9000 ? ' …[cortado]' : '') };
    }
  };

  // "hoje", "amanhã", "ontem", DD/MM, DD/MM/AAAA ou AAAA-MM-DD → AAAA-MM-DD
  function dataIso(v) {
    const t = norm(v).trim();
    const d = new Date(); d.setHours(12, 0, 0, 0);
    const iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
    if (!t || t === 'hoje') return iso(d);
    if (t === 'amanha') { d.setDate(d.getDate() + 1); return iso(d); }
    if (t === 'ontem') { d.setDate(d.getDate() - 1); return iso(d); }
    let m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return t;
    m = t.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if (m) {
      const ano = m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : d.getFullYear();
      return `${ano}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
    }
    return null;
  }

  /* As ações não rodam sozinhas: viram um botão no chat e só acontecem no
     clique da pessoa. É a confirmação — e é também o que o navegador exige
     para abrir janela nova (OC, PDF das etiquetas) sem bloquear o pop-up. */
  let acoesPendentes = [];
  let seqAcao = 0;
  function oferecerAcao(rotulo, executar) {
    acoesPendentes.push({ id: 'a' + (++seqAcao), rotulo, executar, feita: false });
  }

  function rodarFerramenta(nome, entrada) {
    const f = FERRAMENTAS[nome];
    if (!f) return { erro: `Ferramenta desconhecida: ${nome}` };
    try {
      if (typeof STATE === 'undefined' || !STATE) return { erro: 'Os dados ainda não carregaram nesta tela.' };
      return f(entrada || {});
    } catch (e) {
      return { erro: 'Falhou ao consultar: ' + (e && e.message ? e.message : e) };
    }
  }

  /* ---------------- conversa ---------------- */
  let historico = [];
  let ocupado = false;
  let contaDoHistorico = null;
  try {
    const g = JSON.parse(sessionStorage.getItem('assistente:hist') || 'null');
    if (g && Array.isArray(g.h)) { historico = g.h; contaDoHistorico = g.conta || null; }
  } catch (e) { historico = []; }
  function guardar() {
    if (historico.length > HIST_MAX) {
      // corta pelo começo, sempre num "user" com texto, para não deixar
      // um tool_result órfão (a API recusa)
      let i = historico.length - HIST_MAX;
      while (i < historico.length && !(historico[i].role === 'user' && typeof historico[i].content === 'string')) i++;
      historico = historico.slice(i);
    }
    try { sessionStorage.setItem('assistente:hist', JSON.stringify({ conta: login(), h: historico })); } catch (e) { /* sem storage */ }
  }

  async function chamarServidor(corpo) {
    if (typeof supa === 'undefined' || !supa) return { erro: 'Sem conexão com o servidor' };
    const { data: { session } } = await supa.auth.getSession();
    if (!session) return { erro: 'Sua sessão expirou — entre de novo' };
    try {
      const r = await fetch(_servidorUrlAtivo + '/functions/v1/assistente', {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + session.access_token, 'Content-Type': 'application/json' },
        body: JSON.stringify(corpo)
      });
      if (r.status === 404) return { erro: 'O assistente só funciona no servidor da fábrica (e precisa estar instalado lá).' };
      const d = await r.json().catch(() => ({}));
      if (!r.ok) return { erro: d.erro || ('Erro ' + r.status) };
      return d;
    } catch (e) {
      return { erro: 'O servidor não respondeu: ' + (e && e.message ? e.message : e) };
    }
  }

  function contexto() {
    let nome = login();
    try { if (typeof _obsNomeLogin === 'function') nome = _obsNomeLogin(nome); } catch (e) { /* fica o e-mail */ }
    return { nome, papel: ehAdmin() ? 'admin' : 'usuario', hoje: new Date().toLocaleString('pt-BR') };
  }

  async function perguntar(texto) {
    if (ocupado || !texto.trim()) return;
    ocupado = true; atualizarBotao();
    acoesPendentes = [];
    historico.push({ role: 'user', content: texto.trim() });
    guardar(); desenhar();
    const pensando = mostrarPensando('Pensando…');
    try {
      for (let rodada = 0; rodada < RODADAS_MAX; rodada++) {
        const d = await chamarServidor({ mensagens: historico, contexto: contexto() });
        if (d.erro) { mostrarErro(d.erro); desfazerUltimaPergunta(); break; }
        historico.push({ role: 'assistant', content: d.content });
        const usos = (d.content || []).filter(b => b.type === 'tool_use');
        if (d.stop_reason !== 'tool_use' || !usos.length) break;
        pensando.textContent = 'Consultando ' + usos.map(u => NOMES_FERR[u.name] || u.name).join(', ') + '…';
        const resultados = usos.map(u => ({
          type: 'tool_result', tool_use_id: u.id, content: limitar(rodarFerramenta(u.name, u.input))
        }));
        historico.push({ role: 'user', content: resultados });
        if (rodada === RODADAS_MAX - 1) mostrarErro('A pergunta exigiu consultas demais. Tente algo mais específico.');
      }
    } finally {
      pensando.remove();
      ocupado = false;
      guardar(); desenhar(); atualizarBotao();
    }
  }
  // Pergunta que deu erro sai do histórico — senão a próxima carrega uma
  // conversa quebrada (user seguido de user) e a API recusa.
  function desfazerUltimaPergunta() {
    for (let i = historico.length - 1; i >= 0; i--) {
      if (historico[i].role === 'user' && typeof historico[i].content === 'string') {
        const t = historico[i].content;
        historico = historico.slice(0, i);
        const caixa = document.getElementById('assist-entrada');
        if (caixa && !caixa.value) caixa.value = t;
        return;
      }
    }
  }
  const NOMES_FERR = {
    resumo_status: 'o resumo por status', buscar_os: 'as OS', detalhe_os: 'a OS',
    listar_bases: 'as bases', consultar_base: 'os dados', tela_atual: 'a tela',
    preparar_folha: 'a folha'
  };

  /* ---------------- markdown simples (seguro) ---------------- */
  const escH = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function inline(s) {
    return escH(s)
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
      .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<i>$2</i>');
  }
  function markdown(txt) {
    const linhas = String(txt || '').split('\n');
    let html = '', i = 0;
    while (i < linhas.length) {
      const l = linhas[i];
      if (/^\s*\|/.test(l)) {
        const bloco = [];
        while (i < linhas.length && /^\s*\|/.test(linhas[i])) bloco.push(linhas[i++]);
        const cel = r => r.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        const corpo = bloco.filter(r => !/^\s*\|?\s*:?-{2,}/.test(r));
        html += '<div class="assist-tab"><table>' + corpo.map((r, j) =>
          '<tr>' + cel(r).map(c => j === 0 ? `<th>${inline(c)}</th>` : `<td>${inline(c)}</td>`).join('') + '</tr>').join('') + '</table></div>';
        continue;
      }
      if (/^\s*[-*•]\s+/.test(l)) {
        html += '<ul>';
        while (i < linhas.length && /^\s*[-*•]\s+/.test(linhas[i])) html += '<li>' + inline(linhas[i++].replace(/^\s*[-*•]\s+/, '')) + '</li>';
        html += '</ul>'; continue;
      }
      if (/^\s*\d+[.)]\s+/.test(l)) {
        html += '<ol>';
        while (i < linhas.length && /^\s*\d+[.)]\s+/.test(linhas[i])) html += '<li>' + inline(linhas[i++].replace(/^\s*\d+[.)]\s+/, '')) + '</li>';
        html += '</ol>'; continue;
      }
      const h = l.match(/^\s*#{1,4}\s+(.*)/);
      if (h) { html += `<p class="assist-h">${inline(h[1])}</p>`; i++; continue; }
      if (!l.trim()) { i++; continue; }
      html += '<p>' + inline(l) + '</p>'; i++;
    }
    return html;
  }

  /* ---------------- interface ---------------- */
  const CSS = `
  #assist-botao{position:fixed;right:18px;bottom:18px;z-index:9000;display:none;align-items:center;gap:8px;
    padding:10px 16px;border-radius:24px;border:1px solid var(--assist-bd);background:var(--assist-acento);color:#fff;
    font:600 14px 'IBM Plex Sans',system-ui,sans-serif;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.25)}
  #assist-botao:hover{filter:brightness(1.08)}
  #assist-painel{position:fixed;right:18px;bottom:18px;z-index:9001;width:400px;max-width:calc(100vw - 32px);
    height:600px;max-height:calc(100vh - 36px);display:none;flex-direction:column;background:var(--assist-fundo);
    color:var(--assist-txt);border:1px solid var(--assist-bd);border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.35);
    font:14px/1.45 'IBM Plex Sans',system-ui,sans-serif;overflow:hidden}
  #assist-painel.aberto{display:flex}
  .assist-cab{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid var(--assist-bd);background:var(--assist-cab)}
  .assist-cab b{flex:1;font-size:14px}
  .assist-cab button{background:none;border:1px solid var(--assist-bd);color:var(--assist-txt);border-radius:6px;padding:4px 8px;cursor:pointer;font-size:12px}
  .assist-msgs{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:10px}
  .assist-m{max-width:92%;padding:8px 11px;border-radius:10px;word-wrap:break-word}
  .assist-m p{margin:0 0 6px}.assist-m p:last-child{margin:0}
  .assist-m ul,.assist-m ol{margin:4px 0 6px 18px;padding:0}
  .assist-m code{font-family:'IBM Plex Mono',monospace;font-size:12px;background:var(--assist-cab);padding:0 3px;border-radius:3px}
  .assist-h{font-weight:700}
  .assist-eu{align-self:flex-end;background:var(--assist-acento);color:#fff;white-space:pre-wrap}
  .assist-ele{align-self:flex-start;background:var(--assist-bolha);border:1px solid var(--assist-bd)}
  .assist-erro{align-self:stretch;background:#fde8e8;color:#8a1c1c;border:1px solid #f0b4b4;font-size:13px}
  .assist-pensando{align-self:flex-start;font-size:12px;opacity:.75;font-style:italic}
  .assist-tab{overflow-x:auto;margin:4px 0}
  .assist-tab table{border-collapse:collapse;font-size:12px}
  .assist-tab th,.assist-tab td{border:1px solid var(--assist-bd);padding:3px 6px;text-align:left;white-space:nowrap}
  .assist-tab th{background:var(--assist-cab)}
  .assist-acoes{align-self:flex-start;display:flex;flex-direction:column;gap:6px}
  .assist-acoes button{padding:8px 12px;border-radius:8px;border:1px solid var(--assist-acento);background:var(--assist-fundo);
    color:var(--assist-acento);font:600 13px 'IBM Plex Sans',system-ui,sans-serif;cursor:pointer;text-align:left}
  .assist-acoes button:hover{background:var(--assist-acento);color:#fff}
  .assist-acoes button.feita{opacity:.55}
  .assist-vazio{opacity:.8;font-size:13px}
  .assist-vazio button{display:block;width:100%;text-align:left;margin:6px 0 0;padding:7px 9px;border-radius:8px;cursor:pointer;
    border:1px solid var(--assist-bd);background:var(--assist-bolha);color:var(--assist-txt);font:13px 'IBM Plex Sans',system-ui,sans-serif}
  .assist-pe{display:flex;gap:8px;padding:10px;border-top:1px solid var(--assist-bd);background:var(--assist-cab)}
  .assist-pe textarea{flex:1;resize:none;height:44px;max-height:120px;padding:8px;border-radius:8px;border:1px solid var(--assist-bd);
    background:var(--assist-fundo);color:var(--assist-txt);font:14px 'IBM Plex Sans',system-ui,sans-serif}
  .assist-pe button{padding:0 14px;border-radius:8px;border:none;background:var(--assist-acento);color:#fff;font-weight:600;cursor:pointer}
  .assist-pe button:disabled{opacity:.5;cursor:default}
  .assist-nota{font-size:11px;opacity:.6;padding:0 12px 8px;background:var(--assist-cab)}
  :root{--assist-fundo:#ffffff;--assist-txt:#1a1a1a;--assist-bd:#d6d9de;--assist-cab:#f4f5f7;--assist-bolha:#f7f8fa;--assist-acento:#c2410c}
  [data-tema="escuro"]{--assist-fundo:#1b1d21;--assist-txt:#e8e8e8;--assist-bd:#3a3e45;--assist-cab:#24272c;--assist-bolha:#2a2d33;--assist-acento:#d9662b}
  @media (max-width:640px){#assist-painel{right:0;bottom:0;width:100vw;max-width:100vw;height:100vh;max-height:100vh;border-radius:0}
    #assist-botao{right:12px;bottom:12px}}
  @media print{#assist-botao,#assist-painel{display:none!important}}`;

  const SUGESTOES = [
    'Quantas OS e peças estão em cada status agora?',
    'Quais OS estão cortando neste momento?',
    'Imprima a folha de OE de hoje.',
    'Quais OS da coleção Inverno 2026 ainda não começaram?'
  ];

  let painel, msgs, entrada, botaoEnviar, botaoAbrir;

  function montar() {
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);

    botaoAbrir = document.createElement('button');
    botaoAbrir.id = 'assist-botao';
    botaoAbrir.title = 'Perguntar ao assistente (Claude) sobre a produção';
    botaoAbrir.innerHTML = '<span aria-hidden="true">✦</span> Assistente';
    botaoAbrir.onclick = abrir;
    document.body.appendChild(botaoAbrir);

    painel = document.createElement('div');
    painel.id = 'assist-painel';
    painel.setAttribute('role', 'dialog');
    painel.setAttribute('aria-label', 'Assistente Claude');
    painel.innerHTML = `
      <div class="assist-cab"><b>✦ Assistente (Claude)</b>
        <button type="button" id="assist-nova" title="Apagar esta conversa e começar outra">Nova conversa</button>
        <button type="button" id="assist-fechar" title="Fechar (Esc)">✕</button></div>
      <div class="assist-msgs" id="assist-msgs"></div>
      <div class="assist-pe"><textarea id="assist-entrada" placeholder="Pergunte sobre OS, status, estoque…"></textarea>
        <button type="button" id="assist-enviar">Enviar</button></div>
      <div class="assist-nota">Consulta e prepara folhas para imprimir — não altera dados. Confira números importantes na tela.</div>`;
    document.body.appendChild(painel);

    msgs = painel.querySelector('#assist-msgs');
    entrada = painel.querySelector('#assist-entrada');
    botaoEnviar = painel.querySelector('#assist-enviar');
    painel.querySelector('#assist-fechar').onclick = fechar;
    painel.querySelector('#assist-nova').onclick = () => {
      if (ocupado) return;
      historico = []; guardar(); desenhar(); entrada.focus();
    };
    botaoEnviar.onclick = enviar;
    entrada.addEventListener('keydown', e => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); enviar(); }
      if (e.key === 'Escape') fechar();
      e.stopPropagation();   // os atalhos do app (F6, setas…) não disparam enquanto se digita aqui
    });
    msgs.addEventListener('click', e => {
      const b = e.target.closest('button[data-sug]');
      if (b) { entrada.value = b.getAttribute('data-sug'); enviar(); return; }
      const a = e.target.closest('button[data-acao]');
      if (a) {
        const acao = acoesPendentes.find(x => x.id === a.getAttribute('data-acao'));
        if (!acao) return;
        acao.feita = true;
        a.classList.add('feita');
        try { acao.executar(); }
        catch (err) { mostrarErro('Não consegui abrir: ' + (err && err.message ? err.message : err)); }
      }
    });

    desenhar();
    setInterval(vigiarLogin, 1500);
    vigiarLogin();
  }

  let contaVista = null;
  function vigiarLogin() {
    const logado = typeof currentUser !== 'undefined' && !!currentUser;
    botaoAbrir.style.display = logado && !painel.classList.contains('aberto') ? 'flex' : 'none';
    if (!logado && painel.classList.contains('aberto')) fechar();
    // Trocou de conta (saiu e outra pessoa entrou): a conversa anterior não é dela.
    if (logado) {
      const agora = login();
      const anterior = contaVista !== null ? contaVista : contaDoHistorico;
      if (anterior && anterior !== agora && historico.length) { historico = []; guardar(); desenhar(); }
      contaVista = agora;
    }
  }
  function abrir() { painel.classList.add('aberto'); botaoAbrir.style.display = 'none'; entrada.focus(); msgs.scrollTop = msgs.scrollHeight; }
  function fechar() { painel.classList.remove('aberto'); vigiarLogin(); }
  function enviar() { const t = entrada.value; if (!t.trim() || ocupado) return; entrada.value = ''; perguntar(t); }
  function atualizarBotao() { if (botaoEnviar) botaoEnviar.disabled = ocupado; }

  function desenhar() {
    if (!msgs) return;
    const partes = [];
    historico.forEach(m => {
      if (m.role === 'user' && typeof m.content === 'string') {
        partes.push(`<div class="assist-m assist-eu">${escH(m.content)}</div>`);
      } else if (m.role === 'assistant') {
        const txt = (Array.isArray(m.content) ? m.content : []).filter(b => b.type === 'text').map(b => b.text).join('\n\n').trim();
        if (txt) partes.push(`<div class="assist-m assist-ele">${markdown(txt)}</div>`);
      }
    });
    if (acoesPendentes.length) {
      partes.push('<div class="assist-acoes">' + acoesPendentes.map(a =>
        `<button type="button" data-acao="${a.id}" class="${a.feita ? 'feita' : ''}">${escH(a.rotulo)}</button>`).join('') + '</div>');
    }
    if (!partes.length) {
      partes.push(`<div class="assist-vazio">Olá! Eu consulto os dados do Gerador de OS e preparo folhas de OS, etiquetas, OE e ordens de compra para imprimir. Exemplos:
        ${SUGESTOES.map(s => `<button type="button" data-sug="${escH(s)}">${escH(s)}</button>`).join('')}</div>`);
    }
    msgs.innerHTML = partes.join('');
    msgs.scrollTop = msgs.scrollHeight;
  }
  function mostrarPensando(t) {
    const d = document.createElement('div'); d.className = 'assist-pensando'; d.textContent = t;
    msgs.appendChild(d); msgs.scrollTop = msgs.scrollHeight; return d;
  }
  function mostrarErro(t) {
    // desenhar() roda no finally e apagaria o aviso: ele volta logo depois
    setTimeout(() => {
      const d = document.createElement('div'); d.className = 'assist-m assist-erro'; d.textContent = t;
      msgs.appendChild(d); msgs.scrollTop = msgs.scrollHeight;
    }, 0);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', montar);
  else montar();
})();
