/* =====================================================================
   ASSISTENTE CLAUDE — janela de conversa dentro do Gerador-OS (19/10/2026)
   =====================================================================
   (Junior: chat do Claude dentro do programa, para TODOS os usuários, SÓ
   CONSULTA.)

   Arquivo separado do app.js de propósito: o app.js tem 2,4 MB e mexer nele
   para pendurar uma janela seria arriscar o resto. Este arquivo LÊ o que o
   app.js já deixou carregado (STATE, currentUser, currentRole, supa,
   _servidorUrlAtivo, _statusOS, produtosOS, STATUS_OS, contaVeArea).

   A ÚNICA GRAVAÇÃO É O PLANEJAMENTO (19/10/2026, Junior: "o Assistente deve
   ser capaz de imprimir todas as Ordens, preencher e salvar todos os
   planejamentos, perguntar ao usuário quando perceber conflitos. O Assistente
   não pode alterar status ou alterar condições de usuários como eleger admin
   ou rebaixar admin"). E ela não tem caminho próprio: a ferramenta `planejar`
   abre o MESMO formulário do Novo planejamento, preenche os campos como uma
   pessoa e chama o MESMO salvarPlanejamento — as contas de camadas, as
   validações e a trava de gravação são as do programa. Nada daqui muda status,
   aprova planejamento, apaga, ou mexe em conta e papel de usuário.

   E A OS CONJUGADA (19/10/2026, Junior: "autonomia para criar OS com número
   que seja conjugada seguindo outra OS... e imprimir essa OS logo em
   seguida"): `criar_os_conjugada`, pela regra da grade ou pelo formulário +
   salvarOS e a amarra do "⋯ → conjugar". É a única OS com número que nasce
   daqui, sempre presa a uma OS que já existe, e a folha vai para a impressão
   sem esperar clique.

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
    expedicaoCargas: 'expedicao', expedicaoJanelas: 'expedicao', expedicaoExcecoes: 'expedicao',
    operacoes: 'operacoes', costuraProd: 'costura', compraOCs: 'compra', compraPlano: 'compra'
  };
  const BASES_NUNCA = new Set(['_device']);

  /* A CONTA DO ASSISTENTE (19/10/2026, Junior: "configure uma conta de acesso
     para que o Assistente, por meio dessa conta, possa gerar folhas de OS, OE,
     ordens de compra...").

     As FOLHAS (OS, etiquetas, OE, OC) precisam de DUAS licenças: a visão da
     conta "claude.assistent" (que o admin cria e ajusta em Configurações → Contas de
     acesso, como qualquer outra; sem marcação ela vê tudo) E a visão de quem
     está logado. A conta do assistente funciona como um limite a mais do que o
     chat pode fazer — nunca como atalho para alguém abrir pelo chat uma folha
     que a própria conta não abre. Nada aqui grava dados nem usa senha. */
  // A conta que o Junior criou em 19/10/2026 (login claude.assistent@diverse.local).
  // Precisa ter só VISÃO: edição marcada também conta como visão em contaVeArea.
  const CONTA_ASSISTENTE = 'claude.assistent';
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
      const AREA = { os: 'os', etiquetas: 'os', oe: 'expedicao', oc: 'compra' };
      if (!AREA[tipo]) return { erro: 'tipo deve ser os, etiquetas, oe ou oc' };
      if (!assistenteVe(AREA[tipo])) {
        return { sem_acesso: true, mensagem: `A conta do assistente ("${CONTA_ASSISTENTE}") não tem visão da área "${AREA[tipo]}". O admin libera em Configurações → Contas de acesso.` };
      }
      if (!veArea(AREA[tipo])) return SEM_ACESSO(AREA[tipo]);
      const alvos = (Array.isArray(p.alvos) && p.alvos.length ? p.alvos : [p.alvo]).map(String).filter(Boolean);
      if (!alvos.length) return { erro: 'Diga qual folha: alvo ou alvos.' };
      if (alvos.length > FOLHAS_MAX) return { erro: `São ${alvos.length} folhas; o máximo por pedido é ${FOLHAS_MAX}. Divida em partes.` };

      const achadas = [], erros = [];
      alvos.forEach(a => { const f = folhaDe(tipo, a); if (f.erro) erros.push(f.erro); else achadas.push(f); });
      if (!achadas.length) return { erro: erros.join(' ') };

      if (achadas.length === 1) {
        const f = achadas[0];
        const imprimir = !!p.imprimir || f.sempreImprime;
        const rotulo = `${imprimir ? '🖨 Imprimir' : '📄 Abrir'} ${f.nome}`;
        oferecerAcao(rotulo, () => {
          if (f.janelaPropria) { f.abrir(); return; }
          fechar(); f.abrir();
          if (imprimir) setTimeout(() => window.print(), 700);
        });
        return { botao_mostrado: rotulo, ...f.info, nao_achadas: erros.length ? erros : undefined };
      }

      // VÁRIAS. A folha da OS e a da OE moram numa tela do próprio programa: um
      // botão só abre uma e imprime, abre a seguinte e imprime, até a última.
      // Etiquetas e OC abrem janela nova, e o navegador só deixa abrir janela no
      // clique da pessoa — por isso é um botão para cada uma.
      if (achadas.every(f => !f.janelaPropria)) {
        const rotulo = `🖨 Imprimir todas (${achadas.length}) — ${achadas.map(f => f.curto).join(', ')}`;
        oferecerAcao(rotulo, () => imprimirEmSequencia(achadas));
        return { botao_mostrado: rotulo, folhas: achadas.map(f => f.nome), nao_achadas: erros.length ? erros : undefined,
          aviso: 'Abre o diálogo de impressão uma vez por folha, na ordem da lista.' };
      }
      achadas.forEach(f => oferecerAcao(`🖨 ${f.nome}`, () => f.abrir()));
      return { botoes_mostrados: achadas.map(f => f.nome), nao_achadas: erros.length ? erros : undefined,
        aviso: 'Cada uma abre numa janela própria: um clique por folha.' };
    },

    async planejar(p) {
      return planejar(p);
    },

    async criar_os_conjugada(p) {
      return criarConjugada(p);
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

  const esperar = ms => new Promise(r => setTimeout(r, ms));

  /* ---------------- folhas ---------------- */
  const FOLHAS_MAX = 60;

  // Uma folha pedida → o que abrir e como chamar. Erro se não existir.
  function folhaDe(tipo, alvo) {
    if (tipo === 'os' || tipo === 'etiquetas') {
      const n = numOS(alvo);
      const o = (STATE.ordens || []).find(x => numOS(x.os) === n);
      if (!o) return { erro: `Não achei a OS ${alvo}.` };
      if (tipo === 'etiquetas') {
        return { nome: `etiquetas da OS ${o.os}`, curto: o.os, janelaPropria: true, sempreImprime: true,
          abrir: () => imprimirEtiquetasPdf(o.id),
          info: { os: resumoOS(o), aviso: 'As etiquetas abrem como PDF numa aba nova; a impressão é feita por lá.' } };
      }
      return { nome: `folha da OS ${o.os}`, curto: o.os, abrir: () => verOS(o.id), info: { os: resumoOS(o) } };
    }
    if (tipo === 'oe') {
      const data = dataIso(alvo);
      if (!data) return { erro: `Não entendi a data "${alvo}". Use AAAA-MM-DD, DD/MM, "hoje" ou "amanhã".` };
      const rot = data.split('-').reverse().join('/');
      return { nome: `folha de OE de ${rot}`, curto: rot.slice(0, 5),
        abrir: () => {
          expPlanoModo = 'dia';
          expPlanoAncora = data;
          try { sessionStorage.setItem('gos:exp:modo', 'dia'); sessionStorage.setItem('gos:exp:ancora', data); } catch (e) { /* sem storage */ }
          goto('print-expedicao');
        },
        info: { data, aviso: 'A folha só traz as expedições com carga alocada; dia sem carga sai vazio.' } };
    }
    // oc — a janela da OC já abre o diálogo de impressão sozinha.
    const n = String(alvo || '').replace(/\D/g, '').replace(/^0+/, '');
    const ocs = Array.isArray(STATE.compraOCs) ? STATE.compraOCs : [];
    const oc = ocs.find(x => String(x.numero || '').replace(/\D/g, '').replace(/^0+/, '') === n);
    if (!oc) return { erro: `Não achei a ordem de compra ${alvo} (as últimas: ${ocs.slice(-8).map(x => x.numero).join(', ')}).` };
    return { nome: `${oc.numero}`, curto: oc.numero, janelaPropria: true, sempreImprime: true,
      abrir: () => imprimirOC(oc.id), info: { oc: enxugar(oc, 1, 200, 20) } };
  }

  // window.print() segura o código até a pessoa fechar o diálogo; então é só
  // abrir a próxima folha, dar tempo de ela se desenhar e imprimir de novo.
  async function imprimirEmSequencia(folhas) {
    fechar();
    for (const f of folhas) {
      f.abrir();
      await esperar(900);
      window.print();
      await esperar(200);
    }
  }

  /* ---------------- planejamento ----------------
     Pelo formulário de verdade, como o "Planejar" do Relatório de demanda
     (_demAbrirParte): escolhe desenho, linha e grade nos próprios selects,
     dispara os mesmos onchange, e salva por salvarPlanejamento.

     OS CONFLITOS SÃO CONFERIDOS ANTES, e com eles nada é salvo: o formulário
     fica preenchido na tela e o Claude pergunta. Só com o sim da pessoa ele
     chama de novo com `aceitar`. As perguntas que o próprio programa faz ao
     salvar (medida sem risco, camadas acima do limite) são respondidas com o
     que a pessoa aceitou no chat — e qualquer OUTRA pergunta que aparecer é
     respondida "não": o salvamento para e a pergunta volta para o chat. */
  const ROTULO_CONFLITO = {
    grade_de_outro_produto: 'grade de outro produto', camadas_acima_do_limite: 'camadas acima do limite',
    medida_sem_risco: 'medida sem risco', planejamento_repetido: 'planejamento repetido',
    os_igual_em_producao: 'OS igual em produção', data_no_passado: 'data passada',
    ativa_encerrada: 'OS ativa encerrada', conjugada_repetida: 'conjugada repetida', falta_de_tecido: 'falta tecido'
  };
  const CONFIRMA_DO_PROGRAMA = [
    { codigo: 'medida_sem_risco', casa: /não veio de um risco/i },
    { codigo: 'camadas_acima_do_limite', casa: /camadas, mas o limite/i },
    { codigo: 'falta_de_tecido', casa: /Não há tecido disponível/i }
  ];

  function acharDesenho(texto) {
    const t = norm(texto).trim();
    const ds = STATE.desenhos || [];
    const exato = ds.filter(d => norm(d.codigo) === t);
    if (exato.length === 1) return { desenho: exato[0] };
    const termos = t.split(/\s+/).filter(Boolean);
    const casam = ds.filter(d => {
      const h = norm(typeof _textoBuscaDesenho === 'function' ? _textoBuscaDesenho(d) : [d.codigo, d.desc].join(' '));
      return termos.every(x => h.includes(x));
    });
    if (!casam.length) return { erro: `Não achei desenho com "${texto}".` };
    if (casam.length > 1) {
      return { pergunta: 'Mais de um desenho combina — pergunte à pessoa qual.',
        opcoes: casam.slice(0, 20).map(d => _rotuloDesenhoOS(d)) };
    }
    return { desenho: casam[0] };
  }

  const opcoesDeGrade = sel => Array.from(sel.options).filter(o => o.value).map(o => o.text).slice(0, 30);

  /* O formulário da Nova OS já aberto (OS ou planejamento): desenho, linha,
     grade, camadas, data e obs nos próprios campos, com os mesmos onchange.
     Devolve `parar` quando falta algo — é o que volta para o Claude. */
  async function preencherFormulario(p, des, exigirGrade) {
    const el = id => document.getElementById(id);
    if (des) {
      const sd = el('f-desenho');
      sd.value = des.id;
      if (sd.value !== des.id) return { parar: { erro: `O desenho ${des.codigo} não está na lista do formulário.` } };
      previewDesenhoSelecionado();
      sincCodigoDesenho('desenho');
      preencherDropdownGradesOS();
    }

    const sg = el('f-grade-preset');
    let grade = null;
    if (p.grade) {
      const t = norm(p.grade).trim();
      const termos = t.split(/\s+/).filter(Boolean);
      const todas = STATE.grades || [];
      let casam = todas.filter(g => norm(g.nome).trim() === t);
      if (!casam.length) casam = todas.filter(g => termos.every(x => norm(g.nome).includes(x)));
      if (casam.length > 1) {
        // Entre várias, as que o formulário oferece para este desenho primeiro.
        const naLista = new Set(Array.from(sg.options).map(o => o.value));
        const doDesenho = casam.filter(g => naLista.has(g.id));
        if (doDesenho.length) casam = doDesenho;
      }
      if (!casam.length) return { parar: { erro: `Não achei grade com "${p.grade}".`, grades_para_este_desenho: opcoesDeGrade(sg) } };
      if (casam.length > 1) return { parar: { pergunta: 'Mais de uma grade combina — pergunte à pessoa qual.', opcoes: casam.slice(0, 20).map(g => g.nome) } };
      grade = casam[0];
    } else if (exigirGrade) {
      return { parar: { pergunta: 'Falta a grade — pergunte à pessoa qual.', preenchido_na_tela: true,
        grades_para_este_desenho: opcoesDeGrade(sg) } };
    }
    if (grade) {
      const ls = el('f-linha-tipo');
      const linha = _linhaTipoGrade(grade);
      if (ls && linha) { ls.value = linha; filtrarLinhaGradesOS(linha); }
      preencherDropdownGradesOS(grade.id);
      sg.value = grade.id;
      aplicarGradePreset();
    }

    if (p.camadas != null) { el('f-enf-camadas').value = Math.max(1, parseInt(p.camadas, 10) || 1); calcularAlvoDeCamadas(); }
    if (p.grades_iguais != null) { el('f-plan-vezes').value = Math.max(1, parseInt(p.grades_iguais, 10) || 1); atualizarCalculosEnfesto(); }
    if (p.data) { const d = dataIso(p.data); if (!d) return { parar: { erro: `Não entendi a data "${p.data}".` } }; el('f-data').value = d; }
    if (p.tipo_producao) el('f-tipo-producao').value = p.tipo_producao === 'puxada' ? 'puxada' : 'empurrada';
    if (p.obs != null) el('f-obs').value = String(p.obs);
    await esperar(80);

    const gradeId = sg.value;
    const g = (STATE.grades || []).find(x => x.id === gradeId);
    const dados = coletaOS();
    const gr = dados.grade || {};
    if (['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].some(k => gr[k] > 0) && ['t2', 't4', 't6', 't8', 't10', 't12', 't14', 't16'].some(k => gr[k] > 0)) {
      return { parar: { salvo: false, preenchido_na_tela: true, erro: 'A grade mistura P ao G3 com 2 ao 16 — isso o programa não aceita. Pergunte qual linha vale.' } };
    }
    return { des, g, gradeId, desenhoId: el('f-desenho').value, dados,
      cam: parseInt(el('f-enf-camadas').value, 10) || 0, dataForm: el('f-data').value };
  }

  // Os conflitos de qualquer formulário preenchido. `fora` = OS que não contam
  // como "igual em produção" (a ativa de uma conjugada é, de propósito, igual).
  function conflitosDoFormulario(f, fora) {
    const conflitos = [];
    if (f.g && f.des) {
      const motivo = _motivoGradeForaDoDropdownOS(f.g, _ctxDropdownGradesOS([]));
      if (motivo) conflitos.push({ codigo: 'grade_de_outro_produto',
        texto: `A grade ${f.g.nome} não é das que o formulário oferece para este desenho (${(_MOTIVO_GRADE_LABEL || {})[motivo] || motivo}).` });
    }
    const { limite } = calcularLimiteCamadas();
    if (f.cam > 0 && f.cam > limite) conflitos.push({ codigo: 'camadas_acima_do_limite', texto: `${f.cam} camadas, acima do limite do tecido (${limite}).` });
    const semProva = f.gradeId ? fasesSemProvaDeMedida(f.gradeId) : [];
    if (semProva.length) conflitos.push({ codigo: 'medida_sem_risco',
      texto: `A medida de enfesto destas fases foi digitada à mão, sem PDF de risco: ${semProva.map(x => x.nome).join(', ')}.` });
    const naoConta = new Set(fora || []);
    const emProducao = (STATE.ordens || []).filter(o => !naoConta.has(o.id) && o.desenhoId === f.desenhoId && o.gradeId === f.gradeId
      && !['estoque', 'cancelado'].includes(statusDe(o)));
    if (emProducao.length) conflitos.push({ codigo: 'os_igual_em_producao',
      texto: `Já há OS do mesmo desenho e grade em produção: ${emProducao.slice(0, 10).map(o => `${o.os} (${rotuloStatus(statusDe(o))})`).join(', ')}.` });
    if (f.dataForm && f.dataForm < dataIso('hoje')) conflitos.push({ codigo: 'data_no_passado', texto: `A data ${f.dataForm.split('-').reverse().join('/')} já passou.` });
    return conflitos;
  }

  // As perguntas que o programa faz ao salvar: sim só para o que foi aceito;
  // qualquer outra é "não", e o texto dela volta para o chat.
  async function salvarRespondendo(salvar, aceitos) {
    const confirmOriginal = window.confirm;
    const perguntas = [];
    window.confirm = msg => {
      const c = CONFIRMA_DO_PROGRAMA.find(x => x.casa.test(String(msg)));
      if (c && aceitos.has(c.codigo)) return true;
      perguntas.push(String(msg).slice(0, 400));
      return false;
    };
    try { await salvar(); }
    finally { window.confirm = confirmOriginal; }
    return perguntas;
  }

  /* ---------------- OS conjugada ----------------
     (19/10/2026, Junior: "O assistente deve ter autonomia para criar OS com
     número que seja conjugada seguindo outra OS. Ela deve ter autonomia para
     imprimir essa OS logo em seguida.")

     É a ÚNICA OS com número que o assistente cria, e sempre amarrada a uma OS
     que já existe. Os dois caminhos são os do programa:
     · PELA GRADE — a grade da OS ativa tem conjugada no cadastro e a segunda
       OS ainda não existe: aplicarRegraConjugadaSeAplicavel, a mesma de quando
       a OS é salva (número de baixo; não reserva pano, é o mesmo enfesto).
     · À MÃO — a pessoa diz a grade (e o desenho, se não for o mesmo): a Nova
       OS pelo formulário e pelo salvarOS (número, pano dela, etiqueta) e depois
       a marca conjugadaStatusPaiId, a mesma do "⋯ → conjugar": o status passa
       a seguir a ativa. Nenhum status é carimbado agora — alinhar seria mudar
       status, e isso o assistente não faz.
     A folha é impressa logo a seguir, sem esperar clique. */
  async function criarConjugada(p) {
    if (typeof contaTemAcesso !== 'function' || !contaTemAcesso('os', CONTA_ASSISTENTE, 'usuario')) {
      return { sem_acesso: true, mensagem: `A conta do assistente ("${CONTA_ASSISTENTE}") não tem edição em Ordens de Serviço. O admin libera em Configurações → Contas de acesso.` };
    }
    if (typeof podeGravar === 'function' && !podeGravar()) {
      return { erro: 'Esta tela não grava (celular ou fora do servidor da fábrica). Use um computador da fábrica.' };
    }
    if (!(ehAdmin() || temAcesso('os'))) {
      return { sem_acesso: true, mensagem: 'Esta conta não pode criar OS. Peça ao admin.' };
    }
    let ativa = (STATE.ordens || []).find(x => numOS(x.os) === numOS(p.os_ativa));
    if (!ativa) return { erro: `Não achei a OS ${p.os_ativa}.` };
    const notas = [];
    // Quem manda no grupo é uma só: sobe para ela, como a tela de conjugar.
    const pai = ativa.conjugadaPaiId ? (STATE.ordens || []).find(x => x.id === ativa.conjugadaPaiId) : _ativaStatusDaOS(ativa);
    if (pai) { notas.push(`A OS ${ativa.os} já segue a OS ${pai.os}; a nova segue a ${pai.os}, que é quem manda no grupo.`); ativa = pai; }

    const aceitos = new Set((Array.isArray(p.aceitar) ? p.aceitar : []).map(String));
    const conflitos = [];
    const st = statusDe(ativa);
    if (['estoque', 'cancelado'].includes(st)) conflitos.push({ codigo: 'ativa_encerrada', texto: `A OS ${ativa.os} já está em ${rotuloStatus(st)}.` });

    const gAtiva = (STATE.grades || []).find(g => g.id === ativa.gradeId);
    const gConj = gradeConjugaAlguma(ativa.gradeId) ? (STATE.grades || []).find(g => g.id === gAtiva.conjugadaGradeId) : null;
    const termosGrade = norm(p.grade || '').trim().split(/\s+/).filter(Boolean);
    const pelaGrade = gConj && !p.desenho && (!termosGrade.length || termosGrade.every(x => norm(gConj.nome).includes(x)));

    const pendentesOuBotao = (lista, preenchido) => {
      const pendentes = lista.filter(c => !aceitos.has(c.codigo));
      if (!pendentes.length) return null;
      const todos = [...aceitos, ...pendentes.map(c => c.codigo)];
      oferecerAcao(`➕ Criar assim mesmo (${pendentes.map(c => ROTULO_CONFLITO[c.codigo] || c.codigo).join(', ')})`, async () => {
        const r = await criarConjugada(Object.assign({}, p, { aceitar: todos }));
        anotarNoChat(r.criada
          ? `✔ OS ${r.os.os} criada, conjugada à OS ${r.segue_a_os}.${r.impressa ? ' A folha foi para a impressão.' : ''}`
          : `✖ Não criei: ${r.erro || (r.conflitos || []).map(c => c.texto).join(' ') || 'o programa recusou'}`
            + (r.perguntas_do_programa ? `\n\n${r.perguntas_do_programa.join('\n\n')}` : ''));
      }, true);
      return { criada: false, preenchido_na_tela: preenchido, conflitos: pendentes, botao_mostrado: 'Criar assim mesmo',
        instrucao: 'Nada foi criado ainda. Explique os conflitos em poucas linhas e diga que a pessoa pode clicar em "Criar assim mesmo" ou responder sim. Com o sim dela por escrito, chame de novo com aceitar.' };
    };
    const marcar = o => { o.viaAssistente = { conta: CONTA_ASSISTENTE, pedidoPor: login(), em: new Date().toISOString(), conflitosAceitos: [...aceitos] }; };

    let nova = null, caminho = '';
    const outras = [];
    if (pelaGrade) {
      if (!deveGerarConjugada(ativa)) {
        const ja = (STATE.ordens || []).find(o => o.id === ativa.conjugadaId);
        return { criada: false, ja_existe: ja ? resumoOS(ja) : undefined,
          pergunta: `A conjugada da grade da OS ${ativa.os}${ja ? ` já existe: OS ${ja.os}` : ' não pôde ser gerada'}. Para outra OS, à mão, pergunte à pessoa a grade (e o desenho, se não for o mesmo).` };
      }
      const r = pendentesOuBotao(conflitos, { pela_grade: gConj.nome });
      if (r) return r;
      nova = await aplicarRegraConjugadaSeAplicavel(ativa);
      if (!nova) return { criada: false, erro: 'O programa não gerou a conjugada — veja a mensagem que apareceu na tela.' };
      caminho = `pela grade (${gConj.nome}): mesmo enfesto, não reserva pano`;
      marcar(nova);
      try { await saveState('ordens'); } catch (e) { /* a OS já está salva; a marca vai no próximo */ }
    } else {
      if (!p.grade) {
        return { pergunta: `A grade da OS ${ativa.os}${gAtiva ? ` (${gAtiva.nome})` : ''} não tem conjugada no cadastro. Pergunte à pessoa a grade da nova OS (e o desenho, se não for o mesmo).` };
      }
      let des = null;
      if (p.desenho) {
        const r = acharDesenho(p.desenho);
        if (!r.desenho) return r;
        des = r.desenho;
      } else {
        des = (STATE.desenhos || []).find(d => d.id === ativa.desenhoId) || null;
        if (!des) return { pergunta: 'O desenho da OS ativa não está no cadastro — pergunte à pessoa qual desenho usar.' };
      }
      // A Nova OS, como o botão da tela.
      osEditId = null;
      _osPlanAbrir = false;
      goto('nova-os');
      await esperar(300);
      if (_osPlanModo || osEditId) return { erro: 'Não consegui abrir o formulário da Nova OS.' };
      const f = await preencherFormulario(p, des, true);
      if (f.parar) return f.parar;
      const grupo = [ativa.id, ..._conjugadasManuaisDaOS(ativa).map(o => o.id), ativa.conjugadaId].filter(Boolean);
      conflitos.push(...conflitosDoFormulario(f, grupo));
      const repetida = (STATE.ordens || []).filter(o => grupo.includes(o.id) && o.id !== ativa.id
        && o.desenhoId === f.desenhoId && o.gradeId === f.gradeId);
      if (repetida.length) conflitos.push({ codigo: 'conjugada_repetida',
        texto: `A OS ${ativa.os} já tem conjugada com o mesmo desenho e a mesma grade: ${repetida.map(o => o.os).join(', ')}.` });
      try {
        const falta = faltaDeTecidoParaOS(f.dados);
        if (falta.length) conflitos.push({ codigo: 'falta_de_tecido',
          texto: `Falta tecido: ${falta.map(x => `${x.tecidoNome} ${x.corNome || ''} (faltam ${Number(x.falta || 0).toFixed(1).replace('.', ',')} kg)`).join('; ')}.` });
      } catch (e) { /* a prova do salvarOS pergunta de novo */ }
      const preenchido = { desenho: _rotuloDesenhoOS(des), grade: f.g ? f.g.nome : undefined, camadas: f.cam || undefined,
        numero: document.getElementById('f-os').value, data: f.dataForm || undefined };
      const r = pendentesOuBotao(conflitos, preenchido);
      if (r) return r;

      const antes = new Set((STATE.ordens || []).map(o => o.id));
      const perguntas = await salvarRespondendo(salvarOS, aceitos);
      const novas = (STATE.ordens || []).filter(o => !antes.has(o.id));
      nova = novas.find(o => !o.conjugadaPaiId) || null;
      if (!nova) {
        return { criada: false, preenchido_na_tela: preenchido, perguntas_do_programa: perguntas.length ? perguntas : undefined,
          erro: perguntas.length ? 'O programa fez uma pergunta ao salvar: leve-a para a pessoa.' : 'O programa não salvou — veja a mensagem que apareceu na tela.' };
      }
      // A grade da nova OS pode ter conjugada própria: o salvarOS já a gerou.
      novas.filter(o => o !== nova).forEach(o => { marcar(o); outras.push(o.os); });
      nova.conjugadaStatusPaiId = ativa.id;
      marcar(nova);
      caminho = 'à mão: reserva o pano dela, e o status segue a OS ativa';
      try { await saveState('ordens'); }
      catch (e) { return { criada: true, os: resumoOS(nova), erro: 'A OS foi salva, mas a amarra com a ativa não gravou. Conjugue pelo "⋯ → conjugar" da lista.' }; }
    }

    const ret = { criada: true, caminho, os: resumoOS(nova), segue_a_os: ativa.os,
      status_da_ativa: rotuloStatus(st),
      aviso_status: statusDe(nova) !== st
        ? `A nova OS está em ${rotuloStatus(statusDe(nova))} e a ${ativa.os} em ${rotuloStatus(st)}: ela passa a seguir a partir da próxima mudança de status da ${ativa.os}. Alinhar agora é com quem carimba status.`
        : undefined,
      tambem_gerada_pela_grade: outras.length ? outras : undefined, notas: notas.length ? notas : undefined };
    if (p.imprimir !== false) {
      const folha = folhaDe('os', nova.os);
      if (!folha.erro) {
        // O salvarOS termina em goto('lista-os'), que escreve #lista-os no
        // endereço; o hashchange chega DEPOIS e, como a folha (print) não tem
        // endereço, voltava para a lista — e a impressão saía da lista de OS
        // (19/10/2026). Espera esse evento passar, e só imprime com a folha na tela.
        await esperar(400);
        folha.abrir();
        await esperar(900);
        const naTela = document.querySelector('section.page:not(.hidden)')?.dataset?.page;
        if (naTela === 'print' && printOsAtual && printOsAtual.id === nova.id) {
          window.print();
          ret.impressa = true;
        } else {
          ret.impressa = false;
          ret.aviso_impressao = 'A folha não ficou na tela; nada foi impresso. Diga à pessoa para clicar no botão "Imprimir de novo".';
        }
        oferecerAcao(`🖨 Imprimir de novo a folha da OS ${nova.os}`, () => { fechar(); folha.abrir(); setTimeout(() => window.print(), 700); });
      }
    }
    return ret;
  }

  async function planejar(p) {
    const acao = p.acao === 'editar' ? 'editar' : 'novo';
    // Duas licenças, como nas folhas: a conta do assistente precisa de EDIÇÃO no
    // Planejamento, e quem pede precisa poder planejar.
    if (typeof contaTemAcesso !== 'function' || !contaTemAcesso('planejamento', CONTA_ASSISTENTE, 'usuario')) {
      return { sem_acesso: true, mensagem: `A conta do assistente ("${CONTA_ASSISTENTE}") não tem edição no Planejamento de produção. O admin libera em Configurações → Contas de acesso.` };
    }
    if (typeof podeGravar === 'function' && !podeGravar()) {
      return { erro: 'Esta tela não grava (celular ou fora do servidor da fábrica). Use um computador da fábrica.' };
    }
    if (!(ehAdmin() || temAcesso('planejamento') || temAcesso('os'))) {
      return { sem_acesso: true, mensagem: 'Esta conta não pode planejar a produção. Peça ao admin.' };
    }

    let plan = null;
    if (acao === 'editar') {
      plan = _planLista().find(x => x.id === p.planejamento_id);
      if (!plan) return { erro: 'Não achei esse planejamento.',
        existentes: _planLista().slice(-15).map(x => ({ id: x.id, resumo: _planResumo(x) })) };
    }
    let des = null;
    if (p.desenho) {
      const r = acharDesenho(p.desenho);
      if (!r.desenho) return r;
      des = r.desenho;
    } else if (acao === 'novo') {
      return { pergunta: 'Falta o desenho técnico — pergunte à pessoa qual (código ou descrição).' };
    }

    // Abre o formulário, como os botões da tela.
    if (acao === 'novo') novoPlanejamento(); else editarPlanejamento(plan.id);
    await esperar(300);
    if (!_osPlanModo) return { erro: 'Não consegui abrir o formulário de planejamento.' };
    const el = id => document.getElementById(id);
    const f = await preencherFormulario(p, des, acao === 'novo');
    if (f.parar) return f.parar;
    const { g, gradeId, desenhoId, dados, cam } = f;
    const dataPlan = f.dataForm;

    // ---- os conflitos ----
    const conflitos = conflitosDoFormulario(f);
    const iguais = _planLista().filter(x => x.id !== (plan && plan.id) && x.desenhoId === desenhoId && x.gradeId === gradeId);
    if (iguais.length) conflitos.push({ codigo: 'planejamento_repetido',
      texto: `Já existe planejamento com o mesmo desenho e a mesma grade: ${iguais.slice(0, 5).map(_planResumo).join(' / ')}.` });

    const aceitos = new Set((Array.isArray(p.aceitar) ? p.aceitar : []).map(String));
    const pendentes = conflitos.filter(c => !aceitos.has(c.codigo));
    const preenchido = { desenho: des ? _rotuloDesenhoOS(des) : undefined, grade: g ? g.nome : undefined,
      camadas: cam || undefined, grades_iguais: parseInt(el('f-plan-vezes').value, 10) || 1, data: dataPlan || undefined,
      produtos: (() => { try { return produtosOS(dados) * (parseInt(el('f-plan-vezes').value, 10) || 1); } catch (e) { return undefined; } })() };
    if (pendentes.length) {
      // A PERGUNTA TAMBÉM É UM BOTÃO (19/10/2026, Junior: "falta o assistente,
      // além de preencher, salvar"). Esperar o "sim" voltar pela conversa
      // deixava o planejamento só preenchido na tela; o clique é a resposta.
      const todos = [...aceitos, ...pendentes.map(c => c.codigo)];
      oferecerAcao(`💾 Salvar assim mesmo (${pendentes.map(c => ROTULO_CONFLITO[c.codigo] || c.codigo).join(', ')})`, async () => {
        const r = await planejar(Object.assign({}, p, { aceitar: todos }));
        anotarNoChat(r.salvo
          ? `✔ Planejamento salvo: ${r.resumo}. Aguarda a aprovação do admin no Planejamento de produção.`
          : `✖ Não salvei: ${r.erro || (r.conflitos || []).map(c => c.texto).join(' ') || 'o programa recusou'}`
            + (r.perguntas_do_programa ? `\n\n${r.perguntas_do_programa.join('\n\n')}` : ''));
      }, true);
      return { salvo: false, preenchido_na_tela: preenchido, conflitos: pendentes,
        botao_mostrado: 'Salvar assim mesmo',
        instrucao: 'Nada foi salvo ainda. Explique os conflitos em poucas linhas e diga que a pessoa pode clicar em "Salvar assim mesmo" ou responder sim. Com o sim dela por escrito, chame de novo com aceitar.' };
    }

    // ---- salvar pelo caminho do programa ----
    const antes = new Set(_planLista().map(x => x.id));
    const editadoAntes = plan ? plan.editadoEm : null;
    const perguntasNovas = await salvarRespondendo(salvarPlanejamento, aceitos);

    const salvo = acao === 'novo'
      ? _planLista().find(x => !antes.has(x.id))
      : (_planLista().find(x => x.id === plan.id && x.editadoEm && x.editadoEm !== editadoAntes) || null);
    if (!salvo) {
      return { salvo: false, preenchido_na_tela: preenchido,
        perguntas_do_programa: perguntasNovas.length ? perguntasNovas : undefined,
        erro: perguntasNovas.length ? 'O programa fez uma pergunta ao salvar: leve-a para a pessoa.' : 'O programa não salvou — veja a mensagem que apareceu na tela.' };
    }
    // Quem fez pelo chat fica escrito no próprio planejamento.
    salvo.viaAssistente = { conta: CONTA_ASSISTENTE, pedidoPor: login(), em: new Date().toISOString(),
      conflitosAceitos: [...aceitos] };
    try { await saveState('osPlanejadas'); } catch (e) { /* o planejamento já está salvo; a marca vai no próximo */ }
    return { salvo: true, acao, id: salvo.id, resumo: _planResumo(salvo), ...preenchido,
      proximo_passo: 'Fica na lista do Planejamento de produção até o admin aprovar (a aprovação gera a OS com número).' };
  }

  /* As ações não rodam sozinhas: viram um botão no chat e só acontecem no
     clique da pessoa. É a confirmação — e é também o que o navegador exige
     para abrir janela nova (OC, PDF das etiquetas) sem bloquear o pop-up. */
  let acoesPendentes = [];
  let seqAcao = 0;
  // O resultado de um botão clicado: fica embaixo dos botões até a próxima pergunta.
  let notasDoChat = [];
  function anotarNoChat(t) { notasDoChat.push(String(t)); desenhar(); }
  function oferecerAcao(rotulo, executar, umaVez) {
    acoesPendentes.push({ id: 'a' + (++seqAcao), rotulo, executar, feita: false, umaVez: !!umaVez });
  }

  async function rodarFerramenta(nome, entrada) {
    const f = FERRAMENTAS[nome];
    if (!f) return { erro: `Ferramenta desconhecida: ${nome}` };
    try {
      if (typeof STATE === 'undefined' || !STATE) return { erro: 'Os dados ainda não carregaram nesta tela.' };
      return await f(entrada || {});
    } catch (e) {
      return { erro: 'Falhou: ' + (e && e.message ? e.message : e) };
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
    notasDoChat = [];
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
        // Uma de cada vez: o planejamento usa o formulário, que é um só.
        const resultados = [];
        for (const u of usos) {
          resultados.push({ type: 'tool_result', tool_use_id: u.id, content: limitar(await rodarFerramenta(u.name, u.input)) });
        }
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
    preparar_folha: 'a folha', planejar: 'o planejamento', criar_os_conjugada: 'a OS conjugada'
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
      <div class="assist-nota">Consulta, prepara folhas e salva planejamentos quando você pede — pergunta antes se houver conflito. Não muda status nem contas.</div>`;
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
        if (acao.feita && acao.umaVez) return;
        acao.feita = true;
        a.classList.add('feita');
        Promise.resolve().then(() => acao.executar())
          .catch(err => mostrarErro('Não consegui: ' + (err && err.message ? err.message : err)));
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
    notasDoChat.forEach(t => partes.push(`<div class="assist-m assist-ele">${markdown(t)}</div>`));
    if (!partes.length) {
      partes.push(`<div class="assist-vazio">Olá! Eu consulto os dados do Gerador de OS, preparo folhas de OS, etiquetas, OE e ordens de compra para imprimir e monto planejamentos de produção. Exemplos:
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
