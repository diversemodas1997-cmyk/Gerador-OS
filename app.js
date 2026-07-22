/* ========================================================= */
/*                  ESTADO E PERSISTÊNCIA                    */
/* ========================================================= */
const SUPA_URL = 'https://ckkqrjkhorvaahyazqsr.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNra3Fyamtob3J2YWFoeWF6cXNyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4MTY2MjMsImV4cCI6MjA5MjM5MjYyM30.yT3Tb6KKx4sDNJXetwIoA77WudWUqQ2gCgT7JLi0iT8';
const supa = (window.supabase && typeof window.supabase.createClient === 'function')
  ? window.supabase.createClient(SUPA_URL, SUPA_KEY)
  : null;

let cloudCache = null;
let currentUser = null;
let currentRole = null; // 'admin' | 'usuario' | null
let saveTimer = null;
let inRecoveryFlow = false;
// Compras de materiais vindas do programa de Contabilidade (tabela própria
// compras_materiais no Supabase). O Gerador-OS só LÊ — entram como ENTRADAS
// no estoque. Não fazem parte do blob shared_data (fonte separada).
let comprasCache = [];
let comprasChannel = null;
// Catálogo de SKUs publicado pelo Estoque-Confeccao (tabela skus_catalogo).
// O Gerador-OS só LÊ — usado no dropdown de SKU dos cadastros de Desenho/Modelo.
let catalogoSkus = [];

async function cloudLoad() {
  if (!supa || !currentUser) return;
  const { data, error } = await supa
    .from('shared_data')
    .select('data')
    .eq('id', 'main')
    .maybeSingle();
  if (error) {
    console.error('cloudLoad', error);
    cloudCache = {};
    // Visibilidade do problema: a causa quase sempre e RLS bloqueando o
    // SELECT pra esse usuario. Sem o toast, falha silenciosa esconde
    // o motivo de "OS de outro computador nao aparecem".
    setTimeout(() => toast(
      `Falha ao ler dados compartilhados (${error.code || 'erro'}). ` +
      `Verifique as politicas RLS da tabela shared_data no Supabase.`,
      'err'
    ), 50);
    return;
  }
  cloudCache = (data && data.data) || {};
}

async function cloudFlush() {
  if (!supa || !currentUser || !cloudCache) return;
  // TRAVA ANTI-APAGAMENTO (causa raiz dos incidentes de perda de dados):
  // se estamos prestes a gravar um blob VAZIO (sem OS e sem desenhos),
  // isso é quase sempre um cloudCache zerado por uma leitura que falhou.
  // Antes de gravar vazio, confere o SERVIDOR: se ele ainda tem dados,
  // bloqueia o flush pra não sobrescrever o bom com vazio. Cobre inclusive
  // o caso da leitura ter falhado no carregamento (não dependemos de ter
  // visto dados nesta sessão). Ações intencionais liberam via _permitirFlushVazio.
  if (!_permitirFlushVazio && _blobEstaVazio(cloudCache)) {
    let servidorTemDados = false;
    try {
      const { data } = await supa.from('shared_data').select('data').eq('id', 'main').maybeSingle();
      const d = (data && data.data) || {};
      servidorTemDados = _contarItens(d, 'ordens') > 0 || _contarItens(d, 'desenhos') > 0;
    } catch (e) { console.warn('checagem anti-apagamento', e); }
    if (servidorTemDados) {
      console.error('cloudFlush BLOQUEADO: tentativa de gravar dados vazios sobre servidor com dados.');
      setSyncStatus('error');
      mostrarAlertaSalvamento('bloqueio',
        'A tela está sem dados (OS e desenhos), mas o servidor ainda tem seus dados. '
        + 'Para evitar um apagamento, a gravação foi bloqueada e NADA foi sobrescrito. '
        + 'Não continue editando — clique em "Recarregar agora" para trazer os dados de volta.');
      toast('⛔ Gravação bloqueada — nada foi sobrescrito. Recarregue a página.', 'err');
      return;
    }
  }
  setSyncStatus('saving');
  try {
    const { error } = await supa.from('shared_data').upsert({
      id: 'main',
      data: cloudCache,
      updated_at: new Date().toISOString(),
      updated_by: currentUser.id
    }, { onConflict: 'id' });
    if (error) throw error;
    setSyncStatus('ok');
    if (!_blobEstaVazio(cloudCache)) _appJaTeveDados = true;
    // Snapshot de contingência (local + pasta) do estado recém-salvo.
    salvarSnapshotContingencia();
    // Backup local automatico (silencioso; falha nao bloqueia o save).
    // Funcao definida mais abaixo, perto da pasta de PDFs.
    if (typeof escreverBackupJson === 'function') {
      escreverBackupJson().catch(e => console.warn('backup local', e));
    }
  } catch (e) {
    console.error('cloudFlush', e);
    setSyncStatus('error');
    mostrarAlertaSalvamento('erro',
      'Suas últimas alterações podem NÃO ter sido salvas no servidor (' + ((e && e.message) || 'erro de conexão') + '). '
      + 'Verifique a internet. Antes de recarregar, evite fechar a página para não perder o que digitou. '
      + 'Se o problema persistir, avise o suporte.');
  }
}

let realtimeChannel = null;

function iniciarRealtime() {
  if (!supa || !currentUser || realtimeChannel) return;
  realtimeChannel = supa
    .channel('shared_data_main')
    .on('postgres_changes',
      { event: 'UPDATE', schema: 'public', table: 'shared_data', filter: 'id=eq.main' },
      async (payload) => {
        if (!payload.new) return;
        if (payload.new.updated_by === currentUser.id) return;
        cloudCache = payload.new.data || {};
        await loadState();
        // Atualiza o marcador do polling pra evitar reload duplo
        if (payload.new.updated_at) lastSeenUpdatedAt = payload.new.updated_at;
        // Nao re-renderiza nova-os em edicao pra preservar o que o usuario
        // estava digitando. cloudCache ja foi atualizado — proxima
        // navegacao ja le valores frescos.
        // Detecta pagina ativa via .page:not(.hidden) (mais confiavel que
        // .nav-btn.active, que nao cobre 'print' — ela nao tem botao de menu).
        const ativa = document.querySelector('section.page:not(.hidden)');
        const pagina = ativa?.dataset?.page || 'home';
        if (pagina === 'print' && printOsAtual) {
          // OS pronta aberta: atualiza os checkboxes inline em vez de re-render
          // total — preserva scroll e nao pisca. Outros campos eventualmente
          // alterados ficam pra proxima visita.
          const fresh = STATE.ordens.find(x => x.id === printOsAtual.id);
          if (fresh) {
            printOsAtual = fresh;
            aplicarProgressoCheckboxes(fresh);
          }
        } else if (pagina !== 'nova-os') {
          goto(pagina);
        }
      })
    .subscribe();
  // Polling tambem e iniciado — se Realtime nao funcionar (publication
  // nao habilitada, rede bloqueia WebSocket, etc.), o polling cobre.
  iniciarPolling();
}

function pararRealtime() {
  if (supa && realtimeChannel) {
    supa.removeChannel(realtimeChannel);
    realtimeChannel = null;
  }
  if (supa && comprasChannel) {
    supa.removeChannel(comprasChannel);
    comprasChannel = null;
  }
  pararPolling();
}

// Lê as compras de materiais lançadas pela Contabilidade. Falha silenciosa se
// a tabela ainda não existe (integração não configurada) — o estoque segue
// funcionando só com entradas manuais + saídas de OS.
async function carregarComprasMateriais() {
  if (!supa || !currentUser) { comprasCache = []; return; }
  try {
    const { data, error } = await supa
      .from('compras_materiais')
      .select('*')
      .order('data', { ascending: false });
    if (error) { comprasCache = []; return; }
    comprasCache = Array.isArray(data) ? data : [];
  } catch (e) {
    comprasCache = [];
  }
}

// Lê o catálogo de SKUs (skus_catalogo, linha id='main') publicado pelo
// Estoque-Confeccao. Falha silenciosa se a tabela não existir ainda.
async function carregarCatalogoSkus() {
  if (!supa || !currentUser) { catalogoSkus = []; return; }
  try {
    const { data, error } = await supa
      .from('skus_catalogo').select('data').eq('id', 'main').maybeSingle();
    if (error || !data) { return; }
    let d = data.data || {};
    if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = {}; } }
    catalogoSkus = Array.isArray(d.skus) ? d.skus : [];
  } catch (e) { /* tabela ausente / sem permissão — ignora */ }
}

// Realtime das compras: quando a Contabilidade insere/atualiza uma compra,
// recarrega e re-renderiza o painel de estoque se ele estiver aberto.
function iniciarRealtimeCompras() {
  if (!supa || !currentUser || comprasChannel) return;
  comprasChannel = supa
    .channel('compras_materiais_all')
    .on('postgres_changes',
      { event: '*', schema: 'public', table: 'compras_materiais' },
      async () => {
        await carregarComprasMateriais();
        const ativa = document.querySelector('section.page:not(.hidden)');
        if ((ativa?.dataset?.page || '') === 'estoque') renderEstoque();
      })
    .subscribe();
}

// Polling fallback: a cada 15s consulta shared_data.updated_at. Se mudou
// desde a ultima vez vista (e nao foi este usuario que escreveu), recarrega.
// Garante sync mesmo se o canal Realtime falhar (rede instavel, publication
// nao habilitada, etc.). 15s e curto o suficiente pra parecer 'tempo real'
// sem pesar nas API calls.
let pollIntervalId = null;
let lastSeenUpdatedAt = null;

function iniciarPolling() {
  if (!supa || !currentUser || pollIntervalId) return;
  pollIntervalId = setInterval(async () => {
    if (!supa || !currentUser) return;
    try {
      const { data, error } = await supa.from('shared_data')
        .select('updated_at, updated_by, data')
        .eq('id', 'main')
        .maybeSingle();
      if (error || !data) return;
      // Inicializa o marcador na primeira leitura sem disparar reload
      if (lastSeenUpdatedAt === null) {
        lastSeenUpdatedAt = data.updated_at;
        return;
      }
      // Sem mudanca ou mudanca propria: ignora
      if (data.updated_at === lastSeenUpdatedAt) return;
      if (data.updated_by === currentUser.id) {
        lastSeenUpdatedAt = data.updated_at;
        return;
      }
      // Mudanca de outro usuario: aplica
      cloudCache = data.data || {};
      await loadState();
      // Mesma logica do realtime: print pronta atualiza so checkboxes;
      // demais paginas re-renderizam normalmente.
      const ativa = document.querySelector('section.page:not(.hidden)');
      const pagina = ativa?.dataset?.page || 'home';
      if (pagina === 'print' && printOsAtual) {
        const fresh = STATE.ordens.find(x => x.id === printOsAtual.id);
        if (fresh) {
          printOsAtual = fresh;
          aplicarProgressoCheckboxes(fresh);
        }
      } else if (pagina !== 'nova-os') {
        goto(pagina);
      }
      lastSeenUpdatedAt = data.updated_at;
      toast('Dados atualizados por outro usuário', 'ok');
    } catch (e) {
      console.warn('polling shared_data', e);
    }
  }, 15000);
}

function pararPolling() {
  if (pollIntervalId) clearInterval(pollIntervalId);
  pollIntervalId = null;
  lastSeenUpdatedAt = null;
}

function scheduleCloudSave() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { saveTimer = null; cloudFlush(); }, 800);
}

function setSyncStatus(status) {
  const el = document.getElementById('authSync');
  if (!el) return;
  el.classList.remove('saving', 'error');
  if (status === 'saving') { el.textContent = '☁ Salvando...'; el.classList.add('saving'); }
  else if (status === 'error') { el.textContent = '☁ Erro ao salvar'; el.classList.add('error'); }
  else { el.textContent = '☁ Sincronizado'; esconderAlertaSalvamento(); }
}

// Banner de aviso no topo do conteúdo. tipo 'bloqueio' (grave, trava
// anti-apagamento) ou 'erro' (falha de gravação por qualquer motivo).
function mostrarAlertaSalvamento(tipo, msg) {
  const box = document.getElementById('alertaSalvamento');
  if (!box) return;
  const ic = document.getElementById('alertaSalvamentoIcone');
  const tit = document.getElementById('alertaSalvamentoTitulo');
  const m = document.getElementById('alertaSalvamentoMsg');
  box.classList.remove('erro', 'bloqueio');
  box.classList.add(tipo === 'bloqueio' ? 'bloqueio' : 'erro');
  if (ic) ic.textContent = tipo === 'bloqueio' ? '⛔' : '⚠';
  if (tit) tit.textContent = tipo === 'bloqueio'
    ? 'Gravação bloqueada — seus dados no servidor estão protegidos'
    : 'Falha ao salvar no servidor';
  if (m) m.textContent = msg || '';
  box.classList.remove('hidden');
}

function esconderAlertaSalvamento() {
  const box = document.getElementById('alertaSalvamento');
  if (box) box.classList.add('hidden');
}

/* Snapshot diário: grava cópia do blob atual em shared_data_backups uma vez por dia. */
async function snapshotDiario() {
  if (!supa || !currentUser || !cloudCache) return;
  const hoje = new Date().toISOString().slice(0, 10);
  const { data: existente } = await supa
    .from('shared_data_backups')
    .select('id')
    .eq('snapshot_date', hoje)
    .maybeSingle();
  if (existente) return;
  const { error } = await supa.from('shared_data_backups').insert({
    snapshot_date: hoje,
    data: cloudCache,
    created_by: currentUser.id
  });
  if (error) { console.warn('snapshot falhou', error); return; }
  // Retenção: 30 dias
  const corte = new Date();
  corte.setDate(corte.getDate() - 30);
  const corteStr = corte.toISOString().slice(0, 10);
  await supa.from('shared_data_backups').delete().lt('snapshot_date', corteStr);
}

async function listarSnapshots() {
  if (!supa) return;
  const container = document.getElementById('snapshotsList');
  if (!container) return;
  container.innerHTML = '<div class="empty" style="padding:20px;">Carregando...</div>';
  const { data, error } = await supa
    .from('shared_data_backups')
    .select('id, snapshot_date, created_at')
    .order('snapshot_date', { ascending: false });
  if (error) { container.innerHTML = `<div class="empty" style="padding:20px;">Erro: ${esc(error.message)}</div>`; return; }
  if (!data || !data.length) { container.innerHTML = '<div class="empty" style="padding:20px;">Nenhum snapshot ainda — o primeiro é criado ao carregar o app.</div>'; return; }
  container.innerHTML = `<table class="table">
    <thead><tr><th>Data</th><th>Criado em</th><th class="col-actions">Ação</th></tr></thead>
    <tbody>${data.map(s => `
      <tr>
        <td><strong>${esc(s.snapshot_date)}</strong></td>
        <td>${esc(new Date(s.created_at).toLocaleString('pt-BR'))}</td>
        <td class="col-actions"><button class="btn small danger" onclick="restaurarSnapshot(${s.id}, '${esc(s.snapshot_date)}')">Restaurar</button></td>
      </tr>`).join('')}
    </tbody></table>`;
}

/* ========================================================= */
/*                        PRESENCE                           */
/* ========================================================= */
let presenceChannel = null;
let presenceOsId = null;

function iniciarPresenceOS(osKey) {
  if (!supa || !currentUser) return;
  if (presenceChannel && presenceOsId === osKey) return;
  pararPresenceOS();
  presenceOsId = osKey;
  const channelName = 'os_edit:' + osKey;
  presenceChannel = supa.channel(channelName, {
    config: { presence: { key: currentUser.id } }
  });
  presenceChannel
    .on('presence', { event: 'sync' }, () => renderizarPresence())
    .on('presence', { event: 'join' }, () => renderizarPresence())
    .on('presence', { event: 'leave' }, () => renderizarPresence())
    .subscribe(async (status) => {
      if (status === 'SUBSCRIBED') {
        await presenceChannel.track({
          email: currentUser.email || 'sem e-mail',
          at: Date.now()
        });
      }
    });
}

function pararPresenceOS() {
  if (supa && presenceChannel) {
    try { presenceChannel.untrack(); } catch (e) {}
    supa.removeChannel(presenceChannel);
  }
  presenceChannel = null;
  presenceOsId = null;
  const bar = document.getElementById('presenceBar');
  if (bar) bar.classList.add('hidden');
}

function renderizarPresence() {
  if (!presenceChannel) return;
  const state = presenceChannel.presenceState();
  const bar = document.getElementById('presenceBar');
  const usersEl = document.getElementById('presenceUsers');
  const countEl = document.getElementById('presenceCount');
  if (!bar || !usersEl || !countEl) return;
  const outros = [];
  for (const key in state) {
    if (key === (currentUser && currentUser.id)) continue;
    const meta = state[key][0];
    if (meta && meta.email) outros.push(meta.email);
  }
  if (!outros.length) { bar.classList.add('hidden'); return; }
  bar.classList.remove('hidden');
  countEl.textContent = outros.length;
  usersEl.innerHTML = outros.map(e => `<span class="user-chip">${esc(e)}</span>`).join(' ');
}

/* ========================================================= */
/*                  PAPÉIS / PERMISSÕES                      */
/* ========================================================= */
async function carregarPapel() {
  if (!supa || !currentUser) { currentRole = null; return; }
  const { data, error } = await supa
    .from('user_roles')
    .select('role')
    .eq('user_id', currentUser.id)
    .maybeSingle();
  if (error) { console.warn('carregarPapel', error); currentRole = 'usuario'; return; }
  currentRole = (data && data.role) || 'usuario';
}

function aplicarPermissoesUI() {
  const body = document.body;
  body.classList.remove('is-admin', 'is-usuario');
  if (currentRole === 'admin') body.classList.add('is-admin');
  else if (currentRole === 'usuario') body.classList.add('is-usuario');
}

function exigirAdmin(acao) {
  if (currentRole !== 'admin') {
    toast(`Apenas admin pode ${acao}`, 'err');
    return false;
  }
  return true;
}

async function setUserRole(novoPapel) {
  if (!supa) return;
  if (!exigirAdmin('gerenciar usuários')) return;
  const email = (document.getElementById('roleEmail').value || '').trim().toLowerCase();
  if (!email) { toast('Informe o e-mail do usuário', 'err'); return; }
  const { error } = await supa.rpc('set_user_role', { user_email: email, new_role: novoPapel });
  if (error) { toast('Erro: ' + error.message, 'err'); return; }
  toast(`${email} agora é ${novoPapel}`, 'ok');
  document.getElementById('roleEmail').value = '';
  listarUsuariosComPapel();
}

async function listarUsuariosComPapel() {
  if (!supa) return;
  const container = document.getElementById('usersList');
  if (!container) return;
  const { data, error } = await supa
    .from('user_roles')
    .select('user_id, role, created_at')
    .order('created_at', { ascending: true });
  if (error) { container.innerHTML = `<div class="empty" style="padding:20px;">Erro: ${esc(error.message)}</div>`; return; }
  if (!data || !data.length) { container.innerHTML = '<div class="empty" style="padding:20px;">Nenhum papel atribuído ainda.</div>'; return; }
  container.innerHTML = `<table class="table">
    <thead><tr><th>User ID</th><th>Papel</th><th>Desde</th></tr></thead>
    <tbody>${data.map(u => `
      <tr>
        <td style="font-family:'IBM Plex Mono',monospace; font-size:11px;">${esc(u.user_id)}</td>
        <td><span class="badge">${esc(u.role)}</span></td>
        <td>${esc(new Date(u.created_at).toLocaleDateString('pt-BR'))}</td>
      </tr>`).join('')}
    </tbody></table>`;
}

async function restaurarSnapshot(id, dataStr) {
  if (!exigirAdmin('restaurar snapshots')) return;
  if (!supa) return;
  const confirmTxt = prompt(
    `Restaurar o snapshot de ${dataStr}?\n\n` +
    `Isso vai SOBRESCREVER todos os cadastros e OS atuais com a versão daquele dia.\n\n` +
    `Para confirmar, digite RESTAURAR:`
  );
  if (confirmTxt === null) return;
  if ((confirmTxt || '').trim().toUpperCase() !== 'RESTAURAR') {
    toast('Palavra não conferiu — nada foi restaurado.', 'err');
    return;
  }
  const { data, error } = await supa
    .from('shared_data_backups')
    .select('data')
    .eq('id', id)
    .maybeSingle();
  if (error || !data) { toast('Snapshot não encontrado', 'err'); return; }
  cloudCache = data.data || {};
  const { error: upErr } = await supa.from('shared_data').upsert({
    id: 'main', data: cloudCache,
    updated_at: new Date().toISOString(),
    updated_by: currentUser.id
  }, { onConflict: 'id' });
  if (upErr) { toast('Erro ao gravar: ' + upErr.message, 'err'); return; }
  await loadState();
  goto('home');
  toast(`Snapshot de ${dataStr} restaurado`, 'ok');
}

const DB = {
  async get(key) {
    if (cloudCache) {
      const v = cloudCache[key];
      return (v !== undefined && v !== null) ? { key, value: v } : null;
    }
    try {
      const v = localStorage.getItem(key);
      return v !== null ? { key, value: v } : null;
    } catch (e) { return null; }
  },
  async set(key, value) {
    if (cloudCache) {
      cloudCache[key] = value;
      scheduleCloudSave();
      return { key, value };
    }
    try {
      localStorage.setItem(key, value);
      return { key, value };
    } catch (e) {
      console.error('localStorage cheio ou indisponível', e);
      return null;
    }
  },
  async delete(key) {
    if (cloudCache) {
      delete cloudCache[key];
      scheduleCloudSave();
      return { key, deleted: true };
    }
    localStorage.removeItem(key);
    return { key, deleted: true };
  }
};

/* ========================================================= */
/*                     AUTENTICAÇÃO                          */
/* ========================================================= */
const CAD_KEYS = ['tecidos','cores','materiais','modelos','colecoes','grades','desenhos','marcas','linhas','bases','blocos','equipe','funcoes','tarefas','etapas','componentes','ordens','estoqueMov','corteMov','costurandoMov','fiosMov','expedicaoMov','expedicaoJanelas','expedicaoCargas','expedicaoExcecoes','operacoes','osCounter','meta'];

async function inicializarAuth() {
  if (!supa) return;
  const { data: { session } } = await supa.auth.getSession();
  if (session && session.user) {
    currentUser = session.user;
    await cloudLoad();
    await carregarComprasMateriais();
    await carregarCatalogoSkus();
    await revalidarSkusDesenhos();
    iniciarRealtime();
    iniciarRealtimeCompras();
  }
  atualizarUIAuth();
  supa.auth.onAuthStateChange(async (event, session) => {
    if (event === 'PASSWORD_RECOVERY' && session) {
      inRecoveryFlow = true;
      currentUser = session.user;
      document.getElementById('modalAuth').classList.remove('hidden');
      trocarAbaAuth('reset_confirm');
      return;
    }
    if (event === 'SIGNED_IN' && session && !inRecoveryFlow) {
      currentUser = session.user;
      await cloudLoad();
      await carregarPapel();
      await carregarComprasMateriais();
      await carregarCatalogoSkus();
      await revalidarSkusDesenhos();
      iniciarRealtime();
      iniciarRealtimeCompras();
    } else if (event === 'SIGNED_OUT') {
      pararRealtime();
      pararPresenceOS();
      currentUser = null;
      cloudCache = null;
      currentRole = null;
      comprasCache = [];
      inRecoveryFlow = false;
    }
    aplicarPermissoesUI();
    atualizarUIAuth();
  });
}

function atualizarUIAuth() {
  const out = document.getElementById('authLoggedOut');
  const inn = document.getElementById('authLoggedIn');
  const emailEl = document.getElementById('authEmail');
  const appEl = document.querySelector('.app');
  const modal = document.getElementById('modalAuth');
  if (!out || !inn || !appEl || !modal) return;
  if (currentUser) {
    out.classList.add('hidden');
    inn.classList.remove('hidden');
    if (emailEl) emailEl.textContent = currentUser.email || currentUser.id;
    appEl.classList.remove('hidden');
    modal.classList.add('hidden');
  } else {
    out.classList.remove('hidden');
    inn.classList.add('hidden');
    appEl.classList.add('hidden');
    modal.classList.remove('hidden');
    const erroEl = document.getElementById('authErro');
    if (erroEl) erroEl.textContent = '';
  }
}

function abrirLogin() {
  document.getElementById('modalAuth').classList.remove('hidden');
  document.getElementById('authErro').textContent = '';
  document.getElementById('authEmailInput').focus();
}

function fecharLogin() {
  document.getElementById('modalAuth').classList.add('hidden');
}

function trocarAbaAuth(modo) {
  const modal = document.getElementById('modalAuth');
  const tabs = document.getElementById('authTabs');
  const emailGroup = document.getElementById('authEmailGroup');
  const senhaGroup = document.getElementById('authSenhaGroup');
  const senha2Group = document.getElementById('authSenha2Group');
  const actionBtn = document.getElementById('authActionBtn');
  const linkEsqueci = document.getElementById('linkEsqueci');
  const linkVoltar = document.getElementById('linkVoltar');
  const titulo = document.getElementById('authTitle');
  const sub = document.getElementById('authSub');
  modal.dataset.mode = modo;
  document.getElementById('authErro').textContent = '';
  document.querySelectorAll('.modal-auth .tab').forEach(t => t.classList.remove('active'));

  if (modo === 'login' || modo === 'signup') {
    tabs.classList.remove('hidden');
    emailGroup.classList.remove('hidden');
    senhaGroup.classList.remove('hidden');
    senha2Group.classList.add('hidden');
    linkEsqueci.classList.remove('hidden');
    linkVoltar.classList.add('hidden');
    titulo.textContent = 'Acesso restrito';
    sub.textContent = 'Faça login para usar o gerador. Seus cadastros ficam sincronizados na nuvem e acessíveis de qualquer computador.';
    document.getElementById(modo === 'login' ? 'tabLogin' : 'tabSignup').classList.add('active');
    actionBtn.textContent = modo === 'login' ? 'Entrar' : 'Criar conta';
    actionBtn.setAttribute('onclick', 'submeterAuth()');
  } else if (modo === 'reset_request') {
    tabs.classList.add('hidden');
    emailGroup.classList.remove('hidden');
    senhaGroup.classList.add('hidden');
    senha2Group.classList.add('hidden');
    linkEsqueci.classList.add('hidden');
    linkVoltar.classList.remove('hidden');
    titulo.textContent = 'Recuperar senha';
    sub.textContent = 'Informe seu e-mail. Vamos enviar um link para você criar uma nova senha.';
    actionBtn.textContent = 'Enviar link de recuperação';
    actionBtn.setAttribute('onclick', 'enviarEmailRecuperacao()');
  } else if (modo === 'reset_confirm') {
    tabs.classList.add('hidden');
    emailGroup.classList.add('hidden');
    senhaGroup.classList.remove('hidden');
    senha2Group.classList.remove('hidden');
    linkEsqueci.classList.add('hidden');
    linkVoltar.classList.add('hidden');
    titulo.textContent = 'Definir nova senha';
    sub.textContent = 'Digite e confirme sua nova senha. Ela precisa ter pelo menos 6 caracteres.';
    document.getElementById('authSenhaInput').value = '';
    document.getElementById('authSenha2Input').value = '';
    document.getElementById('authSenhaInput').setAttribute('autocomplete', 'new-password');
    document.getElementById('authSenhaInput').setAttribute('placeholder', 'nova senha');
    actionBtn.textContent = 'Atualizar senha';
    actionBtn.setAttribute('onclick', 'definirNovaSenha()');
  }
}

function abrirRecuperacaoSenha() {
  trocarAbaAuth('reset_request');
  document.getElementById('authEmailInput').focus();
}

async function enviarEmailRecuperacao() {
  if (!supa) { toast('Supabase não carregado', 'err'); return; }
  const email = document.getElementById('authEmailInput').value.trim();
  const erroEl = document.getElementById('authErro');
  const btn = document.getElementById('authActionBtn');
  if (!email) { erroEl.textContent = 'Informe seu e-mail'; return; }
  btn.disabled = true;
  erroEl.textContent = '';
  try {
    const { error } = await supa.auth.resetPasswordForEmail(email, {
      redirectTo: window.location.href.split('#')[0]
    });
    if (error) { erroEl.textContent = traduzirErroAuth(error.message); return; }
    toast('E-mail enviado. Verifique sua caixa de entrada (e spam).', 'ok');
    trocarAbaAuth('login');
  } finally {
    btn.disabled = false;
  }
}

async function definirNovaSenha() {
  if (!supa) { toast('Supabase não carregado', 'err'); return; }
  const senha = document.getElementById('authSenhaInput').value;
  const senha2 = document.getElementById('authSenha2Input').value;
  const erroEl = document.getElementById('authErro');
  const btn = document.getElementById('authActionBtn');
  if (!senha || !senha2) { erroEl.textContent = 'Preencha os dois campos'; return; }
  if (senha.length < 6) { erroEl.textContent = 'Senha precisa ter ao menos 6 caracteres'; return; }
  if (senha !== senha2) { erroEl.textContent = 'As senhas não conferem'; return; }
  btn.disabled = true;
  erroEl.textContent = '';
  try {
    const { data, error } = await supa.auth.updateUser({ password: senha });
    if (error) { erroEl.textContent = traduzirErroAuth(error.message); return; }
    inRecoveryFlow = false;
    currentUser = data.user;
    await cloudLoad();
    await carregarPapel();
    aplicarPermissoesUI();
    iniciarRealtime();
    fecharLogin();
    await loadState();
    atualizarUIAuth();
    goto('home');
    toast('Senha atualizada. Bem-vindo(a)!', 'ok');
  } catch (e) {
    erroEl.textContent = e.message || 'Erro inesperado';
  } finally {
    btn.disabled = false;
  }
}

function traduzirErroAuth(msg) {
  if (/invalid login credentials/i.test(msg)) return 'E-mail ou senha incorretos';
  if (/user already registered/i.test(msg)) return 'Este e-mail já está cadastrado — use Entrar';
  if (/password should be at least/i.test(msg)) return 'Senha muito curta';
  if (/rate limit/i.test(msg)) return 'Muitas tentativas — aguarde um minuto';
  if (/email.*invalid/i.test(msg)) return 'E-mail inválido';
  if (/signup.*disabled|signups are disabled/i.test(msg)) return 'Cadastro público desabilitado. Peça acesso ao administrador.';
  if (/for security purposes.*seconds/i.test(msg)) return 'Muitas tentativas. Aguarde alguns segundos.';
  if (/user not found/i.test(msg)) return 'E-mail não cadastrado';
  if (/new password should be different/i.test(msg)) return 'A nova senha precisa ser diferente da atual';
  return msg;
}

async function submeterAuth() {
  if (!supa) { toast('Supabase não carregado — verifique conexão', 'err'); return; }
  const modo = document.getElementById('modalAuth').dataset.mode || 'login';
  const email = document.getElementById('authEmailInput').value.trim();
  const senha = document.getElementById('authSenhaInput').value;
  const erroEl = document.getElementById('authErro');
  const btn = document.getElementById('authActionBtn');
  if (!email || !senha) { erroEl.textContent = 'Informe e-mail e senha'; return; }
  if (senha.length < 6) { erroEl.textContent = 'Senha precisa ter ao menos 6 caracteres'; return; }
  btn.disabled = true;
  erroEl.textContent = '';
  try {
    const resp = modo === 'signup'
      ? await supa.auth.signUp({ email, password: senha })
      : await supa.auth.signInWithPassword({ email, password: senha });
    if (resp.error) { erroEl.textContent = traduzirErroAuth(resp.error.message); return; }
    if (!resp.data || !resp.data.session) {
      erroEl.textContent = 'Conta criada — confirme seu e-mail ou tente entrar';
      return;
    }
    currentUser = resp.data.session.user;
    await cloudLoad();
    await carregarPapel();
    aplicarPermissoesUI();
    iniciarRealtime();
    const temLocal = CAD_KEYS.some(k => localStorage.getItem(k) !== null);
    const cloudVazio = !cloudCache || Object.keys(cloudCache).length === 0;
    if (temLocal && cloudVazio) {
      if (confirm('Você tem cadastros salvos neste navegador. Enviar para a nuvem agora? (ficarão visíveis pra toda equipe)')) {
        await migrarLocalParaNuvem();
      }
    }
    fecharLogin();
    atualizarUIAuth();
    await loadState();
    await migrarEtapasOS();        // padroniza etapas das OSs (1×, admin)
    // Publica o snapshot de estoque p/ a Contabilidade ao entrar (só admin
    // escreve no blob). Garante que exista mesmo sem nenhuma edição na sessão.
    if (currentRole === 'admin') atualizarContabSnapshot();
    goto('home');
    toast('Conectado — cadastros sincronizados na nuvem', 'ok');
  } catch (e) {
    erroEl.textContent = e.message || 'Erro inesperado';
  } finally {
    btn.disabled = false;
  }
}

async function migrarLocalParaNuvem() {
  let n = 0;
  for (const k of CAD_KEYS) {
    const v = localStorage.getItem(k);
    if (v !== null) { cloudCache[k] = v; n++; }
  }
  if (n > 0) { await cloudFlush(); toast(`${n} item(ns) enviado(s) para a nuvem`, 'ok'); }
}

async function sairConta() {
  if (!supa) return;
  if (!confirm('Sair da conta? Seus dados continuam salvos na nuvem.')) return;
  if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; await cloudFlush(); }
  pararRealtime();
  pararPresenceOS();
  await supa.auth.signOut();
  currentUser = null;
  cloudCache = null;
  currentRole = null;
  CAD_KEYS.forEach(k => { if (Array.isArray(STATE[k])) STATE[k] = []; });
  STATE.osCounter = 0;
  aplicarPermissoesUI();
  atualizarUIAuth();
  toast('Desconectado.', 'ok');
}

const STATE = {
  tecidos: [],
  cores: [],
  materiais: [],
  modelos: [],
  colecoes: [],
  grades: [],
  desenhos: [],
  marcas: [],
  linhas: [],
  bases: [],
  blocos: [],
  equipe: [],
  funcoes: [],
  tarefas: [],
  etapas: [],
  componentes: [],
  ordens: [],
  // Movimentações de estoque de tecidos (entradas manuais e saídas automáticas
  // por OS). Cada item: { id, tipo:'entrada'|'saida', tecidoNome, corNome, kg,
  // data, origem:'manual'|'os', osId, osNumero, obs }.
  estoqueMov: [],
  // Movimentações MANUAIS do estoque de corte (contagem física e ajustes).
  // As entradas (OS) e saídas (etapa Costura marcada) são DERIVADAS das OS em
  // tempo de render — não persistem aqui. Cada item manual:
  // { id, tipo:'entrada'|'saida', tecidoNome, corNome, qtd, data, obs }.
  corteMov: [],
  // Contagem manual das fases seguintes do fluxo (mesmo formato de corteMov):
  // Costurando (Costura), Retirada de fios e Expedição.
  costurandoMov: [],
  fiosMov: [],
  expedicaoMov: [],
  // ---------- Planejamento de expedição ----------
  // Janelas = quando a expedição acontece, cadastradas pelo usuário. Duas
  // naturezas: 'semanal' repete nos diasSemana pra sempre; 'data' acontece
  // uma vez só. Toda expedição é interna, ida e volta entre as 2 unidades —
  // daí duas horas por janela.
  // { id, nome, tipo:'semanal'|'data', diasSemana:[0..6], data, horaIda,
  //   horaVolta, volMin, volMax, ativo, obs }
  // volMin/volMax em '' herdam o padrão de meta.expedicao.
  expedicaoJanelas: [],
  // OS alocada numa ocorrência (janela + data de origem) e perna do trajeto.
  // A data guardada é sempre a ORIGINAL da ocorrência, não a remarcada: assim
  // remarcar leva a carga junto em vez de órfã-la.
  // { id, janelaId, data, perna:'ida'|'volta', osId, volumes, obs }
  expedicaoCargas: [],
  // Ocorrência cancelada ou remarcada pontualmente (só janelas semanais
  // precisam disso — uma janela de data avulsa se edita direto).
  // { id, janelaId, data, tipo:'cancelada'|'remarcada', novaData, horaIda,
  //   horaVolta, motivo }
  expedicaoExcecoes: [],
  // ---------- Planejamento diário de operações ----------
  // A jornada planejada de cada POSTO (função cadastrada) em cada dia. Uma
  // operação aqui é o processo completo do posto — início + duração total, com
  // todas as etapas internas subentendidas — e não uma tarefa por OS. Diferente
  // do checklist da OS, que registra o que já aconteceu, aqui é o que vai
  // acontecer e em que horário.
  // { id, data:'YYYY-MM-DD', funcaoId, funcaoNome, operacao,
  //   escopo:'completa'|'etapa', etapa, inicio:'HH:MM', duracaoMin,
  //   responsavelId, responsavelNome, referencia, ordem,
  //   prioridade:'urgente'|'emergente'|'eletiva',
  //   status:'pendente'|'andamento'|'feita', obs }
  // ordem = posição manual dentro do dia (gravada ao mover); sem ela a operação
  // se ordena pelo horário de início.
  // escopo 'completa' (padrão) = todas as etapas da função embutidas;
  // 'etapa' = o posto foi planejado por partes e esta linha é só a etapa nomeada.
  // funcaoNome/responsavelNome são cópias de exibição: se o cadastro for
  // renomeado ou excluído, o histórico do dia continua legível. referencia é
  // texto livre (lote, coleção, OSs) — o plano não fica preso a um pedido.
  operacoes: [],
  osCounter: 0,
  // Flags/metadados internos persistidos (ex.: migrações já executadas).
  meta: {},
  // Overrides de rótulo das pastas/subpastas (fixas ou customizadas). A chave
  // técnica (ex.: 'camiseta', 'basica') segue inalterada nas grades — só o
  // texto exibido muda. tpOrder/vrOrder definem ordem manual; chaves ausentes
  // caem no fim, com fixas antes das customizadas alfabéticas.
  gradeFolderLabels: { tp: {}, vr: {}, tpOrder: [], vrOrder: [] },
  etapasPadrao: ['Corte', 'Acabamento de mangas', 'Costura', 'Retirada de fios', 'Estampa', 'Lavanderia', 'Ensaque', 'Expedição'],
  componentesPadrao: ['Frente', 'Costas', 'Capuz', 'Forro do capuz', 'Mangas', 'Bolso canguru', 'Punho', 'Barra', 'Ribana', 'Cobre gola', 'Recorte lateral', 'Cordão', 'Ilhós', 'Etiqueta interna', 'Tag']
};

// Chaves cujo conteúdo afeta o snapshot de estoque lido pela Contabilidade /
// Estoque-Confeccao. Inclui desenhos/modelos/cores porque o SKU de cada OS é
// resolvido a partir deles (skusDaOS) — editar o SKU de um desenho precisa
// republicar o snapshot, senão o Estoque continua com o SKU antigo.
const _CHAVES_CONTAB_SNAPSHOT = ['ordens', 'estoqueMov', 'corteMov', 'desenhos', 'modelos', 'cores'];

async function saveState(key) {
  try {
    await DB.set(key, JSON.stringify(STATE[key]));
    // Republica o snapshot de estoque p/ a Contabilidade quando muda algo
    // que altera os saldos. Best-effort; não bloqueia nem quebra o save.
    if (_CHAVES_CONTAB_SNAPSHOT.includes(key) && typeof construirContabSnapshot === 'function') {
      atualizarContabSnapshot();
    }
  } catch (e) {
    console.error('Erro ao salvar', key, e);
    toast('Erro ao salvar no armazenamento', 'err');
  }
}

// Normaliza nome de função pra comparar (remove acentos, baixa caixa, colapsa
// qualquer pontuação/espaço múltiplo) — assim "Produção", "producao",
// "Enfestadeira / Esteira" etc. caem todos no mesmo canônico.
function _normFuncaoNome(s) {
  return (s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}
const _FUNCAO_COORD_ENFEST_NORM = _normFuncaoNome('Coordenador de produção Enfestadeira/Esteira de corte');
function ehFuncaoCoordEnfestEsteira(nome) {
  return _normFuncaoNome(nome) === _FUNCAO_COORD_ENFEST_NORM;
}

async function loadState() {
  const keys = ['tecidos','cores','materiais','modelos','colecoes','grades','desenhos','marcas','linhas','bases','blocos','equipe','funcoes','tarefas','etapas','componentes','ordens','estoqueMov','corteMov','costurandoMov','fiosMov','expedicaoMov','expedicaoJanelas','expedicaoCargas','expedicaoExcecoes','operacoes','meta'];
  for (const k of keys) {
    try {
      const r = await DB.get(k);
      if (r && r.value) {
        try { STATE[k] = JSON.parse(r.value); } catch { STATE[k] = []; }
      }
    } catch (e) { /* chave não existe ainda, ok */ }
  }
  // Marca que o app já viu dados reais — arma a trava anti-apagamento e
  // libera o snapshot de contingência (que só grava blobs não-vazios).
  if ((STATE.ordens && STATE.ordens.length) || (STATE.desenhos && STATE.desenhos.length)) {
    _appJaTeveDados = true;
  }
  // Carrega overrides de rótulos das pastas de grades (objeto, não array)
  try {
    const r = await DB.get('gradeFolderLabels');
    if (r && r.value) {
      try {
        const parsed = JSON.parse(r.value);
        STATE.gradeFolderLabels = {
          tp: parsed?.tp || {},
          vr: parsed?.vr || {},
          tpOrder: Array.isArray(parsed?.tpOrder) ? parsed.tpOrder : [],
          vrOrder: Array.isArray(parsed?.vrOrder) ? parsed.vrOrder : []
        };
      } catch { STATE.gradeFolderLabels = { tp: {}, vr: {}, tpOrder: [], vrOrder: [] }; }
    }
  } catch (e) { /* ok */ }
  // Carrega o contador de OS (não é array, é número)
  try {
    const c = await DB.get('osCounter');
    if (c && c.value) STATE.osCounter = parseInt(c.value) || 0;
  } catch (e) { /* ok */ }
  // Seed default etapas se vazio (primeira execução)
  if (!STATE.etapas || !STATE.etapas.length) {
    STATE.etapas = (STATE.etapasPadrao || []).map((nome, i) => ({
      id: 'id_' + Date.now() + '_' + i,
      nome,
      ordem: (i + 1) * 10,
      funcoesIds: []
    }));
    if (STATE.etapas.length) { try { await saveState('etapas'); } catch (e) {} }
  }
  // Seed default componentes se vazio (primeira execução)
  if (!STATE.componentes || !STATE.componentes.length) {
    STATE.componentes = (STATE.componentesPadrao || []).map((nome, i) => ({
      id: 'id_' + (Date.now() + 1000) + '_' + i,
      nome,
      desc: ''
    }));
    if (STATE.componentes.length) { try { await saveState('componentes'); } catch (e) {} }
  }
  // Migração: operações planejadas no formato antigo (uma linha por OS, com
  // peças e sem duração) viram jornada de posto.
  try {
    if (migrarOperacoesParaJornada()) await saveState('operacoes');
  } catch (e) { console.warn('migrarOperacoesParaJornada', e); }
  // Migração: remove em definitivo a função "Coordenador de produção
  // Enfestadeira/Esteira de corte" (decisão do admin).
  if (Array.isArray(STATE.funcoes)) {
    const antes = STATE.funcoes.length;
    STATE.funcoes = STATE.funcoes.filter(f => !ehFuncaoCoordEnfestEsteira(f?.nome));
    if (STATE.funcoes.length !== antes) {
      try { await saveState('funcoes'); } catch (e) {}
    }
  }

  // Limpa equipe SEMPRE — mesmo após STATE.funcoes já ter sido purgada numa
  // execução anterior, p.funcao pode estar stale (ex.: alguém salvou um membro
  // escolhendo a opção "(não cadastrada)" no dropdown depois da migração ter
  // rodado). Sem isso, o nome volta a aparecer como opção no select de função.
  if (Array.isArray(STATE.equipe)) {
    let limpou = 0;
    STATE.equipe.forEach(p => {
      if (ehFuncaoCoordEnfestEsteira(p.funcao)) {
        p.funcao = '';
        p.funcaoId = '';
        limpou++;
      }
    });
    if (limpou) { try { await saveState('equipe'); } catch (e) {} }
  }

  // Sincronização: garante que equipe.funcao reflete o nome atual da função vinculada
  if (Array.isArray(STATE.equipe) && Array.isArray(STATE.funcoes)) {
    let mudou = 0;
    STATE.equipe.forEach(p => {
      // Caso 1: pessoa tem funcaoId → sincroniza p.funcao com o nome atual da função
      if (p.funcaoId) {
        const f = STATE.funcoes.find(x => x.id === p.funcaoId);
        if (f && f.nome && f.nome !== p.funcao) {
          p.funcao = f.nome;
          mudou++;
        }
        return;
      }
      // Caso 2: pessoa sem funcaoId mas com p.funcao → tenta vincular
      if (p.funcao) {
        const f = STATE.funcoes.find(x => (x.nome || '').trim().toLowerCase() === (p.funcao || '').trim().toLowerCase());
        if (f) {
          p.funcaoId = f.id;
          p.funcao = f.nome;
          mudou++;
        }
      }
    });
    if (mudou > 0) { try { await saveState('equipe'); } catch (e) {} }
  }

  // Auto-preenche a "Linha de SKU" dos modelos cujo SKU é dedutível pelo nome.
  // Roda só p/ admin, só quando o campo está VAZIO (nunca sobrescreve edição
  // manual) — então roda no máximo uma vez por modelo. Camiseta Bicolor fica de
  // fora (não há linha clara no catálogo de SKUs).
  if (currentRole === 'admin' && Array.isArray(STATE.modelos)) {
    const padroes = [
      { re: /blusa\s+moletom\s+tricolor/, linha: 'BM.TRI' },
      { re: /blusa\s+moletom\s+basica/,   linha: 'BM.LISA' },
      { re: /camiseta\s+tricolor/,        linha: 'CM.TRI.LISA' },
      { re: /camiseta\s+polo/,            linha: 'PM.LISA' },
      { re: /camiseta\s+basica/,          linha: 'CM.LISA' },
    ];
    let mudou = 0;
    STATE.modelos.forEach(m => {
      if (m.skuLinha) return;
      const nome = _normNome(m.nome || '');
      const p = padroes.find(x => x.re.test(nome));
      if (p) { m.skuLinha = p.linha; mudou++; }
    });
    if (mudou) { try { await saveState('modelos'); } catch (e) {} }
  }

  // Auto-preenche a "Sigla SKU" das cores cuja sigla é inequívoca no catálogo de
  // SKUs. Fora: Grafite (catálogo tem GRA e GRAF — ambíguo) e Off-White (sem SKU).
  // Admin, só quando vazio (não sobrescreve).
  if (currentRole === 'admin' && Array.isArray(STATE.cores)) {
    const siglas = {
      'preto': 'PRE', 'branco': 'BRA', 'verde': 'VERDE', 'vermelho': 'VERM',
      'azul': 'AZUL', 'bege': 'BEGE', 'roxo': 'ROXO', 'marrom': 'MARROM',
      'caqui': 'CAQUI', 'mostarda': 'MOSTARDA',
    };
    let mudou = 0;
    STATE.cores.forEach(c => {
      if (c.siglaSku) return;
      // corBaseNome tira o tecido do fim ("Preto Moletom" → "preto"): sem isto as
      // cores no formato novo nunca casariam no mapa e ficariam sem sigla, e aí
      // revalidarSkusDesenhos() não deduziria o SKU dos desenhos.
      const s = siglas[corBaseNome(c.nome)];
      if (s) { c.siglaSku = s; mudou++; }
    });
    if (mudou) { try { await saveState('cores'); } catch (e) {} }
  }
  // O auto-preenchimento do SKU dos DESENHOS roda em revalidarSkusDesenhos(),
  // após o catálogo (skus_catalogo) carregar — para VALIDAR contra a relação de
  // SKUs (só usa SKU que existe; nunca inventa SKU fora do catálogo).
  // Republica o snapshot p/ a Contabilidade/Estoque-Confeccao SEMPRE que o admin
  // carrega o estado (login OU reload). Garante SKUs frescos mesmo sem nenhuma
  // edição na sessão — sem isso, um reload deixava o snapshot antigo no ar.
  if (currentRole === 'admin' && typeof atualizarContabSnapshot === 'function') {
    atualizarContabSnapshot();
  }
}

// Preenche o SKU dos desenhos técnicos VAZIOS, validando contra o catálogo de
// SKUs (regra: só SKUs que constam na relação de referência). Roda após
// carregarCatalogoSkus. NUNCA toca em desenho que já tem SKU — mapeamentos
// manuais (ex.: moletom tricolor preto → BM.TRI-BEGE) ficam intactos. Se o SKU
// deduzido (linha do modelo + sigla da cor) não existir no catálogo, deixa em
// branco para escolha manual no dropdown. Guard por catálogo carregado.
async function revalidarSkusDesenhos() {
  if (currentRole !== 'admin' || !Array.isArray(STATE.desenhos)) return;
  if (!Array.isArray(catalogoSkus) || !catalogoSkus.length) return;
  const validos = new Set(catalogoSkus.map(s => (s.item || '').trim().toUpperCase()).filter(Boolean));
  let mudou = 0;
  STATE.desenhos.forEach(d => {
    if (d.skuLinha) return;                        // já definido (manual/anterior) → não mexe
    const m = (STATE.modelos || []).find(x => x.id === d.modeloId);
    const linha = ((m && m.skuLinha) || '').trim().toUpperCase();
    const c = (STATE.cores || []).find(x => x.id === d.corPrincipalId);
    const sigla = ((c && c.siglaSku) || '').trim().toUpperCase();
    const ded = (linha && sigla) ? (linha + '-' + sigla) : '';
    if (ded && validos.has(ded)) { d.skuLinha = ded; mudou++; }  // só preenche se EXISTIR no catálogo
  });
  if (mudou) { try { await saveState('desenhos'); } catch (e) {} }
}

// Template de etapas "mais atual" por tipo de produto (confirmado pelo Junior).
// Camiseta usa Acabamento de mangas; Blusa Moletom usa Fechamento de punhos/barra.
// A etapa terminal "Estoque" dispara a entrada de produtos acabados.
// Peças-alvo por tamanho que toda OS NOVA já traz preenchido, junto do número
// sequencial e da data. É o padrão da casa; o campo segue editável, e editar uma
// OS existente carrega o valor salvo dela. Mudou o padrão? Troca aqui.
const PECAS_ALVO_PADRAO = 160;

const ETAPAS_TEMPLATE_OS = {
  camiseta: ['Corte', 'Acabamento de mangas', 'Ensaque', 'Expedição', 'Costura', 'Retirada de fios', 'Estoque'],
  moletom:  ['Corte', 'Fechamento de punhos', 'Fechamento de barra', 'Ensaque', 'Expedição', 'Costura', 'Retirada de fios', 'Estoque'],
};

// Migração ÚNICA (admin): padroniza as etapas das OSs pelo template do tipo
// (camiseta x blusa moletom), só copiando a lista. Preserva os checks/seq das
// etapas que continuam; descarta os das que saíram. Roda 1× (flag em STATE.meta).
async function migrarEtapasOS() {
  if (currentRole !== 'admin' || !Array.isArray(STATE.ordens)) return;
  STATE.meta = STATE.meta || {};
  if (STATE.meta.etapasPadronizadasV1) return;        // já rodou — não mexe mais
  let mudou = 0;
  STATE.ordens.forEach(o => {
    const tipo = /moletom|blusa/i.test(o.modeloNome || '') ? 'moletom' : 'camiseta';
    const tmpl = ETAPAS_TEMPLATE_OS[tipo];
    const atual = o.etapas || [];
    const igual = atual.length === tmpl.length && atual.every((n, i) => n === tmpl[i]);
    if (igual) return;
    o.etapas = tmpl.slice();
    // Preserva check/seq só das etapas que continuam no template.
    if (o.progresso) {
      const manter = new Set(tmpl);
      ['etapasCheck', 'etapasSeq'].forEach(k => {
        const obj = o.progresso[k];
        if (obj) Object.keys(obj).forEach(nome => { if (!manter.has(nome)) delete obj[nome]; });
      });
    }
    mudou++;
  });
  STATE.meta.etapasPadronizadasV1 = true;
  try {
    if (mudou) await saveState('ordens');
    await saveState('meta');
    if (mudou) toast(`Etapas padronizadas em ${mudou} OS`, 'ok');
  } catch (e) { console.warn('migrarEtapasOS', e); }
}

function uid() { return 'id_' + Date.now() + '_' + Math.floor(Math.random()*1000); }

/* ========================================================= */
/*                      NAVEGAÇÃO                            */
/* ========================================================= */
// Preservacao de scroll por pagina: ao trocar de pagina, guarda o scrollY
// atual no sessionStorage e restaura ao voltar. Isso evita o reset ao topo
// que aborrecia ao navegar OS -> editar -> voltar.
function _scrollKey(page) { return 'gos:scroll:' + page; }

function _salvarScrollPaginaAtual() {
  try {
    const atual = document.querySelector('section.page:not(.hidden)');
    if (atual && atual.dataset && atual.dataset.page) {
      sessionStorage.setItem(_scrollKey(atual.dataset.page), String(window.scrollY || 0));
    }
  } catch (e) { /* sessionStorage pode estar indisponivel — segue sem ele */ }
}

function _restaurarScrollPagina(page) {
  let y = 0;
  try { y = parseFloat(sessionStorage.getItem(_scrollKey(page)) || '0') || 0; }
  catch (e) { y = 0; }
  // rAF para esperar o layout estabilizar (sections viraram hidden/visible).
  requestAnimationFrame(() => window.scrollTo(0, y));
}

// Drawer mobile: no celular o menu lateral vira overlay full-screen. Aberto =
// usuario ve so o menu; fechado = usuario ve so a pagina + um botao "Menu" no
// topo. Ao escolher uma opcao do menu, fecha automaticamente.
function abrirMenuMobile() {
  document.body.classList.add('mobile-menu-open');
}
function fecharMenuMobile() {
  document.body.classList.remove('mobile-menu-open');
}
window.abrirMenuMobile = abrirMenuMobile;
window.fecharMenuMobile = fecharMenuMobile;

function goto(page) {
  // Bloqueia navegação a páginas de cadastro para usuários não-admin
  if (page && page.startsWith('cad-') && currentRole && currentRole !== 'admin') {
    toast('Apenas admin pode acessar cadastros', 'err');
    page = 'home';
  }
  _salvarScrollPaginaAtual();  // guarda onde o usuario estava na pagina anterior
  document.querySelectorAll('section.page').forEach(s => s.classList.add('hidden'));
  const target = document.querySelector(`section.page[data-page="${page}"]`);
  if (target) target.classList.remove('hidden');
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const btn = document.querySelector(`.nav-btn[data-page="${page}"]`);
  if (btn) btn.classList.add('active');
  _restaurarScrollPagina(page);  // restaura o scroll salvo desta pagina (ou 0 na 1a visita)
  fecharMenuMobile();  // mobile: fecha o overlay do menu quando entra na pagina

  // Presence: só ativa quando está editando OS
  if (page !== 'nova-os') pararPresenceOS();

  // renderiza listas
  if (page === 'home') renderHome();
  if (page === 'cad-tecidos') renderTecidos();
  if (page === 'cad-cores') renderCores();
  if (page === 'cad-materiais') renderMateriais();
  if (page === 'cad-modelos') renderModelos();
  if (page === 'cad-colecoes') renderColecoes();
  if (page === 'cad-grades') renderGrades();
  if (page === 'cad-desenhos') renderDesenhos();
  if (page === 'cad-marcas') renderMarcas();
  if (page === 'cad-linhas') renderLinhas();
  if (page === 'cad-bases') renderBases();
  if (page === 'cad-blocos') renderBlocos();
  if (page === 'cad-equipe') renderEquipe();
  if (page === 'cad-funcoes') renderFuncoes();
  if (page === 'cad-etapas') renderEtapasCad();
  if (page === 'cad-componentes') renderComponentesCad();
  if (page === 'lista-os') renderListaOS();
  if (page === 'estoque') renderEstoque();
  if (page === 'corte') renderFasePainel(0);
  if (page === 'costurando') renderFasePainel(1);
  if (page === 'fios') renderFasePainel(2);
  if (page === 'expedicao') { renderFasePainel(3); trocarAbaExpedicao(expAbaAtiva); }
  if (page === 'operacoes') renderOperacoes();
  if (page === 'print-expedicao') {
    renderPrintPlanoExpedicao();
    // Auto-save da OE (folha do plano) na pasta conectada — mesma ideia do
    // PDF das OS. Silencioso e sem pasta conectada não faz nada.
    salvarPdfOeNaPasta({ silent: true }).catch(e => console.warn('auto-save OE', e));
  }
  if (page === 'nova-os') initOSForm();
  if (page === 'config') {
    atualizarPdfFolderStatus();
    atualizarBackupFolderStatus();
    atualizarOeFolderStatus();
  }
}

document.querySelectorAll('.nav-btn').forEach(b => b.addEventListener('click', () => goto(b.dataset.page)));

// Injeta o botao "≡ Menu" no topo de cada .page-header. Visivel apenas no
// mobile via CSS — desktop nunca o ve. Idempotente: se rodar de novo, nao
// duplica (checa pela classe).
(function injetarBotaoMenuMobile() {
  document.querySelectorAll('section.page .page-header').forEach(header => {
    if (header.querySelector('.btn-menu-mobile')) return;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn-menu-mobile';
    btn.setAttribute('aria-label', 'Abrir menu');
    btn.innerHTML = '<span aria-hidden="true">≡</span> Menu';
    btn.addEventListener('click', abrirMenuMobile);
    header.insertBefore(btn, header.firstChild);
  });
})();

// Recolhe/expande grupos do menu lateral. Estado persiste em localStorage.
function toggleNavGroup(labelEl) {
  const group = labelEl?.closest?.('.nav-group');
  if (!group) return;
  const key = group.dataset.group;
  group.classList.toggle('collapsed');
  if (key) {
    try {
      const colapsadas = JSON.parse(localStorage.getItem('navGroupsCollapsed') || '{}');
      colapsadas[key] = group.classList.contains('collapsed');
      localStorage.setItem('navGroupsCollapsed', JSON.stringify(colapsadas));
    } catch (e) { /* ignora */ }
  }
}
window.toggleNavGroup = toggleNavGroup;

// Restaura estado dos grupos ao carregar
(function restaurarNavGroups() {
  try {
    const colapsadas = JSON.parse(localStorage.getItem('navGroupsCollapsed') || '{}');
    Object.keys(colapsadas).forEach(key => {
      if (!colapsadas[key]) return;
      const g = document.querySelector(`.nav-group[data-group="${key}"]`);
      if (g) g.classList.add('collapsed');
    });
  } catch (e) { /* ignora */ }
})();

/* ========================================================= */
/*                      TOAST                                */
/* ========================================================= */
function toast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  // 'no-print' precisa ser reescrito aqui: esta linha SUBSTITUI o className, então
  // a classe posta no index.html se perderia no primeiro aviso. Sem ela o toast
  // (position:fixed) entra na foto do html2canvas — salvar uma OS dispara várias
  // capturas em sequência e o aviso de um passo caía sobreposto à folha seguinte.
  t.className = 'toast no-print show ' + type;
  setTimeout(() => t.classList.remove('show'), 2400);
}

/* ========================================================= */
/*                    MODAL DE CADASTRO                      */
/* ========================================================= */
let cadastroContext = null;

function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

function openCadastroModal(tipo, editId = null, origin = null) {
  cadastroContext = { tipo, editId, origin };
  const title = document.getElementById('modal-cad-title');
  const box = document.getElementById('modal-cad-fields');

  const titles = {
    tecido: 'Tecido', cor: 'Cor', material: 'Material / Aviamento',
    modelo: 'Modelo', colecao: 'Coleção', grade: 'Grade de tamanhos',
    desenho: 'Desenho técnico',
    marca: 'Marca / Griffe', linha: 'Linha', base: 'Base', bloco: 'Bloco / Revisão',
    equipe: 'Membro da equipe', funcao: 'Função', tarefa: 'Tarefa', etapa: 'Etapa de produção', componente: 'Componente'
  };
  title.textContent = (editId ? 'Editar ' : 'Novo ') + titles[tipo];

  let item = {};
  if (editId) {
    const list = pluralize(tipo);
    item = STATE[list].find(x => x.id === editId) || {};
  }

  if (tipo === 'tecido') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Moletom Bulk"></div>
        <div class="field"><label>Categoria (define limite de enfesto e multiplicador)</label>
          <select id="m-categoria">
            <option value="">— selecione —</option>
            <option value="malha" ${item.categoria==='malha'?'selected':''}>Malha algodão (limite 80 camadas)</option>
            <option value="moletom" ${item.categoria==='moletom'?'selected':''}>Moletom (limite 36 camadas)</option>
            <option value="ribana" ${item.categoria==='ribana'?'selected':''}>Ribana (1 camada = 2 peças)</option>
            <option value="outro" ${item.categoria==='outro'?'selected':''}>Outro (sem limite)</option>
          </select>
        </div>
        <div class="field"><label>Peso / gramatura padrão (g/m²)</label><input type="number" min="0" step="1" id="m-peso" value="${esc(item.peso||'')}" placeholder="Ex.: 300"><div class="field-hint">Fallback: a gramatura principal agora é cadastrada por <b>cor</b>. Este valor só é usado quando a cor não tem gramatura própria.</div></div>
        <div class="field"><label>Composição / observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: 65% algodão 35% poliéster"></div>
      </div>`;
  }
  else if (tipo === 'cor') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Preto Malha Algodão"><div class="field-hint">Inclua o <b>tecido</b> no nome (ex.: <i>Preto Malha Algodão</i>, <i>Preto Moletom</i>). A mesma cor pesa diferente em cada tecido, e é o nome que amarra a gramatura certa.</div></div>
        <div class="field"><label>Cor (hex)</label><input type="color" id="m-hex" value="${item.hex||'#c9a961'}"></div>
        <div class="field"><label>Código (ex.: Linx)</label><input type="text" id="m-codigo" value="${esc(item.codigo||'')}" placeholder="Ex.: AV.CO.129"></div>
        <div class="field"><label>Sigla SKU</label><input type="text" id="m-siglasku" value="${esc(item.siglaSku||'')}" placeholder="Ex.: PRE, VERM, OFF"><div class="field-hint">Compõe o SKU do produto acabado (ex.: CM.LISA-<b>PRE</b>)</div></div>
        <div class="field"><label>Peso / gramatura (g/m²)</label><input type="number" min="0" step="1" id="m-cor-peso" value="${esc(item.peso||'')}" placeholder="Ex.: 300"><div class="field-hint">A gramatura é por COR+TECIDO (o nome da cor traz o tecido). Base da estimativa em kg da folha de OS: comp × larg × camadas × gramatura ÷ 1000. Se vazia, cai no peso do tecido.</div></div>
      </div>`;
  }
  else if (tipo === 'material') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Código *</label><input type="text" id="m-codigo" value="${esc(item.codigo||'')}" placeholder="Ex.: AV.IN.848"></div>
        <div class="field"><label>Tipo</label><input type="text" id="m-tipo" value="${esc(item.tipo||'')}" placeholder="Ex.: Cordão, Ilhós, Tag"></div>
        <div class="field full"><label>Descrição *</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: Cordão 1,30m palha"></div>
      </div>`;
  }
  else if (tipo === 'modelo') {
    const optSel = (list, fld, id) => '<option value="">— selecione —</option>' + list.map(x => `<option value="${esc(x.id)}" ${id===x.id?'selected':''}>${esc(x[fld])}</option>`).join('');
    const optNada = (list, fld, id, labelFn) => '<option value="">— nenhum —</option>' + list.map(x => `<option value="${esc(x.id)}" ${id===x.id?'selected':''}>${esc(labelFn ? labelFn(x) : x[fld])}</option>`).join('');
    const equipeLabel = p => p.nome + (p.funcao ? ' ('+p.funcao+')' : '');
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Descrição *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Camiseta básica, Moletom canguru"></div>
        <div class="field"><label>Tipo *</label>
          <select id="m-categoria">
            <option value="">— selecione —</option>
            <option value="malha" ${item.categoria==='malha'?'selected':''}>Camiseta (malha algodão)</option>
            <option value="moletom" ${item.categoria==='moletom'?'selected':''}>Moletom</option>
            <option value="outro" ${item.categoria==='outro'?'selected':''}>Outro</option>
          </select>
          <div class="field-hint">Define quais tecidos aparecem ao selecionar este modelo na OS</div>
        </div>
        <div class="field"><label>Linha (texto)</label><input type="text" id="m-linha" value="${esc(item.linha||'')}" placeholder="Ex.: Adulto, Infantil"></div>
        <div class="field"><label>SKU</label><input type="text" id="m-skulinha" list="dl-skus" value="${esc(item.skuLinha||'')}" placeholder="Ex.: CM.LISA (linha) ou CM.LISA-PRE">${datalistSkusHtml()}<div class="field-hint">Escolha o <b>SKU completo</b> ou a <b>linha</b> (SKU da OS = linha + cor). Padrão do modelo; o desenho pode sobrescrever.</div></div>
      </div>
      <div style="margin-top:14px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Vínculos padrão (preenchem a OS ao selecionar este modelo)</label>
        <div class="form-grid cols-2" style="margin-top:6px;">
          <div class="field"><label>Base</label><select id="m-vinc-base">${optSel(STATE.bases, 'nome', item.baseId)}</select></div>
          <div class="field"><label>Marca / Griffe</label><select id="m-vinc-marca">${optSel(STATE.marcas, 'nome', item.marcaId)}</select></div>
          <div class="field"><label>Designer</label><select id="m-vinc-designer">${optNada(STATE.equipe.filter(p => (p.funcao||'').toLowerCase().includes('designer')), 'nome', item.designerId, equipeLabel)}</select></div>
          <div class="field"><label>Ficha técnica</label><select id="m-vinc-ftec">${optNada(STATE.equipe, 'nome', item.ftecId, equipeLabel)}</select></div>
          <div class="field"><label>Coordenador</label><select id="m-vinc-coord">${optNada(STATE.equipe, 'nome', item.coordId, equipeLabel)}</select></div>
        </div>
        <div class="field-hint">Vínculos do desenho (quando houver) têm prioridade sobre os do modelo.</div>
      </div>`;
  }
  else if (tipo === 'colecao') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Inverno 2024"></div>
        <div class="field"><label>Temporada</label><input type="text" id="m-temp" value="${esc(item.temporada||'')}" placeholder="Ex.: Outono-Inverno"></div>
      </div>`;
  }
  else if (tipo === 'grade') {
    // Aplica overrides de label (se a pasta fixa foi renomeada)
    const tiposPeca = [
      { v: '', lbl: labelTp('') === 'Sem categoria' ? '— sem categoria —' : labelTp('') },
      { v: 'camiseta', lbl: labelTp('camiseta') },
      { v: 'blusa_moletom', lbl: labelTp('blusa_moletom') },
      { v: 'outro', lbl: labelTp('outro') }
    ];
    const variacoes = [
      { v: '', lbl: labelVr('') === 'Sem variação' ? '— sem variação —' : labelVr('') },
      { v: 'basica', lbl: labelVr('basica') },
      { v: 'bicolor', lbl: labelVr('bicolor') },
      { v: 'tricolor', lbl: labelVr('tricolor') }
    ];
    const fixosTp = new Set(tiposPeca.map(t => t.v));
    const fixosVr = new Set(variacoes.map(t => t.v));
    const tiposPecaCustom = [...new Set(STATE.grades.map(g => g.tipoPeca || '').filter(x => x && !fixosTp.has(x)))]
      .sort((a,b)=>a.localeCompare(b,'pt-BR'));
    const variacoesCustom = [...new Set(STATE.grades.map(g => g.variacao || '').filter(x => x && !fixosVr.has(x)))]
      .sort((a,b)=>a.localeCompare(b,'pt-BR'));
    // garante que o valor atual do item apareça mesmo se ainda não estiver em STATE.grades
    if (item.tipoPeca && !fixosTp.has(item.tipoPeca) && !tiposPecaCustom.includes(item.tipoPeca)) tiposPecaCustom.push(item.tipoPeca);
    if (item.variacao && !fixosVr.has(item.variacao) && !variacoesCustom.includes(item.variacao)) variacoesCustom.push(item.variacao);

    const optsTp = tiposPeca.map(t => `<option value="${esc(t.v)}" ${item.tipoPeca===t.v?'selected':''}>${esc(t.lbl)}</option>`).join('')
      + (tiposPecaCustom.length ? `<optgroup label="Pastas adicionais">${tiposPecaCustom.map(v => `<option value="${esc(v)}" ${item.tipoPeca===v?'selected':''}>${esc(v)}</option>`).join('')}</optgroup>` : '')
      + `<option value="__nova__">+ Nova pasta…</option>`;
    const optsVr = variacoes.map(t => `<option value="${esc(t.v)}" ${item.variacao===t.v?'selected':''}>${esc(t.lbl)}</option>`).join('')
      + (variacoesCustom.length ? `<optgroup label="Subpastas adicionais">${variacoesCustom.map(v => `<option value="${esc(v)}" ${item.variacao===v?'selected':''}>${esc(v)}</option>`).join('')}</optgroup>` : '')
      + `<option value="__nova__">+ Nova subpasta…</option>`;

    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Grade padrão 6 peças"></div>
        <div class="field"><label>Tipo de peça (pasta)</label>
          <select id="m-grade-tipopeca" data-prev="${esc(item.tipoPeca||'')}" onchange="onSelectGradeFolder(this,'pasta')">
            ${optsTp}
          </select>
        </div>
        <div class="field"><label>Variação (subpasta)</label>
          <select id="m-grade-variacao" data-prev="${esc(item.variacao||'')}" onchange="onSelectGradeFolder(this,'subpasta')">
            ${optsVr}
          </select>
        </div>
      </div>
      <div style="margin-top:10px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Distribuição por tamanho</label>
        <div class="grade-inputs" style="margin-top:6px;">
          ${['p','m','g','gg','g1','g2','g3'].map(t => `
            <div class="field"><label>${t.toUpperCase()}</label><input type="number" min="0" id="m-gr-${t}" value="${item.tamanhos?.[t]||0}"></div>
          `).join('')}
        </div>
      </div>
      <div style="margin-top:14px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Fases do enfesto</label>
        <div class="field-hint" style="margin-top:4px;margin-bottom:8px;">
          Peças básicas/unicolor usam 1 fase. Bicolor → 2 fases. Tricolor → 3 fases. Pode adicionar mais conforme precisar.
          Cada fase tem seu próprio tecido, cor e dimensões.
        </div>
        <div id="m-fases-container"></div>
        <button type="button" class="add-row-btn" onclick="addFaseGradeRow()" style="margin-top:8px;">+ Adicionar fase</button>
      </div>`;
    // Popula o container com as fases existentes (ou uma fase vazia em "Novo")
    setTimeout(() => {
      const fasesSalvas = Array.isArray(item.fases) && item.fases.length ? item.fases : null;
      const legacy = fasesSalvas ? null : {
        comp: item.enfestoComprimento
               || item.enfestos?.outro?.comp
               || item.enfestos?.malha?.comp
               || item.enfestos?.moletom?.comp
               || '',
        larg: item.enfestoLargura
               || item.enfestos?.outro?.larg
               || item.enfestos?.malha?.larg
               || item.enfestos?.moletom?.larg
               || ''
      };
      if (fasesSalvas) {
        // Preserva ordem: cria buracos se houver (ordem 2 e 3 sem ordem 1)
        const porOrdem = {};
        fasesSalvas.forEach(f => { if (f.ordem) porOrdem[f.ordem] = f; });
        const maxOrd = Math.max(...fasesSalvas.map(f => f.ordem || 1), 1);
        for (let n = 1; n <= maxOrd; n++) addFaseGradeRow(porOrdem[n] || {});
      } else {
        addFaseGradeRow(legacy || {});
      }
    }, 0);
  }
  else if (tipo === 'desenho') {
    const optSel = (list, fld, id) => '<option value="">— selecione —</option>' + list.map(x => `<option value="${esc(x.id)}" ${id===x.id?'selected':''}>${esc(x[fld])}</option>`).join('');
    const optNada = (list, fld, id, labelFn) => '<option value="">— nenhum —</option>' + list.map(x => `<option value="${esc(x.id)}" ${id===x.id?'selected':''}>${esc(labelFn ? labelFn(x) : x[fld])}</option>`).join('');
    const equipeLabel = p => p.nome + (p.funcao ? ' ('+p.funcao+')' : '');
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Código *</label><input type="text" id="m-codigo" value="${esc(item.codigo||'')}" placeholder="Ex.: Dx7282"></div>
        <div class="field"><label>Descrição</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: Camiseta básica preta"></div>
        <div class="field"><label>SKU</label><input type="text" id="m-desenho-sku" list="dl-skus" value="${esc(item.skuLinha||'')}" placeholder="Escolha o SKU (ex.: CM.LISA-PRE)">${datalistSkusHtml()}<div class="field-hint">Escolha o <b>SKU completo</b> (ex.: CM.LISA-PRE) ou só a <b>linha</b> (ex.: CM.LISA — a cor resolve pela OS). Tem prioridade sobre o modelo.</div></div>
        <div class="field full">
          <label>Imagem (PNG/JPG) *</label>
          <label class="file-label">Escolher arquivo <input type="file" id="m-img" accept="image/*" onchange="previewUploadImg(event)"></label>
          <div class="desenho-preview" id="m-img-preview" style="margin-top:8px;">
            ${item.img ? `<img src="${item.img}">` : '<span>Sem imagem</span>'}
          </div>
          <input type="hidden" id="m-img-data" value="${item.img||''}">
        </div>
      </div>
      <div style="margin-top:14px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Vínculos (preenchem automaticamente a OS quando este desenho for selecionado)</label>
        <div class="form-grid fit-cols" style="margin-top:6px;">
          <div class="field"><label>Modelo</label><select id="m-vinc-modelo">${optSel(STATE.modelos, 'nome', item.modeloId)}</select></div>
          <div class="field"><label>Base</label><select id="m-vinc-base">${optSel(STATE.bases, 'nome', item.baseId)}</select></div>
          <div class="field"><label>Coleção</label><select id="m-vinc-colecao">${optSel(STATE.colecoes, 'nome', item.colecaoId)}</select></div>
          <div class="field"><label>Marca / Griffe</label><select id="m-vinc-marca">${optSel(STATE.marcas, 'nome', item.marcaId)}</select></div>
          <div class="field"><label>Linha</label><select id="m-vinc-linha">${optSel(STATE.linhas, 'nome', item.linhaId)}</select></div>
          <div class="field"><label>Bloco / Revisão</label><select id="m-vinc-bloco">${optNada(STATE.blocos, 'nome', item.blocoId)}</select></div>
          <div class="field"><label>Designer</label><select id="m-vinc-designer">${optNada(STATE.equipe.filter(p => (p.funcao||'').toLowerCase().includes('designer')), 'nome', item.designerId, equipeLabel)}</select></div>
          <div class="field"><label>Coordenador</label><select id="m-vinc-coord">${optNada(STATE.equipe.filter(p => (p.funcao||'').toLowerCase().includes('coordenador')), 'nome', item.coordId, equipeLabel)}</select><div class="field-hint">Auto-preenche o Coordenador na OS</div></div>
          <div class="field"><label>Tecido principal</label><select id="m-vinc-tecido">${optNada(STATE.tecidos, 'nome', item.tecidoPadraoId)}</select><div class="field-hint">Aplicado aos componentes</div></div>
          <div class="field"><label>Cor principal</label><select id="m-vinc-cor">${optNada(STATE.cores, 'nome', item.corPrincipalId)}</select><div class="field-hint">Aplicada aos componentes e à Cor 1 da Variante 1</div></div>
          <div class="field"><label>Cor secundária (bicolor)</label><select id="m-vinc-cor2">${optNada(STATE.cores, 'nome', item.corSecundariaId)}</select><div class="field-hint">Opcional — aplicada à Cor 2 da Variante 1</div></div>
          <div class="field"><label>Cor terciária (tricolor)</label><select id="m-vinc-cor3">${optNada(STATE.cores, 'nome', item.corTerciariaId)}</select><div class="field-hint">Opcional — aplicada à Cor 3 da Variante 1</div></div>
        </div>
        <div style="margin-top:14px;">
          <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Componentes padrão deste desenho</label>
          <div class="field-hint" style="margin-top:4px;margin-bottom:6px;">
            Marque os componentes usados e preencha tecido, cor e quantidade por peça.
            O total por tamanho é calculado automaticamente na OS.
          </div>
          <div style="padding:8px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);">
            ${STATE.componentes.length ? (() => {
              const tecOpts = (selId) => '<option value="">—</option>' + STATE.tecidos.map(t =>
                `<option value="${esc(t.id)}" ${selId===t.id?'selected':''}>${esc(t.nome)}</option>`).join('');
              const corOpts = (selId) => '<option value="">—</option>' + STATE.cores.map(c =>
                `<option value="${esc(c.id)}" ${selId===c.id?'selected':''}>${esc(c.nome)}</option>`).join('');
              // Retrocompat: se tem componentesIds mas não tem componentes (nova estrutura), mapeia
              const compsAtuais = Array.isArray(item.componentes) && item.componentes.length
                ? item.componentes
                : (item.componentesIds || []).map(id => ({
                    componenteId: id, tecidoId: item.tecidoPadraoId || '', corId: item.corPrincipalId || '', qtdPorPeca: 1
                  }));
              const porId = new Map(compsAtuais.map(c => [c.componenteId, c]));
              // Fallback por NOME: após a perda/restauração, os componentes salvos no
              // desenho podem referenciar IDs antigos que não existem mais no cadastro
              // global (recriado com IDs novos). Sem casar por nome, as linhas apareciam
              // desmarcadas e sem cor — e como o save só grava as linhas MARCADAS, as
              // cores de forro/punho/barra eram apagadas a cada gravação. Casando por
              // nome, elas reaparecem marcadas e o save migra o componenteId pro novo.
              const porNome = new Map();
              compsAtuais.forEach(c => { const k = _normNome(c.nome); if (k && !porNome.has(k)) porNome.set(k, c); });
              return `<table class="desenho-comp-table">
                <thead><tr>
                  <th style="width:24px;"></th>
                  <th>Componente</th>
                  <th>Tecido</th>
                  <th>Cor</th>
                  <th style="width:72px;">Qtd/peça</th>
                </tr></thead>
                <tbody>
                ${STATE.componentes.map(c => {
                  const atual = porId.get(c.id) || porNome.get(_normNome(c.nome)) || {};
                  const marcado = porId.has(c.id) || porNome.has(_normNome(c.nome));
                  return `<tr class="desenho-comp-row">
                    <td style="text-align:center;"><input type="checkbox" class="m-componente-chk" value="${esc(c.id)}" ${marcado?'checked':''}></td>
                    <td>${esc(c.nome)}</td>
                    <td><select class="m-comp-tec" data-comp="${esc(c.id)}">${tecOpts(atual.tecidoId)}</select></td>
                    <td><select class="m-comp-cor" data-comp="${esc(c.id)}">${corOpts(atual.corId)}</select></td>
                    <td>${(()=>{
                      const sel = Math.round(Number(atual.qtdPorPeca)) || 1;
                      const opts = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
                        .map(n => `<option value="${n}" ${sel===n?'selected':''}>${n}</option>`).join('');
                      return `<select class="m-comp-qtd" data-comp="${esc(c.id)}">${opts}</select>`;
                    })()}</td>
                  </tr>`;
                }).join('')}
                </tbody>
              </table>`;
            })() : '<em style="color:var(--ink-3);font-size:12px;">Cadastre componentes primeiro em Componentes.</em>'}
          </div>
        </div>
        <div style="margin-top:14px;">
          <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Aviamentos padrão deste desenho</label>
          <div class="field-hint" style="margin-top:4px;margin-bottom:6px;">
            Marque os aviamentos usados e preencha quantidade por peça e aplicação (unidade sempre "un").
          </div>
          <div style="padding:8px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);">
            ${STATE.materiais.length ? (() => {
              const avsAtuais = Array.isArray(item.aviamentos) && item.aviamentos.length
                ? item.aviamentos
                : (item.aviamentosIds || []).map(id => ({ materialId: id, qtdPorPeca: 1, aplicacao: '' }));
              const porId = new Map(avsAtuais.map(a => [a.materialId, a]));
              return `<table class="desenho-comp-table">
                <thead><tr>
                  <th style="width:24px;"></th>
                  <th>Aviamento</th>
                  <th style="width:72px;">Qtd/peça</th>
                  <th>Aplicação</th>
                </tr></thead>
                <tbody>
                ${STATE.materiais.map(m => {
                  const atual = porId.get(m.id) || {};
                  const marcado = porId.has(m.id);
                  return `<tr class="desenho-comp-row">
                    <td style="text-align:center;"><input type="checkbox" class="m-aviamento-chk" value="${esc(m.id)}" ${marcado?'checked':''}></td>
                    <td><strong>${esc(m.codigo)}</strong> · ${esc(m.desc)}${m.tipo ? ' ('+esc(m.tipo)+')' : ''}</td>
                    <td>${(()=>{
                      const sel = Math.round(Number(atual.qtdPorPeca)) || 1;
                      const opts = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
                        .map(n => `<option value="${n}" ${sel===n?'selected':''}>${n}</option>`).join('');
                      return `<select class="m-av-qtd" data-av="${esc(m.id)}">${opts}</select>`;
                    })()}</td>
                    <td><input type="text" class="m-av-app" data-av="${esc(m.id)}" value="${esc(atual.aplicacao || '')}" placeholder="Ex.: V1 camel / V2 preto"></td>
                  </tr>`;
                }).join('')}
                </tbody>
              </table>`;
            })() : '<em style="color:var(--ink-3);font-size:12px;">Cadastre aviamentos primeiro em Materiais / Aviamentos.</em>'}
          </div>
        </div>
        <div style="margin-top:14px;">
          <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Etapas padrão deste desenho (na ordem de execução)</label>
          <div class="field-hint" style="margin-top:4px;margin-bottom:6px;">
            Marque as etapas que este desenho usa e use ▲▼ pra ordená-las. Ao selecionar o desenho na OS, as etapas já vêm marcadas e na ordem certa.
          </div>
          <div id="m-desenho-etapas" style="padding:8px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);">
            ${(() => {
              const cadastradas = etapasOrdenadas();
              if (!cadastradas.length) return '<em style="color:var(--ink-3);font-size:12px;">Cadastre etapas primeiro em Etapas de produção.</em>';
              const salvas = Array.isArray(item.etapasNomes) && item.etapasNomes.length ? item.etapasNomes : [];
              // Primeiro as salvas na ordem, depois as não-marcadas na ordem cadastrada
              const resto = cadastradas.filter(e => !salvas.includes(e.nome));
              const ordem = [
                ...salvas.map(n => cadastradas.find(e => e.nome === n)).filter(Boolean),
                ...resto
              ];
              return ordem.map(e => `
                <label class="etapa-check ${salvas.includes(e.nome)?'checked':''}" style="margin-bottom:4px;">
                  <span class="etapa-reorder">
                    <button type="button" class="etapa-move" onclick="event.preventDefault(); event.stopPropagation(); moverEtapaDesenho(this, -1)" title="Mover para cima">▲</button>
                    <button type="button" class="etapa-move" onclick="event.preventDefault(); event.stopPropagation(); moverEtapaDesenho(this, 1)" title="Mover para baixo">▼</button>
                  </span>
                  <input type="checkbox" class="m-etapa-chk" value="${esc(e.nome)}" ${salvas.includes(e.nome)?'checked':''} onchange="this.parentElement.classList.toggle('checked', this.checked)">
                  <span>${esc(e.nome)}</span>
                </label>`).join('');
            })()}
          </div>
        </div>
      </div>`;
  }
  else if (tipo === 'marca') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Dixie"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: marca principal"></div>
      </div>`;
  }
  else if (tipo === 'linha') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Adulto"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: linha principal"></div>
      </div>`;
  }
  else if (tipo === 'base') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: BASE M MOLETOM"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: molde base nº 12"></div>
      </div>`;
  }
  else if (tipo === 'bloco') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: R1 BLOCO 2"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: segunda revisão do bloco"></div>
      </div>`;
  }
  else if (tipo === 'equipe') {
    const curVal = item.funcao || '';
    const nomesCadastrados = STATE.funcoes.map(f => f.nome);
    const todosNomes = [...new Set([curVal, ...nomesCadastrados].filter(Boolean))];
    const funcoesOpts = todosNomes.map(nome => {
      const inCad = nomesCadastrados.includes(nome);
      return `<option value="${esc(nome)}" ${curVal===nome?'selected':''}>${esc(nome)}${inCad?'':' (não cadastrada)'}</option>`;
    }).join('');
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Marcelo"></div>
        <div class="field"><label>Função principal</label>
          <select id="m-funcao" onchange="mostrarResponsabilidadesFuncao()">
            <option value="">— sem função —</option>
            ${funcoesOpts}
          </select>
          <div id="m-funcao-resp" class="field-hint" style="min-height:16px;"></div>
          <div class="field-hint">Cadastre funções em <a href="#" onclick="closeModal('modal-cad'); goto('cad-funcoes'); return false;">Funções</a></div>
        </div>
      </div>`;
    setTimeout(mostrarResponsabilidadesFuncao, 0);
  }
  else if (tipo === 'funcao') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Costureira"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Opcional"></div>
        <div class="field full"><label>Responsabilidades / ações</label>
          <textarea id="m-acoes" rows="3" placeholder="Ex.: Costurar peças, fazer travete, acabamento. Uma por linha.">${esc(item.acoes||'')}</textarea>
        </div>
      </div>`;
  }
  else if (tipo === 'tarefa') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Costurar manga, Pregar etiqueta"></div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Opcional"></div>
      </div>`;
  }
  else if (tipo === 'componente') {
    const variacoes = [
      { v: '', lbl: '— sem variação —' },
      { v: 'basica', lbl: 'Básica' },
      { v: 'bicolor', lbl: 'Bicolor' },
      { v: 'tricolor', lbl: 'Tricolor' }
    ];
    const corOpts = (selId) => '<option value="">— selecione —</option>' + STATE.cores.map(c =>
      `<option value="${esc(c.id)}" ${selId===c.id?'selected':''}>${esc(c.nome)}</option>`).join('');
    const semCores = !STATE.cores.length;
    const semModelos = !STATE.modelos.length;
    // Retrocompat: se o valor antigo era slug (camiseta/blusa_moletom/outro), tenta achar um modelo equivalente pelo nome/categoria
    const tipoSalvo = item.tipoPeca || '';
    const modeloOpts = '<option value="">— selecione —</option>' + STATE.modelos.map(m =>
      `<option value="${esc(m.id)}" ${tipoSalvo===m.id?'selected':''}>${esc(m.nome)}${m.linha?' ('+esc(m.linha)+')':''}</option>`).join('');
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Frente, Costas, Mangas"></div>
        <div class="field"><label>Tipo (modelo)</label>
          <select id="m-comp-tipopeca" ${semModelos?'disabled':''}>
            ${semModelos ? '<option value="">— cadastre modelos primeiro —</option>' : modeloOpts}
          </select>
          <div class="field-hint">${semModelos ? 'Cadastre em <strong>Modelos</strong> para liberar este campo.' : 'Lista vem de Modelos cadastrados.'}</div>
        </div>
        <div class="field"><label>Variação</label>
          <select id="m-comp-variacao" onchange="atualizarCoresComponente()">
            ${variacoes.map(x => `<option value="${x.v}" ${item.variacao===x.v?'selected':''}>${x.lbl}</option>`).join('')}
          </select>
          <div class="field-hint">Básica = 1 cor · Bicolor = 2 cores · Tricolor = 3 cores</div>
        </div>
        <div class="field full"><label>Observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Opcional"></div>
      </div>
      <div id="m-comp-cores-wrap" style="margin-top:14px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Cores deste componente</label>
        <div class="field-hint" style="margin-top:4px;margin-bottom:6px;">
          ${semCores
            ? 'Cadastre cores primeiro em <strong>Cores</strong> para poder selecioná-las aqui.'
            : 'Selecione as cores na ordem de aplicação (Cor 1, Cor 2, Cor 3).'}
        </div>
        <div class="form-grid cols-3">
          <div class="field" id="m-comp-cor1-wrap"><label>Cor 1</label><select id="m-comp-cor1" ${semCores?'disabled':''}>${corOpts(item.cor1Id)}</select></div>
          <div class="field" id="m-comp-cor2-wrap"><label>Cor 2</label><select id="m-comp-cor2" ${semCores?'disabled':''}>${corOpts(item.cor2Id)}</select></div>
          <div class="field" id="m-comp-cor3-wrap"><label>Cor 3</label><select id="m-comp-cor3" ${semCores?'disabled':''}>${corOpts(item.cor3Id)}</select></div>
        </div>
      </div>`;
    setTimeout(atualizarCoresComponente, 0);
  }
  else if (tipo === 'etapa') {
    const ordemSugerida = item.ordem ?? ((STATE.etapas.length + 1) * 10);
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Corte, Costura, Acabamento"></div>
        <div class="field"><label>Ordem</label><input type="number" id="m-ordem" value="${ordemSugerida}" placeholder="Ex.: 10"><div class="field-hint">Menor primeiro na folha impressa</div></div>
      </div>
      <div style="margin-top:14px;">
        <label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Tarefas desta etapa</label>
        <div class="field-hint" style="margin-top:4px;margin-bottom:8px;">
          Adicione as tarefas que compõem esta etapa (subitens). Cada tarefa tem nome e observação opcional.
        </div>
        <div id="m-tarefas-container"></div>
        <button type="button" class="add-row-btn" onclick="addTarefaEtapaRow()" style="margin-top:8px;">+ Adicionar tarefa</button>
      </div>`;
    // Popula tarefas existentes (ou uma vazia em "Novo")
    setTimeout(() => {
      let tarefasSalvas = Array.isArray(item.tarefas) && item.tarefas.length ? item.tarefas : null;
      // Retrocompat: migra de tarefasIds (modelo antigo) buscando em STATE.tarefas
      if (!tarefasSalvas && Array.isArray(item.tarefasIds) && item.tarefasIds.length) {
        tarefasSalvas = item.tarefasIds
          .map(tid => (STATE.tarefas || []).find(t => t.id === tid))
          .filter(Boolean)
          .map(t => ({ nome: t.nome || '', desc: t.desc || '' }));
      }
      if (tarefasSalvas && tarefasSalvas.length) tarefasSalvas.forEach(t => addTarefaEtapaRow(t));
      else addTarefaEtapaRow();
    }, 0);
  }

  openModal('modal-cad');
  // Recarrega o catálogo de SKUs na hora ao abrir Desenho/Modelo, pra o dropdown
  // não depender do que foi lido no login (auto-cura se o catálogo subiu depois).
  if (tipo === 'desenho' || tipo === 'modelo') refreshDatalistSkus();
}

// Recarrega o catálogo do Supabase e reinjeta as opções no <datalist id="dl-skus">.
async function refreshDatalistSkus() {
  try {
    await carregarCatalogoSkus();
    const dl = document.getElementById('dl-skus');
    if (dl) {
      dl.innerHTML = (catalogoSkus || [])
        .map(s => `<option value="${esc(s.item)}">${esc(s.descricao || s.item)}</option>`)
        .join('');
    }
  } catch (e) { /* silencioso */ }
}

// Interpreta o campo "bobinas previstas": aceita inteiro (14), fração (1/2),
// decimal com vírgula (0,5) e zero. Retorna número, ou null se não informado.
function parseBobinas(str) {
  const s = String(str == null ? '' : str).trim().replace(',', '.');
  if (s === '') return null;
  if (s.includes('/')) {
    const [a, b] = s.split('/').map(x => parseFloat(x));
    return (isFinite(a) && isFinite(b) && b > 0 && a >= 0) ? a / b : null;
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function addFaseGradeRow(fase = {}) {
  const cont = document.getElementById('m-fases-container');
  if (!cont) return;
  const tecOpts = (selId) => '<option value="">— selecione —</option>' + STATE.tecidos.map(t =>
    `<option value="${esc(t.id)}" ${selId===t.id?'selected':''}>${esc(t.nome)}${t.categoria?' ('+esc(t.categoria)+')':''}</option>`).join('');
  const unidadesAtual = parseInt(fase.unidades) || 2;
  const unidadesOpts = [1, 2, 4, 6, 8, 10, 20].map(n =>
    `<option value="${n}" ${unidadesAtual === n ? 'selected' : ''}>${n}x</option>`).join('');
  const div = document.createElement('div');
  div.className = 'fase-grade-bloco';
  div.style.cssText = 'margin-top:8px;padding:10px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);';
  div.innerHTML = `
    <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:6px;">
      <span class="fase-label" style="font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:12px;color:var(--ink);">FASE ?</span>
      <span style="flex:1;"></span>
      <button type="button" class="btn small danger" onclick="removerFaseGrade(this)">✕ Remover</button>
    </div>
    <div class="form-grid cols-2">
      <div class="field full"><label>Nome da fase (opcional)</label><input type="text" class="fase-nome" value="${esc(fase.nome || '')}" placeholder="Ex.: Moletom, Forro de capuz, Punhos, Barra"></div>
      <div class="field"><label>Tecido</label><select class="fase-tec" onchange="toggleUnidadesGrade(this)">${tecOpts(fase.tecidoId)}</select></div>
      <div class="field fase-unid-wrap"><label>Unidades da grade</label><select class="fase-unid">${unidadesOpts}</select><div class="field-hint">1 unidade da grade = N peças por camada (ribana). Ex.: 2x para Barra+Punhos moletom, 10x ou 20x para Gola.</div></div>
      <div class="field"><label>Comprimento (m)</label><input type="number" step="0.01" class="fase-comp" value="${esc(fase.comp || '')}" placeholder="Ex.: 6,50"></div>
      <div class="field"><label>Largura (m)</label><input type="number" step="0.01" class="fase-larg" value="${esc(fase.larg || '')}" placeholder="Ex.: 1,80"></div>
      <div class="field full"><label>Bobinas previstas (consumo esperado)</label><input type="text" class="fase-bobinas" value="${esc(fase.bobinas != null && fase.bobinas !== '' ? String(fase.bobinas).replace('.', ',') : '')}" placeholder="Ex.: 14  ·  1/2  ·  0"><div class="field-hint">Quantas bobinas deste tecido esta grade costuma consumir nesta fase. Aparece na coluna "Consumo" da folha de OS. Aceita fração (1/2) e zero.</div></div>
    </div>`;
  cont.appendChild(div);
  toggleUnidadesGrade(div.querySelector('.fase-tec'));
  renumerarFasesGrade();
}

function toggleUnidadesGrade(selectEl) {
  const bloco = selectEl?.closest?.('.fase-grade-bloco');
  if (!bloco) return;
  const wrap = bloco.querySelector('.fase-unid-wrap');
  if (!wrap) return;
  const tec = STATE.tecidos.find(t => t.id === selectEl.value);
  wrap.style.display = isTecidoRibana(tec) ? '' : 'none';
}

function removerFaseGrade(btn) {
  const bloco = btn.closest('.fase-grade-bloco');
  if (bloco) bloco.remove();
  renumerarFasesGrade();
}

function renumerarFasesGrade() {
  const cont = document.getElementById('m-fases-container');
  if (!cont) return;
  Array.from(cont.querySelectorAll('.fase-grade-bloco')).forEach((b, i) => {
    const lbl = b.querySelector('.fase-label');
    if (lbl) lbl.textContent = `FASE ${i+1}`;
  });
}

function addTarefaEtapaRow(tarefa = {}) {
  const cont = document.getElementById('m-tarefas-container');
  if (!cont) return;
  const div = document.createElement('div');
  div.className = 'tarefa-etapa-bloco';
  div.style.cssText = 'margin-top:8px;padding:10px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);';
  div.innerHTML = `
    <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:6px;">
      <span class="tarefa-label" style="font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:12px;color:var(--ink);">TAREFA ?</span>
      <span style="flex:1;"></span>
      <button type="button" class="btn small danger" onclick="removerTarefaEtapa(this)">✕ Remover</button>
    </div>
    <div class="form-grid cols-2">
      <div class="field full"><label>Nome *</label><input type="text" class="tarefa-nome" value="${esc(tarefa.nome || '')}" placeholder="Ex.: Costurar manga, Pregar etiqueta"></div>
      <div class="field full"><label>Observação</label><input type="text" class="tarefa-desc" value="${esc(tarefa.desc || '')}" placeholder="Opcional"></div>
    </div>`;
  cont.appendChild(div);
  renumerarTarefasEtapa();
}

function removerTarefaEtapa(btn) {
  const bloco = btn.closest('.tarefa-etapa-bloco');
  if (bloco) bloco.remove();
  renumerarTarefasEtapa();
}

function renumerarTarefasEtapa() {
  const cont = document.getElementById('m-tarefas-container');
  if (!cont) return;
  Array.from(cont.querySelectorAll('.tarefa-etapa-bloco')).forEach((b, i) => {
    const lbl = b.querySelector('.tarefa-label');
    if (lbl) lbl.textContent = `TAREFA ${i+1}`;
  });
}

// Forca um reload completo dos dados do Supabase para o STATE em memoria.
// Util quando o cache local diverge do servidor (ex.: cadastros feitos em
// outra sessao/aba que ainda nao chegaram nesta).
async function recarregarDadosDoServidor() {
  if (!exigirAdmin('recarregar dados do servidor')) return;
  if (!supa) { toast('Supabase nao carregado', 'err'); return; }
  toast('Recarregando do servidor...', '');
  try {
    await loadState();
    // Re-renderiza a pagina atual
    const ativa = document.querySelector('.page:not(.hidden)');
    const pagina = ativa?.dataset?.page || 'home';
    goto(pagina);
    toast('Dados atualizados do servidor', 'ok');
  } catch (e) {
    console.error('Falha ao recarregar:', e);
    toast('Erro ao recarregar dados', 'err');
  }
}

// Handler do botao de Configuracoes: le o codigo digitado, confirma e dispara a copia.
// Chave do MODELO de um desenho, usada nas operações em massa para só copiar
// entre desenhos do MESMO modelo E mesma variação. O modeloId já distingue
// "Camiseta Básica" de "Camiseta Bicolor"/"Camiseta Tricolor" e "Blusa Moletom
// Básica" de "Blusa Moletom Tricolor" — então camiseta básica NÃO copia para
// bicolor/tricolor, e moletom básico NÃO copia para moletom tricolor. Se o
// desenho não tiver modeloId (legado), cai para a parte do desc antes do "|"
// (ex.: "Camiseta Básica | Preto" -> "camiseta básica").
function chaveModeloDesenho(d) {
  if (!d) return '';
  const id = (d.modeloId || '').trim();
  if (id) return 'm:' + id;
  const nomeDesc = ((d.desc || '').split('|')[0] || '').trim().toLowerCase();
  return nomeDesc ? 'd:' + nomeDesc : '';
}

// Rótulo amigável do modelo de um desenho (parte do desc antes do "|"), para
// mensagens. Fallback para o nome do modelo vinculado, senão o modeloId.
function rotuloModeloDesenho(d) {
  if (!d) return 'mesmo modelo';
  const nomeDesc = ((d.desc || '').split('|')[0] || '').trim();
  if (nomeDesc) return nomeDesc;
  const m = (STATE.modelos || []).find(x => x.id === d.modeloId);
  return (m && m.nome) || d.modeloId || 'mesmo modelo';
}

async function rodarCopiarEtapasParaTodos() {
  const input = document.getElementById('copyEtapasOrigem');
  const codigo = (input?.value || '').trim();
  if (!codigo) { toast('Informe o codigo do desenho de origem', 'err'); return; }
  const origem = STATE.desenhos.find(d => (d.codigo || '').trim() === codigo);
  if (!origem) { toast(`Desenho "${codigo}" nao encontrado`, 'err'); return; }
  const etapas = Array.isArray(origem.etapasNomes) ? origem.etapasNomes : [];
  if (!etapas.length) { toast(`Desenho "${codigo}" nao tem etapas configuradas`, 'err'); return; }
  const chaveOrigem = chaveModeloDesenho(origem);
  const modeloLabel = rotuloModeloDesenho(origem);
  const alvos = STATE.desenhos.filter(d => d.id !== origem.id && chaveModeloDesenho(d) === chaveOrigem);
  if (!alvos.length) {
    toast(`Nenhum outro desenho do modelo "${modeloLabel}" para receber as etapas`, 'err');
    return;
  }
  const ok = confirm(
    `Copiar as ${etapas.length} etapas do desenho "${codigo}" para os outros ${alvos.length} desenhos do modelo "${modeloLabel}"?\n\n`
    + `Etapas: ${etapas.join(', ')}\n\n`
    + `Só desenhos do MESMO modelo/variação (${modeloLabel}) são afetados — outros modelos, e as variações bicolor/tricolor, ficam intactos. Apenas o campo "etapas" será sobrescrito. Esta ação não pode ser desfeita automaticamente.`
  );
  if (!ok) return;
  await copiarEtapasEntreDesenhos(codigo);
}

// Utilitario admin: copia as etapasNomes (e a ordem) de um desenho de origem
// para os demais desenhos DO MESMO MODELO/variação (mesmo modeloId). Uso:
// copiarEtapasEntreDesenhos('001'). Outros modelos ficam intactos.
async function copiarEtapasEntreDesenhos(codigoOrigem) {
  if (!exigirAdmin('copiar etapas entre desenhos')) return;
  const origem = STATE.desenhos.find(d => (d.codigo || '').trim() === String(codigoOrigem).trim());
  if (!origem) {
    toast(`Desenho "${codigoOrigem}" nao encontrado`, 'err');
    return;
  }
  const etapasNomes = Array.isArray(origem.etapasNomes) ? [...origem.etapasNomes] : [];
  if (!etapasNomes.length) {
    toast(`Desenho "${codigoOrigem}" nao tem etapas configuradas`, 'err');
    return;
  }
  const chaveOrigem = chaveModeloDesenho(origem);
  const modeloLabel = rotuloModeloDesenho(origem);
  let alteradas = 0;
  STATE.desenhos.forEach(d => {
    if (d.id === origem.id) return;
    if (chaveModeloDesenho(d) !== chaveOrigem) return;   // só mesmo modelo/variação
    d.etapasNomes = [...etapasNomes];
    alteradas++;
  });
  await saveState('desenhos');
  toast(`Etapas de "${codigoOrigem}" aplicadas a ${alteradas} desenho(s) do modelo "${modeloLabel}"`, 'ok');
  if (typeof renderDesenhos === 'function') renderDesenhos();
  return { origem: codigoOrigem, modelo: modeloLabel, etapas: etapasNomes, alteradas };
}

// Ordem CANÔNICA das cores de um desenho = a sequência escrita no desc, após o
// último "|". Ex.: "Blusa Moletom Tricolor | Verde/Preto/Bege" -> ['verde','preto','bege'].
// É a ordem que o usuário mantém no cadastro (e que aparece no banner). Os campos
// corPrincipalId/Sec/Ter do desenho podem estar numa ordem DIVERGENTE do desc — ex.:
// desenho 0024 tem desc "Verde/Preto/Bege" mas campos "Preto/Verde/Bege" (efeito de
// restauração de dados). Como o enfesto/tecidos mapeiam a 1ª/2ª/3ª fase de corpo à
// cor primária/secundária/terciária POR ÍNDICE, essa divergência trocava as cores das
// fases. Este é o ponto único de verdade: banner, enfesto, tecidos e variante ordenam
// as cores por esta sequência, então tudo sai consistente mesmo com os campos trocados.
function ordemCoresPorDesc(desenho) {
  const tail = ((desenho && desenho.desc) || '').split('|').pop() || '';
  return tail.split('/').map(s => s.trim().toLowerCase()).filter(Boolean);
}

// Reordena uma lista de IDs de cor pela ordem canônica do desc. Resolve cada id ao
// nome via STATE.cores; cores sem correspondência no desc vão pro fim mantendo a
// ordem relativa. Se o desc não tiver cores ou os nomes não resolverem, devolve a
// lista original — fallback seguro. Usada no enfesto/tecidos/variante, onde os
// campos corPrincipal/Sec/Ter podem estar numa ordem divergente do desc.
function ordenarCoresIdsPorDesc(ids, desenho) {
  const ordem = ordemCoresPorDesc(desenho).map(corBaseNome);
  if (!ordem.length) return ids;
  // Compara pela cor BASE dos dois lados: o desc escreve "Verde/Preto/Bege" e o
  // cadastro agora guarda "Verde Malha Algodão". Sem isso o indexOf nunca casaria
  // e a ordem cairia no fallback, trocando as cores das fases do enfesto.
  const nome = id => corBaseNome(((STATE.cores || []).find(c => c.id === id) || {}).nome);
  return ids
    .map((id, i) => ({ id, i, pos: ordem.indexOf(nome(id)) }))
    .sort((a, b) => (a.pos < 0 ? 99 : a.pos) - (b.pos < 0 ? 99 : b.pos) || a.i - b.i)
    .map(x => x.id);
}

// Reordena uma lista de NOMES de cor pela ordem canônica do desc (sem depender de
// STATE.cores — usada no banner impresso, que já tem os nomes).
function ordenarCoresNomesPorDesc(nomes, desenho) {
  const ordem = ordemCoresPorDesc(desenho).map(corBaseNome);
  if (!ordem.length) return nomes;
  return nomes
    .map((n, i) => ({ n, i, pos: ordem.indexOf(corBaseNome(n)) }))
    .sort((a, b) => (a.pos < 0 ? 99 : a.pos) - (b.pos < 0 ? 99 : b.pos) || a.i - b.i)
    .map(x => x.n);
}

function atualizarCoresComponente() {
  const sel = document.getElementById('m-comp-variacao');
  const wrap = document.getElementById('m-comp-cores-wrap');
  if (!sel || !wrap) return;
  const v = sel.value;
  const nCores = v === 'tricolor' ? 3 : v === 'bicolor' ? 2 : v === 'basica' ? 1 : 0;
  wrap.style.display = nCores === 0 ? 'none' : '';
  [1, 2, 3].forEach(i => {
    const w = document.getElementById('m-comp-cor'+i+'-wrap');
    if (w) w.style.display = i <= nCores ? '' : 'none';
  });
}

function previewUploadImg(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    document.getElementById('m-img-preview').innerHTML = `<img src="${evt.target.result}">`;
    document.getElementById('m-img-data').value = evt.target.result;
  };
  reader.readAsDataURL(file);
}

/* Converte uma dataURL base64 em Blob (para upload binário ao Storage). */
function dataUrlParaBlob(dataUrl) {
  const [meta, b64] = dataUrl.split(',');
  const mime = (meta.match(/data:([^;]+)/) || [null, 'image/png'])[1];
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return { blob: new Blob([arr], { type: mime }), mime };
}

/* Faz upload de uma imagem de desenho para o bucket 'desenhos' e retorna URL pública. */
async function uploadDesenhoImagem(dataUrl) {
  if (!supa) throw new Error('Supabase não carregado');
  const { blob, mime } = dataUrlParaBlob(dataUrl);
  const ext = mime.split('/')[1] || 'png';
  const nome = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}.${ext}`;
  const { error } = await supa.storage.from('desenhos').upload(nome, blob, {
    contentType: mime,
    upsert: false
  });
  if (error) throw error;
  const { data } = supa.storage.from('desenhos').getPublicUrl(nome);
  return data.publicUrl;
}

/* Migra imagens base64 legadas para o Storage (roda uma vez, só para admin). */
async function migrarImagensBase64() {
  if (!supa || !currentUser) return;
  const pendentes = (STATE.desenhos || []).filter(d => typeof d.img === 'string' && d.img.startsWith('data:image/'));
  if (!pendentes.length) return;
  toast(`Migrando ${pendentes.length} imagem(ns) para o Storage...`, 'ok');
  let migradas = 0;
  for (const d of pendentes) {
    try {
      const url = await uploadDesenhoImagem(d.img);
      d.img = url;
      migradas++;
    } catch (e) {
      console.error('Falha ao migrar imagem do desenho', d.codigo, e);
    }
  }
  if (migradas > 0) {
    await saveState('desenhos');
    toast(`${migradas} imagem(ns) migrada(s) para Storage`, 'ok');
  }
}

function pluralize(tipo) {
  return { tecido:'tecidos', cor:'cores', material:'materiais', modelo:'modelos',
           colecao:'colecoes', grade:'grades', desenho:'desenhos',
           marca:'marcas', linha:'linhas', base:'bases', bloco:'blocos', equipe:'equipe', funcao:'funcoes', tarefa:'tarefas', etapa:'etapas', componente:'componentes' }[tipo];
}

async function salvarCadastro() {
  if (!exigirAdmin('criar ou editar cadastros')) return;
  const { tipo, editId } = cadastroContext;
  const list = pluralize(tipo);
  const v = id => document.getElementById(id)?.value || '';
  let item = editId ? STATE[list].find(x => x.id === editId) : { id: uid() };
  if (!item) item = { id: uid() };

  if (tipo === 'tecido') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.desc = v('m-desc');
    item.categoria = v('m-categoria');
    item.peso = parseFloat(String(v('m-peso')).replace(',', '.')) || 0;
  }
  else if (tipo === 'cor') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.hex = v('m-hex');
    item.codigo = v('m-codigo');
    item.siglaSku = (v('m-siglasku') || '').trim().toUpperCase();
    item.peso = parseFloat(String(v('m-cor-peso')).replace(',', '.')) || 0;
  }
  else if (tipo === 'material') {
    if (!v('m-codigo') || !v('m-desc')) return toast('Código e descrição obrigatórios', 'err');
    item.codigo = v('m-codigo');
    item.tipo = v('m-tipo');
    item.desc = v('m-desc');
  }
  else if (tipo === 'modelo') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.linha = v('m-linha');
    item.skuLinha = (v('m-skulinha') || '').trim().toUpperCase();
    item.categoria = v('m-categoria');
    item.baseId = v('m-vinc-base');
    item.marcaId = v('m-vinc-marca');
    item.designerId = v('m-vinc-designer');
    item.ftecId = v('m-vinc-ftec');
    item.coordId = v('m-vinc-coord');
  }
  else if (tipo === 'colecao') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.temporada = v('m-temp');
  }
  else if (tipo === 'grade') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.tipoPeca = v('m-grade-tipopeca');
    item.variacao = v('m-grade-variacao');
    item.tamanhos = {};
    ['p','m','g','gg','g1','g2','g3'].forEach(t => {
      item.tamanhos[t] = parseInt(v('m-gr-'+t)) || 0;
    });
    item.fases = Array.from(document.querySelectorAll('#m-fases-container .fase-grade-bloco')).map((b, i) => {
      const pb = parseBobinas(b.querySelector('.fase-bobinas')?.value);
      return {
        ordem: i + 1,
        nome: b.querySelector('.fase-nome')?.value || '',
        tecidoId: b.querySelector('.fase-tec')?.value || '',
        unidades: parseInt(b.querySelector('.fase-unid')?.value) || 2,
        comp: b.querySelector('.fase-comp')?.value || '',
        larg: b.querySelector('.fase-larg')?.value || '',
        bobinas: pb == null ? '' : pb
      };
    });
    // Retrocompatibilidade: usa a primeira fase para os campos legados
    const f1 = item.fases[0] || {};
    item.enfestoComprimento = f1.comp || '';
    item.enfestoLargura = f1.larg || '';
    // Remove a estrutura antiga "enfestos" (por categoria) que foi substituída pelas fases
    delete item.enfestos;
  }
  else if (tipo === 'desenho') {
    if (!v('m-codigo')) return toast('Código obrigatório', 'err');
    if (!v('m-img-data')) return toast('Imagem obrigatória', 'err');
    item.codigo = v('m-codigo');
    item.desc = v('m-desc');
    item.skuLinha = (v('m-desenho-sku') || '').trim().toUpperCase();
    const imgInput = v('m-img-data');
    if (imgInput.startsWith('data:image/')) {
      try {
        item.img = await uploadDesenhoImagem(imgInput);
      } catch (e) {
        console.error('Upload falhou', e);
        return toast('Falha ao enviar imagem para Storage — tente novamente', 'err');
      }
    } else {
      item.img = imgInput;
    }
    item.modeloId = v('m-vinc-modelo');
    item.baseId = v('m-vinc-base');
    item.colecaoId = v('m-vinc-colecao');
    item.marcaId = v('m-vinc-marca');
    item.linhaId = v('m-vinc-linha');
    item.blocoId = v('m-vinc-bloco');
    item.designerId = v('m-vinc-designer');
    item.coordId = v('m-vinc-coord');
    item.tecidoPadraoId = v('m-vinc-tecido');
    item.corPrincipalId = v('m-vinc-cor');
    item.corSecundariaId = v('m-vinc-cor2');
    item.corTerciariaId = v('m-vinc-cor3');
    // Componentes com tecido + cor + qtd/peça (estrutura nova)
    const componentesAntigos = Array.isArray(item.componentes) ? item.componentes : [];
    const componentesMarcados = Array.from(document.querySelectorAll('.m-componente-chk:checked')).map(chk => {
      const compId = chk.value;
      const cad = STATE.componentes.find(x => x.id === compId);
      const tecEl = document.querySelector(`.m-comp-tec[data-comp="${compId}"]`);
      const corEl = document.querySelector(`.m-comp-cor[data-comp="${compId}"]`);
      const qtdEl = document.querySelector(`.m-comp-qtd[data-comp="${compId}"]`);
      return {
        componenteId: compId,
        nome: cad?.nome || '',
        tecidoId: tecEl?.value || '',
        corId: corEl?.value || '',
        qtdPorPeca: parseFloat(qtdEl?.value) || 1
      };
    });
    // Preserva componentes do desenho que NÃO existem no cadastro global (ex.: as
    // variantes "Frente/Costa/Mangas PARTE 1/2/3" e "Viés" que se perderam quando o
    // cadastro global foi recriado com 15 itens). Sem isso, como o editor só lista os
    // componentes globais, o save apagaria silenciosamente esses componentes e suas
    // cores. Descarta os que já foram capturados por nome nas linhas marcadas.
    const idsGlobais = new Set(STATE.componentes.map(c => c.id));
    const nomesGlobais = new Set(STATE.componentes.map(c => _normNome(c.nome)));
    const nomesMarcados = new Set(componentesMarcados.map(c => _normNome(c.nome)));
    const componentesOrfaos = componentesAntigos.filter(c =>
      !idsGlobais.has(c.componenteId)
      && !nomesGlobais.has(_normNome(c.nome))
      && !nomesMarcados.has(_normNome(c.nome)));
    item.componentes = componentesMarcados.concat(componentesOrfaos);
    // Retrocompat: mantém componentesIds sincronizado
    item.componentesIds = item.componentes.map(c => c.componenteId);

    // Aviamentos com qtd/peça + aplicação (estrutura nova)
    item.aviamentos = Array.from(document.querySelectorAll('.m-aviamento-chk:checked')).map(chk => {
      const mId = chk.value;
      const qtdEl = document.querySelector(`.m-av-qtd[data-av="${mId}"]`);
      const appEl = document.querySelector(`.m-av-app[data-av="${mId}"]`);
      return {
        materialId: mId,
        qtdPorPeca: parseFloat(qtdEl?.value) || 1,
        aplicacao: appEl?.value || ''
      };
    });
    item.aviamentosIds = item.aviamentos.map(a => a.materialId);

    // Etapas padrão do desenho (na ordem visual das marcadas)
    item.etapasNomes = Array.from(document.querySelectorAll('#m-desenho-etapas .etapa-check'))
      .filter(l => l.querySelector('input:checked'))
      .map(l => l.querySelector('input').value);
  }
  else if (tipo === 'marca' || tipo === 'linha' || tipo === 'base' || tipo === 'bloco') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.desc = v('m-desc');
  }
  else if (tipo === 'equipe') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.funcao = v('m-funcao');
    // Vincula a função por ID também — assim se o nome da função for renomeado, reflete aqui
    const funcaoMatch = (STATE.funcoes || []).find(f => (f.nome || '').trim().toLowerCase() === (item.funcao || '').trim().toLowerCase());
    item.funcaoId = funcaoMatch?.id || '';
  }
  else if (tipo === 'funcao') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    const nomeAntigo = editId ? (item.nome || '') : '';
    const nomeNovo = v('m-nome');
    // Bloqueia em definitivo o cadastro de "Coordenador de produção
    // Enfestadeira/Esteira de corte" (decisão do admin).
    if (ehFuncaoCoordEnfestEsteira(nomeNovo)) {
      return toast('Esta função foi removida em definitivo. Use outro nome.', 'err');
    }
    item.nome = nomeNovo;
    item.desc = v('m-desc');
    item.acoes = v('m-acoes');
    // Se o nome mudou, propaga pra todas as pessoas da equipe que usavam o nome antigo
    if (editId && nomeAntigo && nomeAntigo !== nomeNovo) {
      let migradas = 0;
      (STATE.equipe || []).forEach(p => {
        if ((p.funcao || '').trim().toLowerCase() === nomeAntigo.trim().toLowerCase()) {
          p.funcao = nomeNovo;
          migradas++;
        }
      });
      if (migradas > 0) {
        await saveState('equipe');
        toast(`${migradas} pessoa(s) da equipe atualizada(s)`, 'ok');
      }
    }
  }
  else if (tipo === 'etapa') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.ordem = parseInt(v('m-ordem')) || 0;
    item.tarefas = Array.from(document.querySelectorAll('#m-tarefas-container .tarefa-etapa-bloco'))
      .map((b, i) => ({
        id: b.dataset.id || uid(),
        ordem: i + 1,
        nome: (b.querySelector('.tarefa-nome')?.value || '').trim(),
        desc: b.querySelector('.tarefa-desc')?.value || ''
      }))
      .filter(t => t.nome);
    // Limpa estrutura antiga (tarefasIds + STATE.tarefas) — agora tarefa vive dentro da etapa.
    delete item.tarefasIds;
  }
  else if (tipo === 'componente') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.desc = v('m-desc');
    item.tipoPeca = v('m-comp-tipopeca');
    item.variacao = v('m-comp-variacao');
    const nCores = item.variacao === 'tricolor' ? 3 : item.variacao === 'bicolor' ? 2 : item.variacao === 'basica' ? 1 : 0;
    item.cor1Id = nCores >= 1 ? v('m-comp-cor1') : '';
    item.cor2Id = nCores >= 2 ? v('m-comp-cor2') : '';
    item.cor3Id = nCores >= 3 ? v('m-comp-cor3') : '';
  }

  if (!editId) STATE[list].push(item);
  await saveState(list);

  closeModal('modal-cad');
  toast('Salvo com sucesso', 'ok');

  if (cadastroContext.origin === 'os-form') {
    refreshOSFormDropdowns();
    if (tipo === 'etapa') renderEtapas();
    if (!editId) {
      const autoMap = {
        marca: { id: 'f-griffe', field: 'nome' },
        colecao: { id: 'f-colecao', field: 'nome' },
        modelo: { id: 'f-modelo', field: 'nome' },
        linha: { id: 'f-linha', field: 'nome' },
        base: { id: 'f-base', field: 'nome' },
        bloco: { id: 'f-bloco', field: 'nome' },
        grade: { id: 'f-grade-preset', field: 'nome' },
        desenho: { id: 'f-desenho', field: 'codigo' }
      };
      const t = autoMap[tipo];
      if (t) {
        const el = document.getElementById(t.id);
        if (el) el.value = item[t.field] || '';
        if (tipo === 'desenho') sincCodigoDesenho('desenho');
      }
    }
  } else {
    goto('cad-' + list);
  }
}

function refreshOSFormDropdowns() {
  const IDS = ['f-colecao','f-modelo','f-desenho','f-grade-preset','f-griffe','f-linha','f-base','f-bloco','f-designer','f-ftec','f-coordenado'];
  const saved = {};
  IDS.forEach(id => { const el = document.getElementById(id); if (el) saved[id] = el.value; });
  fillSelect('f-colecao', STATE.colecoes, 'nome', '— selecione —');
  fillSelect('f-modelo', STATE.modelos, 'nome', '— selecione —');
  fillSelect('f-desenho', STATE.desenhos, 'codigo', '— selecione —', d => `${d.codigo}${d.desc ? ' · '+d.desc : ''}`);
  preencherDropdownGradesOS();
  fillSelect('f-griffe', STATE.marcas, 'nome', '— selecione —');
  fillSelect('f-linha', STATE.linhas, 'nome', '— selecione —');
  fillSelect('f-base', STATE.bases, 'nome', '— selecione —');
  fillSelect('f-bloco', STATE.blocos, 'nome', '— selecione —');
  fillSelect('f-designer', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));
  fillSelect('f-ftec', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));
  fillSelect('f-coordenado', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));
  atualizarDatalistCodigos();
  Object.entries(saved).forEach(([id, val]) => { const el = document.getElementById(id); if (el && val) el.value = val; });
}

async function excluirCadastro(tipo, id) {
  if (!exigirAdmin('excluir cadastros')) return;
  if (!confirm('Excluir este registro?')) return;
  const list = pluralize(tipo);
  STATE[list] = STATE[list].filter(x => x.id !== id);
  await saveState(list);
  // Tarefa excluida: limpa referencias em etapas.tarefasIds
  if (tipo === 'tarefa') {
    let mexeu = false;
    (STATE.etapas || []).forEach(e => {
      if (Array.isArray(e.tarefasIds) && e.tarefasIds.includes(id)) {
        e.tarefasIds = e.tarefasIds.filter(x => x !== id);
        mexeu = true;
      }
    });
    if (mexeu) await saveState('etapas');
  }
  toast('Excluído', 'ok');
  // Tarefa nao tem mais pagina propria — volta para cad-etapas (arvore).
  goto(tipo === 'tarefa' ? 'cad-etapas' : 'cad-' + list);
}

/* ========================================================= */
/*                       RENDER TABELAS                      */
/* ========================================================= */
function esc(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function acoesCell(tipo, id) {
  return `<td class="col-actions row-actions">
    <button class="edit" onclick="openCadastroModal('${tipo}','${id}')">editar</button>
    <button class="edit" onclick="duplicarCadastro('${tipo}','${id}')">duplicar</button>
    <button class="del" onclick="excluirCadastro('${tipo}','${id}')">excluir</button>
  </td>`;
}

async function duplicarCadastro(tipo, id) {
  if (!exigirAdmin('duplicar cadastros')) return;
  const list = pluralize(tipo);
  const original = STATE[list].find(x => x.id === id);
  if (!original) return toast('Cadastro não encontrado', 'err');
  const copia = JSON.parse(JSON.stringify(original));
  copia.id = uid();
  // Marca algum campo identificador com "(cópia)" para distinguir
  if (copia.nome)   copia.nome   = copia.nome + ' (cópia)';
  else if (copia.codigo) copia.codigo = copia.codigo + ' (cópia)';
  else if (copia.desc)   copia.desc   = copia.desc + ' (cópia)';
  STATE[list].push(copia);
  await saveState(list);
  toast('Cadastro duplicado', 'ok');
  // Re-renderiza a página atual
  const activeBtn = document.querySelector('.nav-btn.active');
  const pagina = activeBtn?.dataset.page || ('cad-' + list);
  goto(pagina);
}

function renderTecidos() {
  const tb = document.getElementById('tbl-tecidos');
  if (!STATE.tecidos.length) { tb.innerHTML = `<tr><td colspan="5" class="empty">Nenhum tecido cadastrado.</td></tr>`; return; }
  const catLabel = { malha: 'Malha algodão · máx 80', moletom: 'Moletom · máx 36', outro: 'Outro' };
  tb.innerHTML = STATE.tecidos.map(t => `
    <tr>
      <td><strong>${esc(t.nome)}</strong></td>
      <td>${esc(t.desc)}</td>
      <td><span class="badge">${esc(catLabel[t.categoria] || '—')}</span></td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${t.peso ? esc(t.peso) + ' g/m²' : '—'}</td>
      ${acoesCell('tecido', t.id)}
    </tr>`).join('');
}
function renderCores() {
  const tb = document.getElementById('tbl-cores');
  if (!STATE.cores.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhuma cor cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.cores.map(c => `
    <tr><td><span class="color-swatch" style="background:${esc(c.hex)}"></span><strong>${esc(c.nome)}</strong></td>
    <td><span class="badge">${esc(c.codigo)||'—'}</span></td>
    <td style="font-family:'IBM Plex Mono',monospace;">${c.peso ? esc(c.peso)+' g/m²' : '—'}</td>${acoesCell('cor', c.id)}</tr>`).join('');
}
function renderMateriais() {
  const tb = document.getElementById('tbl-materiais');
  if (!STATE.materiais.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhum material cadastrado.</td></tr>`; return; }
  tb.innerHTML = STATE.materiais.map(m => `
    <tr><td><span class="badge">${esc(m.codigo)}</span></td><td>${esc(m.desc)}</td>
    <td>${esc(m.tipo)||'—'}</td>${acoesCell('material', m.id)}</tr>`).join('');
}
/* ========================================================= */
/*                ESTOQUE DE TECIDOS (kg)                     */
/* ========================================================= */
// Converte as compras vindas da Contabilidade (compras_materiais) em
// movimentos de ENTRADA, no mesmo formato de STATE.estoqueMov.
// A cor sai crua daqui (o de-para da Contabilidade manda a cor pura, "Preto");
// quem canonicaliza para o nome desdobrado por tecido é movimentacoesEstoque,
// num ponto só, junto com o resto do razão.
function comprasComoMovimentos() {
  return (comprasCache || []).map(c => ({
    id: 'nf_' + c.id,
    tipo: 'entrada',
    tecidoNome: c.tecido_nome || '',
    corNome: c.cor_nome || '',
    kg: parseFloat(c.quantidade_kg) || 0,
    data: (c.data || '').slice(0, 10),
    origem: 'nf',
    osId: '',
    osNumero: c.nota_fiscal || '',
    obs: c.fornecedor || ''
  }));
}

// Todos os movimentos do estoque: entradas/saídas locais (estoqueMov) +
// compras da Contabilidade (entradas via NF). Fonte única para saldo e histórico.
//
// A cor de TODO movimento passa por corCanonicaPorTecido antes de sair daqui.
// Motivo: as cores foram desdobradas por tecido ("Preto" virou "Preto Malha
// Algodão", "Preto Ribana Moletom", …), mas o razão já gravado — entradas
// manuais, baixas de OSs antigas, compras por NF — guarda a cor pura "Preto".
// Como a chave do saldo é tecido||cor, sem converter aqui o mesmo tecido
// apareceria em DUAS linhas: uma com o histórico e outra com os lançamentos
// novos. A conversão é em tempo de leitura: não reescreve nada no banco, e
// desfazer é só reverter o código.
function movimentacoesEstoque() {
  return [...(STATE.estoqueMov || []), ...comprasComoMovimentos()]
    .map(m => {
      const canon = corCanonicaPorTecido(m.corNome || '', m.tecidoNome || '');
      return canon === (m.corNome || '') ? m : { ...m, corNome: canon };
    });
}

// Calcula, por tecido+cor:
//   entrada   = compras (NF) + entradas manuais
//   reservado = consumo de OSs salvas mas AINDA NÃO baixadas (status reservado)
//   saida     = baixa definitiva: OSs apontadas como produzidas + saídas manuais
//   disponivel (livre) = entrada − reservado − saida
function calcularSaldosEstoque() {
  const key = (t, c) => _normNome(t) + '||' + _normNome(c);
  const detMap = new Map();
  movimentacoesEstoque().forEach(m => {
    const tNome = m.tecidoNome || '', cNome = m.corNome || '';
    const k = key(tNome, cNome);
    const cur = detMap.get(k) || { tecidoNome: tNome, corNome: cNome, entrada: 0, reservado: 0, saida: 0, fechados: 0, abertos: 0 };
    const kg = parseFloat(m.kg) || 0;
    if (m.tipo === 'entrada') cur.entrada += kg;
    else if (m.origem === 'os' && m.status !== 'consumido') cur.reservado += kg;
    else cur.saida += kg;  // OS já baixada (consumido) + saídas manuais
    // Unidades (contagem física do lançamento manual): entrada soma, saída subtrai.
    const fch = parseInt(m.fechados) || 0, abr = parseInt(m.abertos) || 0;
    const sinal = m.tipo === 'entrada' ? 1 : -1;
    cur.fechados += sinal * fch;
    cur.abertos += sinal * abr;
    if (!cur.tecidoNome && tNome) cur.tecidoNome = tNome;
    if (!cur.corNome && cNome) cur.corNome = cNome;
    detMap.set(k, cur);
  });
  const detalhe = Array.from(detMap.values())
    .map(c => ({ ...c, disponivel: c.entrada - c.reservado - c.saida }))
    .sort((a, b) => (a.tecidoNome || '').localeCompare(b.tecidoNome || '') || (a.corNome || '').localeCompare(b.corNome || ''));
  return { detalhe };
}

// Agrupa os movimentos de OS por OS (para a seção "apontar OS"): total kg e status.
function osComMaterialReservado() {
  const map = new Map();
  (STATE.estoqueMov || []).forEach(m => {
    if (m.origem !== 'os') return;
    const cur = map.get(m.osId) || { osId: m.osId, osNumero: m.osNumero || '', kg: 0, consumido: true };
    cur.kg += parseFloat(m.kg) || 0;
    // OS é "consumida" só se TODOS os movimentos dela estiverem consumidos.
    if (m.status !== 'consumido') cur.consumido = false;
    map.set(m.osId, cur);
  });
  return Array.from(map.values()).map(o => {
    const os = (STATE.ordens || []).find(x => x.id === o.osId);
    return { ...o, modelo: os?.modeloNome || '', data: os?.data || '' };
  }).sort((a, b) => String(b.osNumero).localeCompare(String(a.osNumero), undefined, { numeric: true }));
}

function renderEstoque() {
  const cont = document.getElementById('estoque-painel');
  if (!cont) return;
  const { detalhe } = calcularSaldosEstoque();
  const fmt = n => Number(n || 0).toFixed(3).replace('.', ',');
  const dispCell = s => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;font-weight:700;color:${s < 0 ? '#c0392b' : 'inherit'};">${fmt(s)} kg</td>`;
  const semNada = !movimentacoesEstoque().length;
  // Tecido + cor são UMA categoria combinada. As variações de um mesmo tecido
  // ficam agrupadas e ordenadas juntas, com subtotal por tipo de tecido.
  const grupos = new Map();
  detalhe.forEach(c => {
    const k = _normNome(c.tecidoNome);
    const g = grupos.get(k) || { tecidoNome: c.tecidoNome || '(sem tecido)', entrada: 0, reservado: 0, saida: 0, fechados: 0, abertos: 0, linhas: [] };
    g.entrada += c.entrada; g.reservado += c.reservado; g.saida += c.saida;
    g.fechados += c.fechados || 0; g.abertos += c.abertos || 0; g.linhas.push(c);
    grupos.set(k, g);
  });
  const gruposArr = Array.from(grupos.values()).sort((a, b) => (a.tecidoNome || '').localeCompare(b.tecidoNome || ''));
  gruposArr.forEach(g => g.linhas.sort((a, b) => (a.corNome || '').localeCompare(b.corNome || '')));

  // A linha já mostra o tecido antes do "·", então o sufixo do tecido no nome da
  // cor ("Preto Malha Algodão") sai — evita "Malha Algodão · Preto Malha Algodão".
  const corLabel = (nome, tecido) => esc(corSemTecido(nome, tecido)) || '<span style="color:var(--ink-2)">(sem cor)</span>';
  const numCell = (n, bold) => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;${bold ? 'font-weight:700;' : ''}">${fmt(n)}</td>`;
  // Célula de UNIDADES (inteiro, sem kg). Fundo levemente diferente p/ destacar.
  const uniCell = (n, bold) => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;${bold ? 'font-weight:700;' : ''}">${Number(n) || 0}</td>`;
  // kg: Entradas | Reservado | Saídas | Disponível ; unidades: Fechados | Abertos
  const cellsVals = (o, bold) =>
    numCell(o.entrada, bold) + numCell(o.reservado, bold) + numCell(o.saida, bold) +
    dispCell(o.entrada - o.reservado - o.saida) +
    uniCell(o.fechados, bold) + uniCell(o.abertos, bold);
  const linhasEstoque = gruposArr.map(g => {
    const cores = g.linhas.map(c => `
      <tr>
        <td>${esc(g.tecidoNome)} · <strong>${corLabel(c.corNome, g.tecidoNome)}</strong></td>
        ${cellsVals(c, false)}
      </tr>`).join('');
    // Subtotal do tipo de tecido (só quando há mais de uma cor no grupo).
    const subtotal = g.linhas.length > 1 ? `
      <tr style="background:#eef6f0;">
        <td style="text-align:right;font-weight:700;color:var(--ink-2);">Subtotal ${esc(g.tecidoNome)}</td>
        ${cellsVals(g, true)}
      </tr>` : '';
    return cores + subtotal;
  }).join('');

  const estoqueHtml = `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">Estoque por tecido + cor</h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">
        Colunas em <b>kg</b>: Entradas, Reservado (OSs não produzidas), Saídas (baixa definitiva),
        Disponível (= Entradas − Reservado − Saídas). Colunas em <b>unidades</b> (lançamento manual):
        <b>Fechados</b> (rolos/peças lacrados) e <b>Abertos</b> (em uso).
      </div>
      <table class="table">
        <thead><tr>
          <th>Tecido + cor</th>
          <th style="text-align:right;">Entradas</th><th style="text-align:right;">Reservado</th>
          <th style="text-align:right;">Saídas</th><th style="text-align:right;">Disponível</th>
          <th style="text-align:right;">Fechados (un)</th>
          <th style="text-align:right;">Abertos (un)</th>
        </tr></thead>
        <tbody>
          ${gruposArr.length ? linhasEstoque : `<tr><td colspan="7" class="empty">Sem movimentações ainda.</td></tr>`}
        </tbody>
      </table>
    </div>`;

  // Apontar OS produzida → converte a RESERVA em SAÍDA definitiva.
  const osMat = osComMaterialReservado().filter(o => o.kg > 0);
  const reservadas = osMat.filter(o => !o.consumido);
  const baixadas = osMat.filter(o => o.consumido);
  const linhaOS = (o, baixada) => `
    <tr>
      <td><strong>${esc(o.osNumero) || '—'}</strong></td>
      <td>${esc(o.modelo) || '—'}</td>
      <td style="white-space:nowrap;">${esc(formatDate(o.data))}</td>
      <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(o.kg)} kg</td>
      <td>${baixada ? '<span class="badge" style="background:#f6dcda;">Baixado</span>' : '<span class="badge" style="background:#fde9c8;">Reservado</span>'}</td>
      <td class="col-actions row-actions">${baixada
        ? `<button onclick="estornarBaixaMaterialOS('${esc(o.osId)}')">estornar</button>`
        : `<button onclick="darBaixaMaterialOS('${esc(o.osId)}')">dar baixa</button>`}</td>
    </tr>`;
  const apontarHtml = osMat.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">OSs · baixa de material</h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">
        Aponte a OS como <b>produzida</b> ("dar baixa") para converter a reserva em
        <b>saída definitiva</b> do estoque. Use "estornar" para desfazer.
      </div>
      <table class="table">
        <thead><tr><th>OS</th><th>Modelo</th><th>Data</th><th style="text-align:right;">Material</th><th>Situação</th><th class="col-actions">Ação</th></tr></thead>
        <tbody>
          ${reservadas.map(o => linhaOS(o, false)).join('')}
          ${baixadas.map(o => linhaOS(o, true)).join('')}
        </tbody>
      </table>
    </div>` : '';

  const movs = movimentacoesEstoque().slice()
    .sort((a, b) => String(b.data || '').localeCompare(String(a.data || '')) || String(b.id).localeCompare(String(a.id)))
    .slice(0, 60);
  const origemLabel = m => m.origem === 'os' ? `OS ${esc(m.osNumero || '')}`
    : m.origem === 'nf' ? `NF ${esc(m.osNumero || '')}${m.obs ? ' · ' + esc(m.obs) : ''}`
    : 'Manual';
  const movHtml = `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">Movimentações recentes</h2>
      <table class="table">
        <thead><tr><th>Data</th><th>Tipo</th><th>Tecido</th><th>Cor</th><th style="text-align:right;">Qtd (kg)</th><th style="text-align:right;">Fech.</th><th style="text-align:right;">Abertos</th><th>Origem</th><th class="col-actions">Ações</th></tr></thead>
        <tbody>
          ${movs.length ? movs.map(m => `
            <tr>
              <td style="white-space:nowrap;">${esc(m.data) || '—'}</td>
              <td>${m.tipo === 'entrada'
                ? '<span class="badge" style="background:#d6f0db;">Entrada</span>'
                : (m.origem === 'os'
                    ? (m.status === 'consumido'
                        ? '<span class="badge" style="background:#f6dcda;">Saída (OS)</span>'
                        : '<span class="badge" style="background:#fde9c8;">Reserva</span>')
                    : '<span class="badge" style="background:#f6dcda;">Saída</span>')}</td>
              <td>${esc(m.tecidoNome) || '—'}</td>
              <td>${esc(m.corNome) || '—'}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(m.kg)} kg</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${m.fechados ? Number(m.fechados) : '—'}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${m.abertos ? Number(m.abertos) : '—'}</td>
              <td>${origemLabel(m)}</td>
              <td class="col-actions row-actions">${m.origem === 'manual' ? `<button onclick="excluirMovEstoque('${esc(m.id)}')">excluir</button>` : '<span style="color:var(--ink-2);font-size:11px;">auto</span>'}</td>
            </tr>`).join('') : `<tr><td colspan="9" class="empty">Nenhuma movimentação.</td></tr>`}
        </tbody>
      </table>
    </div>`;

  cont.innerHTML = `
    ${semNada ? `<div class="info-box">Ainda não há movimentações. As <b>entradas</b> vêm das compras lançadas no programa de Contabilidade (por NF) ou de um lançamento manual aqui; as <b>saídas</b> entram sozinhas ao salvar uma OS com enfesto e o tecido com peso (g/m²) cadastrado.</div>` : ''}
    ${estoqueHtml}
    ${apontarHtml}
    ${movHtml}
  `;
}

let movEstoqueTipo = 'entrada';
function abrirMovEstoque(tipo) {
  if (!exigirAdmin('movimentar estoque')) return;
  movEstoqueTipo = tipo === 'saida' ? 'saida' : 'entrada';
  const title = document.getElementById('modal-estoque-title');
  const box = document.getElementById('modal-estoque-fields');
  title.textContent = movEstoqueTipo === 'entrada' ? 'Entrada de estoque (compra)' : 'Saída / ajuste manual';
  const tecOpts = '<option value="">— selecione —</option>' + (STATE.tecidos || []).map(t => `<option value="${esc(t.nome)}">${esc(t.nome)}</option>`).join('');
  const corOpts = '<option value="">— sem cor —</option>' + (STATE.cores || []).map(c => `<option value="${esc(c.nome)}">${esc(c.nome)}</option>`).join('');
  const hoje = new Date().toISOString().slice(0, 10);
  box.innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Tecido *</label><select id="me-tecido">${tecOpts}</select></div>
      <div class="field"><label>Cor</label><select id="me-cor">${corOpts}</select></div>
      <div class="field"><label>Quantidade (kg) *</label><input type="number" min="0" step="0.001" id="me-kg" placeholder="Ex.: 50,000"></div>
      <div class="field"><label>Data</label><input type="date" id="me-data" value="${hoje}"></div>
      <div class="field"><label>Itens fechados (un)</label><input type="number" min="0" step="1" id="me-fechados" placeholder="0"></div>
      <div class="field"><label>Itens abertos em uso (un)</label><input type="number" min="0" step="1" id="me-abertos" placeholder="0"></div>
      <div class="field full"><label>Observação</label><input type="text" id="me-obs" placeholder="Ex.: NF 1234 / fornecedor"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">O kg é o equivalente em peso. As unidades (fechados = rolos/peças lacrados; abertos = em uso) são contagem física e aparecem em colunas próprias no painel.</div>
    ${movEstoqueTipo === 'saida' ? '<div class="info-box" style="margin-top:8px;">Use para corrigir o estoque (perdas, sobras, inventário). O consumo de produção já é lançado sozinho ao salvar a OS.</div>' : ''}`;
  openModal('modal-estoque');
}

async function salvarMovEstoque() {
  if (!exigirAdmin('movimentar estoque')) return;
  const v = id => document.getElementById(id)?.value || '';
  const tecidoNome = v('me-tecido');
  if (!tecidoNome) return toast('Selecione o tecido', 'err');
  const kg = parseFloat(String(v('me-kg')).replace(',', '.')) || 0;
  if (!(kg > 0)) return toast('Informe a quantidade em kg', 'err');
  const fechados = parseInt(v('me-fechados')) || 0;
  const abertos = parseInt(v('me-abertos')) || 0;
  if (!Array.isArray(STATE.estoqueMov)) STATE.estoqueMov = [];
  STATE.estoqueMov.push({
    id: uid(),
    tipo: movEstoqueTipo,
    tecidoNome,
    corNome: v('me-cor'),
    kg: Math.round(kg * 1000) / 1000,
    fechados,
    abertos,
    data: v('me-data') || new Date().toISOString().slice(0, 10),
    origem: 'manual',
    osId: '',
    osNumero: '',
    obs: v('me-obs')
  });
  await saveState('estoqueMov');
  closeModal('modal-estoque');
  toast(movEstoqueTipo === 'entrada' ? 'Entrada registrada' : 'Saída registrada', 'ok');
  renderEstoque();
}

async function excluirMovEstoque(id) {
  if (!exigirAdmin('excluir movimentação')) return;
  const m = (STATE.estoqueMov || []).find(x => x.id === id);
  if (!m) return;
  if (m.origem !== 'manual') return toast('Saídas automáticas de OS são removidas ao excluir a própria OS', 'err');
  if (!confirm('Excluir esta movimentação?')) return;
  STATE.estoqueMov = STATE.estoqueMov.filter(x => x.id !== id);
  await saveState('estoqueMov');
  toast('Movimentação excluída', 'ok');
  renderEstoque();
}

/* ========================================================= */
/*           ESTOQUE DE CORTE (peças cortadas)               */
/* ========================================================= */
// Peças já cortadas, em estoque esperando a costura. Diferente do estoque de
// tecidos (kg), aqui a unidade é PEÇA (componente cortado). Entradas e saídas
// são DERIVADAS das OS em tempo de render; só os ajustes manuais persistem em
// STATE.corteMov. Saldo por tecido+cor:
//   entrada  = soma dos componentes de TODAS as OS (cada OS = um pacote)
//   saida    = idem, mas só das OS com a etapa "Costura" marcada
//   contagem = líquido dos lançamentos manuais (entrada − saída)
//   estoque  = entrada − saida + contagem

// A OS tem uma etapa (casada por regex) marcada no checklist? Genérico — usado
// pelos gatilhos automáticos de saída entre os campos de estoque.
function osEtapaMarcada(o, re) {
  const checks = (o.progresso && o.progresso.etapasCheck) || {};
  const nome = (o.etapas || []).find(n => re.test(n));
  return nome ? !!checks[nome] : false;
}
// "Costura" marcada → gatilho da saída do Estoque de corte (entra em Costurando).
function osCosturaMarcada(o) { return osEtapaMarcada(o, /costura/i); }
// "Limpeza de fios" (ou "Retirada de fios") marcada → gatilho da saída de Costurando.
function osFiosMarcada(o) { return osEtapaMarcada(o, /fios/i); }

// Componentes de uma OS agregados por tecido(material)+cor → unidades cortadas.
function componentesPorTecidoCorOS(o) {
  const mapa = new Map();
  (o.componentes || []).forEach(c => {
    const qtd = Number(c.qtdTotal) || 0;
    if (!(qtd > 0)) return;
    const tecidoNome = c.materialNome || '';
    // Mesma convergência do razão de kg: OSs antigas gravaram a cor pura
    // ("Preto") nos componentes, as novas gravam a desdobrada por tecido. Sem
    // canonicalizar, o Estoque de corte mostraria o mesmo tecido em duas linhas.
    const corNome = corCanonicaPorTecido(c.corNome || '', tecidoNome);
    const k = _normNome(tecidoNome) + '||' + _normNome(corNome);
    const cur = mapa.get(k) || { tecidoNome, corNome, qtd: 0 };
    cur.qtd += qtd;
    mapa.set(k, cur);
  });
  return Array.from(mapa.values());
}

// Campos de estoque em processo, um por ETAPA de produção que tem campo. Modelo
// SOBREPOSTO (cumulativo): o volume de cada OS fica SEMPRE no campo da etapa
// marcada por ÚLTIMO (faseAtualOS = maior etapasSeq). Marcar uma nova etapa move
// o volume para o campo dela; etapas SEM campo (Acabamento de mangas, Ensaque,
// Estampa, Lavanderia…) NÃO movem o volume. Para adicionar uma fase nova, inserir
// uma linha aqui (+ a chave do array manual em STATE/keys + nav/section/rota).
//   entrada.tipo 'etapa' = OS com a etapa (re/label) marcada no checklist.
const FASES_ESTOQUE = [
  { id: 'corte',      titulo: 'Estoque de corte', movKey: 'corteMov',      painelId: 'corte-painel', semContagem: true,
    entrada: { tipo: 'etapa', re: /corte/i, label: 'Corte' } },
  { id: 'costurando', titulo: 'Costurando',       movKey: 'costurandoMov', painelId: 'costurando-painel', semContagem: true, osTodasEntradas: true,
    entrada: { tipo: 'etapa', re: /costura/i, label: 'Costura' } },
  { id: 'fios',       titulo: 'Retirada de fios', movKey: 'fiosMov',       painelId: 'fios-painel', semContagem: true,
    entrada: { tipo: 'etapa', re: /fios/i, label: 'Retirada de fios' } },
  { id: 'expedicao',  titulo: 'Expedição',        movKey: 'expedicaoMov',  painelId: 'expedicao-painel', semContagem: true,
    entrada: { tipo: 'etapa', re: /expedi/i, label: 'Expedição' } },
];

// A OS entrou nesta fase? (etapa da fase marcada no checklist).
function _faseEntrouOS(o, entrada) {
  if (!entrada) return false;
  if (entrada.tipo === 'oscriada') return true;
  return osEtapaMarcada(o, entrada.re);
}

// Nome (no checklist da OS) da etapa que dispara esta fase, p/ ler o etapasSeq.
function _nomeEtapaDaFase(o, fase) {
  if (!fase || !fase.entrada || fase.entrada.tipo === 'oscriada') return null;
  return (o.etapas || []).find(n => fase.entrada.re.test(n)) || null;
}

// Saldo de uma fase por tecido+cor:
//   entrada  = OSs que entraram na fase (× componentes)
//   saida    = OSs que já entraram na PRÓXIMA fase
//   contagem = líquido dos lançamentos manuais da fase (STATE[movKey])
//   estoque  = entrada − saida + contagem
function calcularSaldosFase(idx) {
  const fase = FASES_ESTOQUE[idx];
  const key = (t, c) => _normNome(t) + '||' + _normNome(c);
  const map = new Map();
  const pegar = (tNome, cNome) => {
    const k = key(tNome, cNome);
    let cur = map.get(k);
    if (!cur) { cur = { tecidoNome: tNome, corNome: cNome, entrada: 0, saida: 0, contagem: 0, osNums: new Set() }; map.set(k, cur); }
    if (!cur.tecidoNome && tNome) cur.tecidoNome = tNome;
    if (!cur.corNome && cNome) cur.corNome = cNome;
    return cur;
  };
  (STATE.ordens || []).forEach(o => {
    if (!_faseEntrouOS(o, fase.entrada)) return;
    // Modelo sobreposto: a OS "saiu" desta fase se o volume está em OUTRA fase
    // agora (a última etapa marcada não é a desta fase).
    const atual = faseAtualOS(o);
    const saiu = atual !== idx;
    // Quais OSs listar na coluna OS desta linha:
    //  - padrão: só as que estão ATUALMENTE nesta fase (compõem o saldo).
    //  - fase.osTodasEntradas: TODAS as que entraram na fase (etapa marcada),
    //    mesmo que já tenham avançado (ex.: Costurando lista toda OS com Costura).
    const listarOS = fase.osTodasEntradas ? true : (atual === idx);
    const numOS = (o.os || '').toString().trim();
    componentesPorTecidoCorOS(o).forEach(it => {
      const cur = pegar(it.tecidoNome, it.corNome);
      cur.entrada += it.qtd;
      if (saiu) cur.saida += it.qtd;
      if (listarOS && numOS) cur.osNums.add(numOS);
    });
  });
  (STATE[fase.movKey] || []).forEach(m => {
    const cur = pegar(m.tecidoNome || '', m.corNome || '');
    const q = Number(m.qtd) || 0;
    cur.contagem += (m.tipo === 'entrada' ? q : -q);
  });
  const ordOS = arr => arr.slice().sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const detalhe = Array.from(map.values())
    .map(c => ({ ...c, estoque: c.entrada - c.saida + c.contagem, osList: ordOS(Array.from(c.osNums)) }))
    .sort((a, b) => (a.tecidoNome || '').localeCompare(b.tecidoNome || '') || (a.corNome || '').localeCompare(b.corNome || ''));
  return { detalhe };
}

// Etapa TERMINAL: ao marcar "Estoque", a OS sai do fluxo em processo (foi para o
// estoque de produtos acabados) — some de todos os campos (corte..expedição).
const TERMINAL_ETAPA_RE = /estoque/i;

// Fase atual de uma OS = o campo da etapa marcada por ÚLTIMO (modelo sobreposto).
// Usa etapasSeq (carimbo de quando cada etapa foi marcada): vence o maior seq
// entre as etapas com campo + a terminal. Sem seq (OS antiga), cai no canônico =
// última (na ordem do fluxo, terminal por último) que está marcada. Etapas sem
// campo (Acabamento de mangas, Ensaque…) não contam. Retorna o índice da fase, ou
// -1 quando a OS está FORA do fluxo (terminal "Estoque", ou nenhuma etapa de fase).
function faseAtualOS(o) {
  const seqs = (o.progresso && o.progresso.etapasSeq) || {};
  let achouSeq = false, idxSeq = -1, melhorSeq = -Infinity; // por etapasSeq
  let idxOrd = -1, melhorOrd = -1;                           // fallback canônico
  const considerar = (idx, ord, nome) => {
    const s = (nome && seqs[nome] != null) ? Number(seqs[nome]) : null;
    if (s != null && (!achouSeq || s > melhorSeq)) { achouSeq = true; melhorSeq = s; idxSeq = idx; }
    if (ord > melhorOrd) { melhorOrd = ord; idxOrd = idx; }
  };
  FASES_ESTOQUE.forEach((f, i) => {
    if (!_faseEntrouOS(o, f.entrada)) return;
    considerar(i, i, _nomeEtapaDaFase(o, f));
  });
  if (osEtapaMarcada(o, TERMINAL_ETAPA_RE)) {
    const nomeT = (o.etapas || []).find(n => TERMINAL_ETAPA_RE.test(n));
    considerar(-1, FASES_ESTOQUE.length, nomeT); // -1 = terminal; canônico = depois de todas
  }
  return achouSeq ? idxSeq : idxOrd;
}

// Cada fase do fluxo é um CAMPO próprio no menu (Estoque de corte, Costurando,
// Retirada de fios, Expedição). Renderiza UMA fase no seu painel: saldo tec+cor, OSs
// atualmente nessa fase e os lançamentos manuais da fase.
function renderFasePainel(faseIdx) {
  const fase = FASES_ESTOQUE[faseIdx];
  if (!fase) return;
  const cont = document.getElementById(fase.painelId);
  if (!cont) return;
  const fmt = n => (Number(n) || 0).toLocaleString('pt-BR');
  const fmtSinal = n => { const v = Number(n) || 0; return (v > 0 ? '+' : '') + v.toLocaleString('pt-BR'); };
  // A linha já mostra o tecido antes do "·", então o sufixo do tecido no nome da
  // cor ("Preto Malha Algodão") sai — evita "Malha Algodão · Preto Malha Algodão".
  const corLabel = (nome, tecido) => esc(corSemTecido(nome, tecido)) || '<span style="color:var(--ink-2)">(sem cor)</span>';
  const numCell = (n, bold) => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;${bold ? 'font-weight:700;' : ''}">${fmt(n)}</td>`;
  const contCell = (n, bold) => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;${bold ? 'font-weight:700;' : ''}">${n ? fmtSinal(n) : '—'}</td>`;
  const estCell = (n) => `<td style="text-align:right;font-family:'IBM Plex Mono',monospace;font-weight:700;color:${n < 0 ? '#c0392b' : 'inherit'};">${fmt(n)}</td>`;
  const mostrarCont = !fase.semContagem;
  const cellsVals = (o, bold) => numCell(o.entrada, bold) + numCell(o.saida, bold) + (mostrarCont ? contCell(o.contagem, bold) : '') + estCell(o.estoque);
  const ordOS = arr => (arr || []).slice().sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const osCell = arr => `<td style="font-family:'IBM Plex Mono',monospace;font-size:11px;">${(arr && arr.length) ? ordOS(arr).map(esc).join(', ') : '—'}</td>`;

  const { detalhe } = calcularSaldosFase(faseIdx);
  const grupos = new Map();
  detalhe.forEach(c => {
    const k = _normNome(c.tecidoNome);
    const g = grupos.get(k) || { tecidoNome: c.tecidoNome || '(sem tecido)', entrada: 0, saida: 0, contagem: 0, estoque: 0, linhas: [], osSet: new Set() };
    g.entrada += c.entrada; g.saida += c.saida; g.contagem += c.contagem; g.estoque += c.estoque;
    (c.osList || []).forEach(n => g.osSet.add(n));
    g.linhas.push(c); grupos.set(k, g);
  });
  const gruposArr = Array.from(grupos.values()).sort((a, b) => (a.tecidoNome || '').localeCompare(b.tecidoNome || ''));
  gruposArr.forEach(g => g.linhas.sort((a, b) => (a.corNome || '').localeCompare(b.corNome || '')));
  const linhas = gruposArr.map(g => {
    const cores = g.linhas.map(c => `<tr><td>${esc(g.tecidoNome)} · <strong>${corLabel(c.corNome, g.tecidoNome)}</strong></td>${cellsVals(c, false)}${osCell(c.osList)}</tr>`).join('');
    const sub = g.linhas.length > 1
      ? `<tr style="background:#eef6f0;"><td style="text-align:right;font-weight:700;color:var(--ink-2);">Subtotal ${esc(g.tecidoNome)}</td>${cellsVals(g, true)}${osCell(Array.from(g.osSet))}</tr>`
      : '';
    return cores + sub;
  }).join('');
  const entradaDesc = `OS com a etapa <b>${esc(fase.entrada.label)}</b> marcada`;
  const saidaDesc = 'OS cujo volume já foi para outro campo (uma etapa posterior virou a última marcada)';
  const card = `
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:8px;">
        <h2 style="margin:0;font-size:14px;">${esc(fase.titulo)} — por tecido + cor</h2>
        <div class="admin-only" style="display:flex;gap:6px;">
          <button class="btn primary" onclick="abrirMovFase('${fase.id}','entrada')">+ Entrada</button>
          <button class="btn" onclick="abrirMovFase('${fase.id}','saida')">− Saída / ajuste</button>
        </div>
      </div>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">
        Em <b>peças</b>: <b>Entradas</b> (${entradaDesc}), <b>Saídas</b> (${saidaDesc}),
        ${mostrarCont ? '<b>Contagem de estoque</b> (lançamentos manuais) e <b>Estoque</b> (= Entradas − Saídas + Contagem).' : 'e <b>Estoque</b> (= Entradas − Saídas, ajustado por lançamentos manuais).'}
        <b>OS</b> = números das OS que estão nesta fase agora (várias separadas por vírgula).
      </div>
      <table class="table">
        <thead><tr>
          <th>Tecido + cor</th>
          <th style="text-align:right;">Entradas</th>
          <th style="text-align:right;">Saídas</th>
          ${mostrarCont ? '<th style="text-align:right;">Contagem de estoque</th>' : ''}
          <th style="text-align:right;">Estoque</th>
          <th>OS</th>
        </tr></thead>
        <tbody>${gruposArr.length ? linhas : `<tr><td colspan="${mostrarCont ? 6 : 5}" class="empty">Sem peças nesta fase.</td></tr>`}</tbody>
      </table>
    </div>`;

  // OSs atualmente NESTA fase.
  const pacotes = (STATE.ordens || []).map(o => {
    const total = componentesPorTecidoCorOS(o).reduce((s, it) => s + it.qtd, 0);
    return { osId: o.id, osNumero: o.os || '', modelo: o.modeloNome || '', data: o.data || '', total, faseIdx: faseAtualOS(o) };
  }).filter(p => p.total > 0 && p.faseIdx === faseIdx)
    .sort((a, b) => String(b.osNumero).localeCompare(String(a.osNumero), undefined, { numeric: true }));
  const pacotesHtml = pacotes.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">OSs atualmente em ${esc(fase.titulo)}</h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Cada OS avança de fase automaticamente conforme as etapas do checklist são marcadas.</div>
      <table class="table">
        <thead><tr><th>OS</th><th>Modelo</th><th>Data</th><th style="text-align:right;">Peças</th><th class="col-actions">Ação</th></tr></thead>
        <tbody>
          ${pacotes.map(p => `
            <tr>
              <td><strong>${esc(p.osNumero) || '—'}</strong></td>
              <td>${esc(p.modelo) || '—'}</td>
              <td style="white-space:nowrap;">${esc(formatDate(p.data))}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(p.total)} pç</td>
              <td class="col-actions row-actions"><button onclick="verOS('${esc(p.osId)}')">ver OS</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';

  // Lançamentos manuais DESTA fase.
  const movs = (STATE[fase.movKey] || []).slice()
    .sort((a, b) => String(b.data || '').localeCompare(String(a.data || '')) || String(b.id).localeCompare(String(a.id)))
    .slice(0, 60);
  const movHtml = movs.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">Lançamentos manuais recentes — ${esc(fase.titulo)}</h2>
      <table class="table">
        <thead><tr><th>Data</th><th>Tipo</th><th>Tecido</th><th>Cor</th><th style="text-align:right;">Qtd (pç)</th><th>Obs.</th><th class="col-actions">Ações</th></tr></thead>
        <tbody>
          ${movs.map(m => `
            <tr>
              <td style="white-space:nowrap;">${esc(m.data) || '—'}</td>
              <td>${m.tipo === 'entrada'
                ? '<span class="badge" style="background:#d6f0db;">Entrada</span>'
                : '<span class="badge" style="background:#f6dcda;">Saída</span>'}</td>
              <td>${esc(m.tecidoNome) || '—'}</td>
              <td>${esc(m.corNome) || '—'}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(m.qtd)}</td>
              <td>${esc(m.obs) || '—'}</td>
              <td class="col-actions row-actions"><button onclick="excluirMovFase('${esc(fase.id)}', '${esc(m.id)}')">excluir</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';

  const vazio = !gruposArr.length && !movs.length && !pacotes.length;
  cont.innerHTML = `
    ${vazio ? `<div class="info-box">Sem peças nesta fase ainda. O volume entra sozinho conforme a etapa correspondente é marcada no checklist da OS. Use os botões para contagem física e ajustes manuais.</div>` : ''}
    ${card}
    ${pacotesHtml}
    ${movHtml}
  `;
}

// Re-renderiza o painel de uma fase pelo id (após salvar/excluir lançamento).
function renderFasePorId(faseId) {
  const idx = FASES_ESTOQUE.findIndex(f => f.id === faseId);
  if (idx >= 0) renderFasePainel(idx);
}

// Compat: "Estoque de corte" = primeira fase.
function renderEstoqueCorte() { renderFasePainel(0); }

// Lançamento manual genérico de qualquer fase do fluxo (entra na coluna
// "Contagem de estoque" daquela fase).
let movFaseTipo = 'entrada';
let movFaseId = 'corte';
function abrirMovFase(faseId, tipo) {
  if (!exigirAdmin('movimentar estoque')) return;
  const fase = FASES_ESTOQUE.find(f => f.id === faseId);
  if (!fase) return;
  movFaseId = faseId;
  movFaseTipo = tipo === 'saida' ? 'saida' : 'entrada';
  document.getElementById('modal-corte-title').textContent =
    (movFaseTipo === 'entrada' ? 'Entrada manual' : 'Saída / ajuste') + ' — ' + fase.titulo;
  const tecOpts = '<option value="">— selecione —</option>' + (STATE.tecidos || []).map(t => `<option value="${esc(t.nome)}">${esc(t.nome)}</option>`).join('');
  const corOpts = '<option value="">— sem cor —</option>' + (STATE.cores || []).map(c => `<option value="${esc(c.nome)}">${esc(c.nome)}</option>`).join('');
  const hoje = new Date().toISOString().slice(0, 10);
  document.getElementById('modal-corte-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Tecido *</label><select id="mc-tecido">${tecOpts}</select></div>
      <div class="field"><label>Cor</label><select id="mc-cor">${corOpts}</select></div>
      <div class="field"><label>Quantidade (peças) *</label><input type="number" min="0" step="1" id="mc-qtd" placeholder="Ex.: 50"></div>
      <div class="field"><label>Data</label><input type="date" id="mc-data" value="${hoje}"></div>
      <div class="field full"><label>Observação</label><input type="text" id="mc-obs" placeholder="Ex.: contagem de inventário / sobra"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">Entra na coluna <b>Contagem de estoque</b> de <b>${esc(fase.titulo)}</b> e ajusta o saldo. As entradas/saídas automáticas vêm das etapas do checklist — use isto só para contagem física e correções.</div>`;
  openModal('modal-corte');
}

async function salvarMovFase() {
  if (!exigirAdmin('movimentar estoque')) return;
  const fase = FASES_ESTOQUE.find(f => f.id === movFaseId);
  if (!fase) return;
  const v = id => document.getElementById(id)?.value || '';
  const tecidoNome = v('mc-tecido');
  if (!tecidoNome) return toast('Selecione o tecido', 'err');
  const qtd = parseInt(String(v('mc-qtd')).replace(',', '.')) || 0;
  if (!(qtd > 0)) return toast('Informe a quantidade em peças', 'err');
  if (!Array.isArray(STATE[fase.movKey])) STATE[fase.movKey] = [];
  STATE[fase.movKey].push({
    id: uid(),
    tipo: movFaseTipo,
    tecidoNome,
    corNome: v('mc-cor'),
    qtd,
    data: v('mc-data') || new Date().toISOString().slice(0, 10),
    obs: v('mc-obs')
  });
  await saveState(fase.movKey);
  closeModal('modal-corte');
  toast(movFaseTipo === 'entrada' ? 'Entrada registrada' : 'Saída registrada', 'ok');
  renderFasePorId(fase.id);
}

async function excluirMovFase(faseId, id) {
  if (!exigirAdmin('excluir lançamento')) return;
  const fase = FASES_ESTOQUE.find(f => f.id === faseId);
  if (!fase) return;
  if (!confirm('Excluir este lançamento manual?')) return;
  STATE[fase.movKey] = (STATE[fase.movKey] || []).filter(x => x.id !== id);
  await saveState(fase.movKey);
  toast('Lançamento excluído', 'ok');
  renderFasePorId(faseId);
}

/* ========================================================= */
/*              PLANEJAMENTO DE EXPEDIÇÃO                    */
/* ========================================================= */
// Segunda folha impressa do programa (a primeira é a folha de OS). Toda
// expedição aqui é INTERNA: ida e volta entre duas unidades. Por isso cada
// ocorrência tem DUAS pernas contabilizadas em separado — a carga que sai na
// ida não é a que volta, e cada uma tem seu próprio mínimo/máximo a respeitar.
//
// Vocabulário:
//   janela     = a regra cadastrada ("toda terça e quinta, ida 8h volta 17h")
//   ocorrência = a janela num dia concreto (janela + data)
//   perna      = ida (unidade A -> B) ou volta (B -> A)
//   carga      = uma OS alocada numa perna de uma ocorrência, com seus volumes

const EXP_CFG_PADRAO = {
  unidadeA: 'Unidade 1',
  unidadeB: 'Unidade 2',
  volMin: 0,
  volMax: 0
};

function expCfg() {
  return { ...EXP_CFG_PADRAO, ...((STATE.meta && STATE.meta.expedicao) || {}) };
}

// Número com fallback: '' e null caem no padrão em vez de virar 0 — é o que
// deixa uma janela dizer "sem limite próprio, usa o da configuração".
function _expNum(v, fallback) {
  if (v === '' || v == null) return fallback;
  const n = Number(v);
  return isNaN(n) ? fallback : n;
}

const _EXP_DIAS = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
const _EXP_DIAS_CURTO = ['DOM', 'SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SÁB'];
const _EXP_MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Datas sempre como 'YYYY-MM-DD' em horário LOCAL. new Date('2026-07-17')
// seria UTC e viraria dia 16 à noite no Brasil — daí o parse manual.
function _expIso(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function _expData(iso) {
  const [y, m, d] = String(iso || '').split('-').map(Number);
  return new Date(y || 1970, (m || 1) - 1, d || 1);
}
function _expHoje() { return _expIso(new Date()); }
function _expAddDias(iso, n) {
  const d = _expData(iso);
  d.setDate(d.getDate() + n);
  return _expIso(d);
}

// Período visível a partir do modo e da data-âncora.
function _expRange(modo, ancora) {
  if (modo === 'dia') return { ini: ancora, fim: ancora };
  if (modo === 'semana') {
    const ini = _expAddDias(ancora, -_expData(ancora).getDay()); // semana começa no domingo
    return { ini, fim: _expAddDias(ini, 6) };
  }
  const d = _expData(ancora);
  return {
    ini: _expIso(new Date(d.getFullYear(), d.getMonth(), 1)),
    fim: _expIso(new Date(d.getFullYear(), d.getMonth() + 1, 0))
  };
}

function _expNavegar(modo, ancora, dir) {
  if (modo === 'dia') return _expAddDias(ancora, dir);
  if (modo === 'semana') return _expAddDias(ancora, dir * 7);
  const d = _expData(ancora);
  return _expIso(new Date(d.getFullYear(), d.getMonth() + dir, 1));
}

function _expLabelPeriodo(modo, ancora) {
  const { ini, fim } = _expRange(modo, ancora);
  if (modo === 'dia') return _EXP_DIAS[_expData(ini).getDay()] + ', ' + formatDate(ini);
  if (modo === 'semana') return formatDate(ini) + ' — ' + formatDate(fim);
  const d = _expData(ini);
  return _EXP_MESES[d.getMonth()] + ' de ' + d.getFullYear();
}

function _expNomeModo(modo) {
  return modo === 'dia' ? 'diário' : (modo === 'semana' ? 'semanal' : 'mensal');
}

// Ocorrências das janelas ativas dentro de [ini, fim].
// Uma cancelada continua aparecendo (riscada) no período dela: o usuário
// precisa ver que a expedição foi suspensa, não que sumiu por engano.
function ocorrenciasExpedicao(ini, fim) {
  const out = [];
  const excecoes = STATE.expedicaoExcecoes || [];
  // Range folgado: uma ocorrência remarcada PARA dentro do período nasceu
  // fora dele, então precisa ser gerada antes de ser filtrada.
  const iniG = _expAddDias(ini, -60), fimG = _expAddDias(fim, 60);
  (STATE.expedicaoJanelas || []).forEach(j => {
    if (j.ativo === false) return;
    const datas = [];
    if (j.tipo === 'data') {
      if (j.data && j.data >= iniG && j.data <= fimG) datas.push(j.data);
    } else {
      const dias = (Array.isArray(j.diasSemana) ? j.diasSemana : []).map(Number);
      if (!dias.length) return;
      for (let d = iniG; d <= fimG; d = _expAddDias(d, 1)) {
        if (dias.includes(_expData(d).getDay())) datas.push(d);
      }
    }
    datas.forEach(data => {
      const exc = excecoes.find(e => e.janelaId === j.id && e.data === data);
      const base = { janela: j, dataOrig: data, chave: j.id + '|' + data };
      if (exc && exc.tipo === 'cancelada') {
        if (data >= ini && data <= fim) {
          out.push({ ...base, data, horaIda: j.horaIda || '', horaVolta: j.horaVolta || '', cancelada: true, remarcada: false, motivo: exc.motivo || '' });
        }
        return;
      }
      const dataFinal = (exc && exc.tipo === 'remarcada' && exc.novaData) ? exc.novaData : data;
      if (dataFinal < ini || dataFinal > fim) return;
      out.push({
        ...base,
        data: dataFinal,
        horaIda: (exc && exc.horaIda) || j.horaIda || '',
        horaVolta: (exc && exc.horaVolta) || j.horaVolta || '',
        cancelada: false,
        remarcada: !!(exc && exc.tipo === 'remarcada'),
        motivo: (exc && exc.motivo) || ''
      });
    });
  });
  return out.sort((a, b) =>
    a.data.localeCompare(b.data) ||
    String(a.horaIda || '').localeCompare(String(b.horaIda || '')) ||
    String(a.janela.nome || '').localeCompare(String(b.janela.nome || ''))
  );
}

function _expPecasOS(o) {
  return componentesPorTecidoCorOS(o).reduce((s, it) => s + it.qtd, 0);
}

function _expCargasDa(janelaId, dataOrig, perna) {
  return (STATE.expedicaoCargas || []).filter(c => c.janelaId === janelaId && c.data === dataOrig && c.perna === perna);
}

// Carga de uma perna: as OSs alocadas, os totais e a situação contra os
// limites da janela (que herdam da configuração quando em branco).
function resumoPernaExpedicao(oc, perna) {
  const cfg = expCfg();
  const volMin = _expNum(oc.janela.volMin, _expNum(cfg.volMin, 0));
  const volMax = _expNum(oc.janela.volMax, _expNum(cfg.volMax, 0));
  const itens = _expCargasDa(oc.janela.id, oc.dataOrig, perna).map(c => {
    const o = (STATE.ordens || []).find(x => x.id === c.osId);
    return {
      carga: c,
      os: o,
      osNumero: o ? (o.os || '—') : '(OS excluída)',
      modelo: o ? (o.modeloNome || '') : '',
      pecas: o ? _expPecasOS(o) : 0,
      volumes: Number(c.volumes) || 0
    };
  }).sort((a, b) => String(a.osNumero).localeCompare(String(b.osNumero), undefined, { numeric: true }));
  const volumes = itens.reduce((s, i) => s + i.volumes, 0);
  const pecas = itens.reduce((s, i) => s + i.pecas, 0);
  // OS que entrou no plano pelo checklist sem "peças por volume" configurado
  // chega com 0 volumes: conta como carga, mas ninguém disse quanto ocupa.
  const semVolumes = itens.filter(i => !(i.volumes > 0)).length;
  let situacao = 'ok';
  if (!itens.length) situacao = 'vazio';
  else if (volMax > 0 && volumes > volMax) situacao = 'alto';
  else if (volMin > 0 && volumes < volMin) situacao = 'baixo';
  return { itens, volumes, pecas, volMin, volMax, situacao, semVolumes };
}

const _EXP_SIT_LABEL = { ok: 'dentro', baixo: 'abaixo do mín.', alto: 'acima do máx.', vazio: 'sem carga' };

// Como a folha impressa se refere ao período quando não há nenhuma OE produzida.
const _EXP_VAZIO_PERIODO = { dia: 'neste dia', semana: 'nesta semana', mes: 'neste mês' };

// Texto dos limites da perna, pra não repetir a regra em 4 lugares.
function _expLimitesTexto(volMin, volMax) {
  if (volMin > 0 && volMax > 0) return `mín ${volMin} / máx ${volMax}`;
  if (volMin > 0) return `mín ${volMin}`;
  if (volMax > 0) return `máx ${volMax}`;
  return 'sem limite';
}

function _expRotaTexto(perna) {
  const cfg = expCfg();
  return perna === 'ida' ? `${cfg.unidadeA} → ${cfg.unidadeB}` : `${cfg.unidadeB} → ${cfg.unidadeA}`;
}

// A OS é de blusa de moletom? (algum tecido da OS é categoria 'moletom'.)
// Muda a regra de pacotes: moletom conta 1 por tamanho distinto; camiseta
// conta 1 por unidade (soma das quantidades).
function _osEhMoletom(o) {
  if (!o) return false;
  const ehMol = tecId => {
    const t = (STATE.tecidos || []).find(x => x.id === tecId);
    return !!t && categoriaEfetivaTecido(t) === 'moletom';
  };
  return (o.fases || []).some(f => ehMol(f.tecidoId))
      || (o.tecidos || []).some(t => ehMol(t.tecidoId));
}

// Nº de "vagas" de tamanho da grade = base do volume de expedição (e das
// etiquetas). Duas regras por tipo de produto:
//   • Camiseta: 1 pacote por UNIDADE de tamanho (soma das quantidades; 2M = 2).
//   • Moletom : 1 pacote por TAMANHO distinto (a quantidade/multiplicador não
//     multiplica os pacotes — ex.: "2X P ao G3" = 7, não 14).
// Prefere a grade viva (como a folha impressa), caindo no snapshot salvo na OS.
function _expTotalTamanhosGrade(o) {
  const keys = ['p','m','g','gg','g1','g2','g3'];
  let tam = null;
  if (o && o.gradeId) {
    const g = (STATE.grades || []).find(x => x.id === o.gradeId);
    if (g && g.tamanhos) tam = g.tamanhos;
  }
  if (!tam && o && o.grade) tam = o.grade;
  if (!tam) return 0;
  if (_osEhMoletom(o)) return keys.filter(k => (parseInt(tam[k]) || 0) > 0).length;
  return keys.reduce((s, k) => s + (parseInt(tam[k]) || 0), 0);
}

// Volume (pacotes) de uma OS: nº de vagas de tamanho (_expTotalTamanhosGrade,
// que já aplica a regra por tipo) × nº de TONALIDADES + 1 pacote de reposição.
// Cada tonalidade é ensacada separada, então uma grade em dois tons dobra os
// pacotes. Não depende de peças nem de camadas.
// Ex. moletom P ao G3: 1 tom → 7×1+1=8; 2 tons → 7×2+1=15; 3 tons → 7×3+1=22.
// Ex. camiseta P-G1-G2: 1 tom → 3×1+1=4; 2 tons → 3×2+1=7.
// OS sem tonalidade marcada conta como 1 tom (comportamento antigo preservado).
function _expSugestaoVolumes(o) {
  const nTam = _expTotalTamanhosGrade(o);
  if (!(nTam > 0)) return '';
  const nTons = Math.max(1, tonsEfetivos(((o || {}).progresso || {}).totalTamanhoTons || {}).length);
  return String(nTam * nTons + 1);
}

// Aviso na tela quando o volume GRAVADO na carga não bate com a regra
// (tamanhos × tonalidades + 1). A propagação cobre as cargas futuras a partir do
// momento em que a tonalidade muda, mas não alcança as que já estavam gravadas
// com número velho antes disso — nem um ajuste manual que ficou defasado. Aqui
// elas ficam visíveis, em vez de irem caladas para a OE.
// Só avisa: quem decide é o usuário, que pode ter posto o número à mão de propósito.
function _expBadgeVolumeDivergente(item) {
  if (!item || !item.os || !(item.volumes > 0)) return '';
  const esperado = Number(_expSugestaoVolumes(item.os)) || 0;
  if (!(esperado > 0) || esperado === item.volumes) return '';
  const nTons = Math.max(1, tonsEfetivos((item.os.progresso || {}).totalTamanhoTons || {}).length);
  return ` <span class="exp-badge baixo" title="A grade em ${nTons} tonalidade(s) dá ${esperado} volumes, mas esta carga está com ${item.volumes}. Use ↻ Recalcular volumes, ou deixe como está se o ajuste foi proposital.">≠ ${esperado}</span>`;
}

// Reescreve o volume das cargas AINDA NÃO REALIZADAS desta OS pela regra.
// Sem isto o número fica congelado no instante em que a OS entrou no plano: se
// a tonalidade for marcada depois — o caminho normal, já que o Tom 2 costuma ser
// marcado durante o enfesto — a OE seguiria imprimindo o volume antigo.
// Expedição já realizada é histórico do que saiu no caminhão: não se reescreve.
// Devolve quantas cargas mudaram.
async function propagarVolumesExpedicaoOS(os) {
  if (!os || !Array.isArray(STATE.expedicaoCargas)) return 0;
  const sug = Number(_expSugestaoVolumes(os)) || 0;
  if (!(sug > 0)) return 0;
  const hoje = _expHoje();
  let n = 0;
  STATE.expedicaoCargas.forEach(c => {
    if (c.osId !== os.id) return;
    if (_expDataEfetivaCarga(c) < hoje) return;
    if ((Number(c.volumes) || 0) === sug) return;
    c.volumes = sug;
    n++;
  });
  if (n) { try { await saveState('expedicaoCargas'); } catch (e) { console.warn('propagarVolumesExpedicaoOS', e); } }
  return n;
}

/* ------- seleção da OS pelo checklist da folha de OS ------- */
// Marcar "Ensaque" no checklist da folha de OS é o que seleciona a OS pra ser
// expedida: ensacada = pacote pronto pra embarcar. Ela cai sozinha na próxima
// janela; trocar a janela é depois, no planejamento (moverCargaExp).
//
// Ensaque não tem campo de estoque próprio (não está em FASES_ESTOQUE), então
// a detecção é pela caixinha marcada — não por faseAtualOS, que só enxerga
// etapas com campo. O volume de peças da OS é o mesmo do Estoque de corte
// (ambos = soma dos componentes, via _expPecasOS).

const ENSAQUE_ETAPA_RE = /ensaque/i;

// A OS está ensacada? (caixinha "Ensaque" marcada no checklist da folha.)
function osEnsacada(o) { return osEtapaMarcada(o, ENSAQUE_ETAPA_RE); }

// A carga guarda a data ORIGINAL da ocorrência; se ela foi remarcada, a data
// em que a expedição de fato acontece é outra.
function _expDataEfetivaCarga(c) {
  const exc = (STATE.expedicaoExcecoes || []).find(e => e.janelaId === c.janelaId && e.data === c.data);
  if (exc && exc.tipo === 'remarcada' && exc.novaData) return exc.novaData;
  return c.data;
}

// Primeira expedição de hoje em diante. É onde a OS marcada aterrissa.
function _expProximaOcorrencia() {
  const hoje = _expHoje();
  return ocorrenciasExpedicao(hoje, _expAddDias(hoje, 180)).find(o => !o.cancelada) || null;
}

async function sincronizarPlanoExpedicaoDaOS(os, etapaNome, checked) {
  if (!os || !ENSAQUE_ETAPA_RE.test(etapaNome || '')) return;
  if (!Array.isArray(STATE.expedicaoCargas)) STATE.expedicaoCargas = [];
  const num = (os.os || '').toString().trim();

  if (checked) {
    // Já planejada em algum lugar: respeita onde o usuário a pôs.
    if (STATE.expedicaoCargas.some(c => c.osId === os.id)) return;
    const oc = _expProximaOcorrencia();
    if (!oc) {
      toast('OS marcada para expedir, mas não há janela de expedição cadastrada', 'err');
      return;
    }
    STATE.expedicaoCargas.push({
      id: uid(), janelaId: oc.janela.id, data: oc.dataOrig, perna: 'ida',
      osId: os.id, volumes: Number(_expSugestaoVolumes(os)) || 0, obs: ''
    });
    await saveState('expedicaoCargas');
    toast(`OS ${num} entrou na expedição de ${_EXP_DIAS_CURTO[_expData(oc.data).getDay()]} ${formatDate(oc.data)}${oc.horaIda ? ' ' + oc.horaIda : ''}`, 'ok');
  } else {
    // Desmarcar tira do que ainda vai acontecer. Expedição já passada é
    // histórico do que saiu no caminhão — não se reescreve por um clique.
    const hoje = _expHoje();
    const antes = STATE.expedicaoCargas.length;
    STATE.expedicaoCargas = STATE.expedicaoCargas.filter(c =>
      c.osId !== os.id || _expDataEfetivaCarga(c) < hoje);
    const tiradas = antes - STATE.expedicaoCargas.length;
    if (tiradas) {
      await saveState('expedicaoCargas');
      toast(`OS ${num} saiu do plano de expedição`, 'ok');
    }
  }
  if (expAbaAtiva === 'plano' && document.getElementById('expedicao-plano')) renderExpedicaoPlano();
}

function moverCargaExp(cargaId) {
  const c = (STATE.expedicaoCargas || []).find(x => x.id === cargaId);
  if (!c) return;
  abrirModalExpCarga(c.janelaId, c.data, c.perna, c.osId, cargaId);
}

/* ---------------- estado da tela ---------------- */

let expPlanoModo = 'semana';
let expPlanoAncora = _expHoje();
let expAbaAtiva = 'estoque';
try {
  expPlanoModo = sessionStorage.getItem('gos:exp:modo') || expPlanoModo;
  expPlanoAncora = sessionStorage.getItem('gos:exp:ancora') || expPlanoAncora;
  expAbaAtiva = sessionStorage.getItem('gos:exp:aba') || expAbaAtiva;
} catch (e) { /* sessionStorage indisponível, segue no padrão */ }

function trocarAbaExpedicao(aba) {
  expAbaAtiva = (aba === 'plano') ? 'plano' : 'estoque';
  try { sessionStorage.setItem('gos:exp:aba', expAbaAtiva); } catch (e) {}
  document.querySelectorAll('.exp-tab').forEach(b => b.classList.toggle('active', b.dataset.exptab === expAbaAtiva));
  const est = document.getElementById('expedicao-aba-estoque');
  const plano = document.getElementById('expedicao-aba-plano');
  if (est) est.classList.toggle('hidden', expAbaAtiva !== 'estoque');
  if (plano) plano.classList.toggle('hidden', expAbaAtiva !== 'plano');
  if (expAbaAtiva === 'plano') renderExpedicaoPlano();
}

function expSetModo(modo) {
  expPlanoModo = modo;
  try { sessionStorage.setItem('gos:exp:modo', modo); } catch (e) {}
  renderExpedicaoPlano();
}

function expNav(dir) {
  expPlanoAncora = _expNavegar(expPlanoModo, expPlanoAncora, dir);
  try { sessionStorage.setItem('gos:exp:ancora', expPlanoAncora); } catch (e) {}
  renderExpedicaoPlano();
}

function expHoje() {
  expPlanoAncora = _expHoje();
  try { sessionStorage.setItem('gos:exp:ancora', expPlanoAncora); } catch (e) {}
  renderExpedicaoPlano();
}

/* ---------------- render do planejamento ---------------- */

function renderExpedicaoPlano() {
  const cont = document.getElementById('expedicao-plano');
  if (!cont) return;
  const cfg = expCfg();
  const { ini, fim } = _expRange(expPlanoModo, expPlanoAncora);
  const ocs = ocorrenciasExpedicao(ini, fim);
  const fmt = n => (Number(n) || 0).toLocaleString('pt-BR');

  const toolbar = `
    <div class="exp-toolbar no-print">
      <div class="exp-seg">
        <button class="${expPlanoModo === 'dia' ? 'active' : ''}" onclick="expSetModo('dia')">Diário</button>
        <button class="${expPlanoModo === 'semana' ? 'active' : ''}" onclick="expSetModo('semana')">Semanal</button>
        <button class="${expPlanoModo === 'mes' ? 'active' : ''}" onclick="expSetModo('mes')">Mensal</button>
      </div>
      <div style="display:flex;align-items:center;gap:6px;">
        <button class="btn" onclick="expNav(-1)" title="Período anterior">‹</button>
        <div class="exp-periodo">${esc(_expLabelPeriodo(expPlanoModo, expPlanoAncora))}</div>
        <button class="btn" onclick="expNav(1)" title="Próximo período">›</button>
        <button class="btn" onclick="expHoje()">Hoje</button>
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;">
        <button class="btn accent" onclick="goto('print-expedicao')">🖨 Folha do plano</button>
        <div class="admin-only" style="display:flex;gap:6px;">
          <button class="btn primary" onclick="abrirModalExpJanela()">+ Janela</button>
          <button class="btn" onclick="abrirModalExpConfig()">⚙ Unidades e carga</button>
          <button class="btn" onclick="recalcularVolumesExpedicao()" title="Redefine os volumes das expedições futuras pela regra da grade (1 pacote por tamanho, por tonalidade, + 1 de reposição). Não mexe em expedições já realizadas.">↻ Recalcular volumes</button>
        </div>
      </div>
    </div>`;

  // Totais do período. Ida e volta somam separado: é essa distinção que
  // torna a expedição interna diferente de uma saída simples.
  let volIda = 0, volVolta = 0, pecasIda = 0, pecasVolta = 0, alertas = 0, ativas = 0;
  const osAlocadas = new Set();
  ocs.forEach(oc => {
    if (oc.cancelada) return;
    ativas++;
    ['ida', 'volta'].forEach(perna => {
      const r = resumoPernaExpedicao(oc, perna);
      if (perna === 'ida') { volIda += r.volumes; pecasIda += r.pecas; }
      else { volVolta += r.volumes; pecasVolta += r.pecas; }
      if (r.situacao === 'baixo' || r.situacao === 'alto') alertas++;
      r.itens.forEach(i => { if (i.os) osAlocadas.add(i.os.id); });
    });
  });

  const comoFunciona = `
    <div class="info-box no-print" style="font-size:12px;">
      As OSs entram aqui sozinhas: marcar <b>Ensaque</b> no checklist da folha de OS põe a OS na <b>próxima expedição</b> (perna de ida). Use <b>⇄</b> em cada OS para mudar o dia e o horário em que ela sai.
    </div>`;

  const resumo = `
    <div class="exp-resumo">
      <div class="item"><div class="num">${fmt(ativas)}</div><div class="lbl">Expedições no período</div></div>
      <div class="item"><div class="num">${fmt(volIda)}</div><div class="lbl">Volumes na ida</div></div>
      <div class="item"><div class="num">${fmt(volVolta)}</div><div class="lbl">Volumes na volta</div></div>
      <div class="item"><div class="num">${fmt(volIda + volVolta)}</div><div class="lbl">Volumes no total</div></div>
      <div class="item"><div class="num">${fmt(pecasIda + pecasVolta)}</div><div class="lbl">Peças movimentadas</div></div>
      <div class="item"><div class="num">${fmt(osAlocadas.size)}</div><div class="lbl">OS alocadas</div></div>
      <div class="item ${alertas ? 'alerta' : ''}"><div class="num">${fmt(alertas)}</div><div class="lbl">Cargas fora do limite</div></div>
    </div>`;

  const pernaHtml = (oc, perna) => {
    const r = resumoPernaExpedicao(oc, perna);
    const hora = perna === 'ida' ? oc.horaIda : oc.horaVolta;
    const linhas = r.itens.length ? r.itens.map(i => `
      <div class="exp-os-row">
        <span class="num">${esc(i.osNumero)}</span>
        <span class="mod">${esc(i.modelo) || '—'}</span>
        <span class="qtd">${fmt(i.pecas)} pç</span>
        <span class="vol">${i.volumes > 0 ? fmt(i.volumes) + ' vol' : '<span class="exp-badge baixo" title="Ninguém disse quantos volumes esta OS ocupa">vol?</span>'}${_expBadgeVolumeDivergente(i)}</span>
        <span class="admin-only"><button title="Mudar o dia e o horário em que esta OS será expedida" onclick="moverCargaExp('${esc(i.carga.id)}')">⇄</button><button title="Tirar esta OS da carga" onclick="excluirCargaExp('${esc(i.carga.id)}')">×</button></span>
      </div>`).join('') : '<div class="exp-vazio">Nenhuma OS alocada.</div>';
    return `
      <div class="exp-perna">
        <div class="exp-perna-head">
          <div>
            <div class="exp-perna-tit">${perna === 'ida' ? 'Ida' : 'Volta'}</div>
            <div class="exp-perna-rota">${esc(_expRotaTexto(perna))}</div>
          </div>
          <div class="exp-perna-hora">${esc(hora) || '—'}</div>
        </div>
        <div class="exp-os-list">${linhas}</div>
        <div class="exp-perna-total">
          <span>
            <span class="vol">${fmt(r.volumes)}</span> vol
            <span style="color:var(--ink-3);"> · ${fmt(r.pecas)} pç · ${esc(_expLimitesTexto(r.volMin, r.volMax))}</span>
            ${r.semVolumes ? `<br><span style="color:var(--accent-dark);font-size:11px;">${r.semVolumes} OS sem volumes definidos — o total está incompleto</span>` : ''}
          </span>
          <span class="exp-badge ${r.situacao}">${esc(_EXP_SIT_LABEL[r.situacao])}</span>
        </div>
        ${oc.cancelada ? '' : `<div class="admin-only" style="margin-top:8px;display:flex;gap:6px;">
          <button class="btn" style="flex:1;padding:5px;font-size:12px;" onclick="abrirModalExpCarga('${esc(oc.janela.id)}','${esc(oc.dataOrig)}','${perna}')">+ Alocar OS</button>
          ${perna === 'volta' ? `<button class="btn" style="flex:1;padding:5px;font-size:12px;" title="Traz para esta volta as OSs de uma expedição já montada — normalmente a ida que levou as peças." onclick="abrirModalExpVolta('${esc(oc.janela.id)}','${esc(oc.dataOrig)}')">⟲ Trazer de uma OE</button>` : ''}
        </div>`}
      </div>`;
  };

  const cards = ocs.map(oc => `
    <div class="card exp-ocor ${oc.cancelada ? 'cancelada' : ''}">
      <div class="exp-ocor-head">
        <div>
          <div class="exp-ocor-data" style="${oc.cancelada ? 'text-decoration:line-through;' : ''}">
            ${_EXP_DIAS_CURTO[_expData(oc.data).getDay()]} · ${esc(formatDate(oc.data))}
          </div>
          <div class="exp-ocor-nome">
            ${esc(oc.janela.nome) || 'Janela sem nome'}
            ${oc.janela.tipo === 'data' ? ' · <span class="exp-badge info">data fixa</span>' : ''}
            ${oc.remarcada ? ` · <span class="exp-badge baixo">remarcada de ${esc(formatDate(oc.dataOrig))}</span>` : ''}
            ${oc.cancelada ? ' · <span class="exp-badge alto">cancelada</span>' : ''}
            ${oc.motivo ? ' · ' + esc(oc.motivo) : ''}
          </div>
        </div>
        <div class="admin-only" style="display:flex;gap:6px;">
          <button class="btn" onclick="abrirModalExpOcorrencia('${esc(oc.janela.id)}','${esc(oc.dataOrig)}')">Cancelar / remarcar</button>
          <button class="btn" onclick="abrirModalExpJanela('${esc(oc.janela.id)}')">Editar janela</button>
        </div>
      </div>
      <div class="exp-pernas">
        ${pernaHtml(oc, 'ida')}
        ${pernaHtml(oc, 'volta')}
      </div>
    </div>`).join('');

  const semJanelas = !(STATE.expedicaoJanelas || []).length;
  const vazio = `
    <div class="card">
      <div class="empty" style="padding:24px 0;text-align:center;">
        ${semJanelas
          ? 'Nenhuma janela de expedição cadastrada. Clique em <b>+ Janela</b> para definir os dias e horários em que a expedição acontece.'
          : 'Nenhuma expedição neste período. Navegue entre os períodos ou cadastre uma janela para estes dias.'}
      </div>
    </div>`;

  // OSs ensacadas (prontas) que ninguém colocou em carga nenhuma. É a lista
  // que evita esquecer OS pronta parada no campo.
  const alocadasSempre = new Set((STATE.expedicaoCargas || []).map(c => c.osId));
  const pendentes = (STATE.ordens || [])
    .filter(o => osEnsacada(o) && !alocadasSempre.has(o.id))
    .map(o => ({ o, pecas: _expPecasOS(o) }))
    .filter(x => x.pecas > 0)
    .sort((a, b) => String(b.o.os || '').localeCompare(String(a.o.os || ''), undefined, { numeric: true }));
  const pendentesHtml = pendentes.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">OSs ensacadas sem carga alocada <span class="exp-badge baixo">${pendentes.length}</span></h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Estão com a etapa <b>Ensaque</b> marcada mas não entraram em nenhuma expedição — nem passada, nem planejada. Acontece com OS ensacada antes de existir janela cadastrada. Use <b>alocar</b> para pô-las numa expedição.</div>
      <table class="table">
        <thead><tr><th>OS</th><th>Modelo</th><th>Data</th><th style="text-align:right;">Peças</th><th class="col-actions">Ações</th></tr></thead>
        <tbody>
          ${pendentes.map(({ o, pecas }) => `
            <tr>
              <td><strong>${esc(o.os) || '—'}</strong></td>
              <td>${esc(o.modeloNome) || '—'}</td>
              <td style="white-space:nowrap;">${esc(formatDate(o.data))}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(pecas)} pç</td>
              <td class="col-actions row-actions">
                <button onclick="verOS('${esc(o.id)}')">ver OS</button>
                <button class="edit admin-only" onclick="abrirModalExpCarga('','','ida','${esc(o.id)}')">alocar</button>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';

  // Janelas cadastradas: o "de onde vêm" das ocorrências acima.
  const janelas = (STATE.expedicaoJanelas || []).slice().sort((a, b) =>
    String(a.tipo).localeCompare(String(b.tipo)) || String(a.horaIda || '').localeCompare(String(b.horaIda || ''))
  );
  const janelasHtml = `
    <div class="card admin-only">
      <div class="card-title">Janelas de expedição cadastradas</div>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Uma janela <b>semanal</b> se repete nos dias marcados; uma de <b>data fixa</b> acontece uma vez só. Mín/máx em branco herdam a configuração de <b>Unidades e carga</b> (hoje: ${esc(_expLimitesTexto(_expNum(cfg.volMin, 0), _expNum(cfg.volMax, 0)))}).</div>
      <table class="table">
        <thead><tr><th>Nome</th><th>Quando</th><th>Ida</th><th>Volta</th><th>Volumes</th><th>Situação</th><th class="col-actions">Ações</th></tr></thead>
        <tbody>
          ${janelas.length ? janelas.map(j => `
            <tr>
              <td><strong>${esc(j.nome) || '—'}</strong></td>
              <td>${j.tipo === 'data'
                ? esc(formatDate(j.data))
                : ((j.diasSemana || []).length ? (j.diasSemana || []).slice().sort((a, b) => a - b).map(d => _EXP_DIAS_CURTO[d]).join(', ') : '<span class="exp-badge alto">sem dias</span>')}</td>
              <td style="font-family:'IBM Plex Mono',monospace;">${esc(j.horaIda) || '—'}</td>
              <td style="font-family:'IBM Plex Mono',monospace;">${esc(j.horaVolta) || '—'}</td>
              <td>${esc(_expLimitesTexto(_expNum(j.volMin, _expNum(cfg.volMin, 0)), _expNum(j.volMax, _expNum(cfg.volMax, 0))))}</td>
              <td>${j.ativo === false ? '<span class="exp-badge vazio">inativa</span>' : '<span class="exp-badge ok">ativa</span>'}</td>
              <td class="col-actions row-actions">
                <button class="edit" onclick="abrirModalExpJanela('${esc(j.id)}')">editar</button>
                <button class="del" onclick="excluirJanelaExp('${esc(j.id)}')">excluir</button>
              </td>
            </tr>`).join('') : '<tr><td colspan="7" class="empty">Nenhuma janela cadastrada.</td></tr>'}
        </tbody>
      </table>
    </div>`;

  cont.innerHTML = toolbar + comoFunciona + resumo + (ocs.length ? cards : vazio) + pendentesHtml + janelasHtml;
}

/* ---------------- modais ---------------- */

let _expModalCtx = null;

function _expCampoNum(id, label, valor, hint) {
  return `<div class="field"><label>${label}</label><input type="number" min="0" step="1" id="${id}" value="${valor === '' || valor == null ? '' : esc(valor)}">${hint ? `<div class="field-hint">${hint}</div>` : ''}</div>`;
}

function abrirModalExpJanela(editId = null) {
  if (!exigirAdmin('cadastrar janelas de expedição')) return;
  const j = editId ? (STATE.expedicaoJanelas || []).find(x => x.id === editId) : null;
  if (editId && !j) return;
  _expModalCtx = { tipo: 'janela', editId };
  const cfg = expCfg();
  const tipo = j ? (j.tipo || 'semanal') : 'semanal';
  const dias = (j && Array.isArray(j.diasSemana)) ? j.diasSemana.map(Number) : [];
  document.getElementById('modal-exp-title').textContent = editId ? 'Editar janela de expedição' : 'Nova janela de expedição';
  document.getElementById('modal-exp-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Nome *</label><input type="text" id="ej-nome" value="${esc(j ? j.nome : '')}" placeholder="Ex.: Expedição da manhã"></div>
      <div class="field">
        <label>Tipo *</label>
        <select id="ej-tipo" onchange="_expToggleTipoJanela()">
          <option value="semanal" ${tipo === 'semanal' ? 'selected' : ''}>Semanal (repete nos dias marcados)</option>
          <option value="data" ${tipo === 'data' ? 'selected' : ''}>Data fixa (acontece uma vez)</option>
        </select>
      </div>
      <div class="field full" id="ej-wrap-dias">
        <label>Dias da semana *</label>
        <div style="display:flex;gap:10px;flex-wrap:wrap;padding:4px 0;">
          ${_EXP_DIAS.map((nome, i) => `
            <label style="display:flex;align-items:center;gap:4px;font-size:12px;cursor:pointer;">
              <input type="checkbox" class="ej-dia" value="${i}" ${dias.includes(i) ? 'checked' : ''}> ${_EXP_DIAS_CURTO[i]}
            </label>`).join('')}
        </div>
      </div>
      <div class="field" id="ej-wrap-data"><label>Data *</label><input type="date" id="ej-data" value="${esc(j ? (j.data || '') : _expHoje())}"></div>
      <div class="field"><label>Hora da ida *</label><input type="time" id="ej-hora-ida" value="${esc(j ? (j.horaIda || '') : '08:00')}"><div class="field-hint">${esc(cfg.unidadeA)} → ${esc(cfg.unidadeB)}</div></div>
      <div class="field"><label>Hora da volta *</label><input type="time" id="ej-hora-volta" value="${esc(j ? (j.horaVolta || '') : '17:00')}"><div class="field-hint">${esc(cfg.unidadeB)} → ${esc(cfg.unidadeA)}</div></div>
      ${_expCampoNum('ej-vol-min', 'Volume mínimo', j ? j.volMin : '', `Em branco usa o padrão (${_expNum(cfg.volMin, 0) || 'sem mínimo'}). Vale por perna.`)}
      ${_expCampoNum('ej-vol-max', 'Volume máximo', j ? j.volMax : '', `Em branco usa o padrão (${_expNum(cfg.volMax, 0) || 'sem máximo'}). Vale por perna.`)}
      <div class="field"><label>Situação</label><select id="ej-ativo"><option value="1" ${!j || j.ativo !== false ? 'selected' : ''}>Ativa</option><option value="0" ${j && j.ativo === false ? 'selected' : ''}>Inativa (não gera expedições)</option></select></div>
      <div class="field full"><label>Observação</label><input type="text" id="ej-obs" value="${esc(j ? (j.obs || '') : '')}" placeholder="Ex.: motorista da tarde"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">Toda expedição é interna, de <b>ida e volta</b> entre ${esc(cfg.unidadeA)} e ${esc(cfg.unidadeB)}. Os limites de volume valem para cada perna separadamente.</div>`;
  _expToggleTipoJanela();
  openModal('modal-exp');
}

function _expToggleTipoJanela() {
  const tipo = document.getElementById('ej-tipo')?.value || 'semanal';
  document.getElementById('ej-wrap-dias')?.classList.toggle('hidden', tipo !== 'semanal');
  document.getElementById('ej-wrap-data')?.classList.toggle('hidden', tipo !== 'data');
}

// Alocar uma OS. Aberto de dentro de uma perna (ocorrência já escolhida) ou da
// lista de pendentes (OS já escolhida) — os dois campos ficam editáveis nos
// dois casos.
function abrirModalExpCarga(janelaId, dataOrig, perna, osIdPre = '', cargaId = '') {
  if (!exigirAdmin('alocar OS na expedição')) return;
  if (!(STATE.expedicaoJanelas || []).some(j => j.ativo !== false)) {
    return toast('Cadastre uma janela de expedição antes de alocar OS', 'err');
  }
  const cargaEdit = cargaId ? (STATE.expedicaoCargas || []).find(c => c.id === cargaId) : null;
  _expModalCtx = { tipo: 'carga', editId: cargaEdit ? cargaId : '' };

  // Ocorrências oferecidas: do começo do período (ou de hoje, o que vier antes)
  // até 90 dias após o fim dele — cobre a que foi clicada e as próximas.
  const { ini, fim } = _expRange(expPlanoModo, expPlanoAncora);
  const hoje = _expHoje();
  const ocs = ocorrenciasExpedicao(ini < hoje ? ini : hoje, _expAddDias(fim, 90)).filter(o => !o.cancelada);
  const selecionada = janelaId ? `${janelaId}|${dataOrig}|${perna}` : '';
  const opts = ocs.map(oc => ['ida', 'volta'].map(p => {
    const val = `${oc.janela.id}|${oc.dataOrig}|${p}`;
    const hora = p === 'ida' ? oc.horaIda : oc.horaVolta;
    const label = `${_EXP_DIAS_CURTO[_expData(oc.data).getDay()]} ${formatDate(oc.data)} · ${hora || '—'} · ${p === 'ida' ? 'IDA' : 'VOLTA'} · ${oc.janela.nome || 'sem nome'}`;
    return `<option value="${esc(val)}" ${val === selecionada ? 'selected' : ''}>${esc(label)}</option>`;
  }).join('')).join('');

  // OSs: as ensacadas (prontas pra embarcar) primeiro; as demais ficam
  // disponíveis porque adiantar carga de OS que ainda vai chegar é legítimo.
  const ensacadas = [], outras = [];
  (STATE.ordens || []).forEach(o => {
    const pecas = _expPecasOS(o);
    if (!(pecas > 0)) return;
    const label = `${o.os || '(sem nº)'} · ${o.modeloNome || 'sem modelo'} · ${pecas.toLocaleString('pt-BR')} pç`;
    (osEnsacada(o) ? ensacadas : outras).push({ id: o.id, label });
  });
  const ordena = arr => arr.sort((a, b) => String(b.label).localeCompare(String(a.label), undefined, { numeric: true }));
  const optOS = arr => ordena(arr).map(x => `<option value="${esc(x.id)}" ${x.id === osIdPre ? 'selected' : ''}>${esc(x.label)}</option>`).join('');
  const osPre = osIdPre ? (STATE.ordens || []).find(o => o.id === osIdPre) : null;

  document.getElementById('modal-exp-title').textContent = cargaEdit ? 'Mudar a expedição desta OS' : 'Alocar OS na expedição';
  document.getElementById('modal-exp-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field full">
        <label>Expedição (data · hora · perna) *</label>
        <select id="ec-ocorrencia">${opts || '<option value="">— nenhuma expedição planejada —</option>'}</select>
        <div class="field-hint">${cargaEdit ? 'Escolha o dia e o horário em que esta OS será expedida. ' : ''}A perna define o trajeto: IDA é ${esc(expCfg().unidadeA)} → ${esc(expCfg().unidadeB)}; VOLTA é o caminho inverso.</div>
      </div>
      <div class="field full">
        <label>OS *</label>
        ${cargaEdit ? '' : `<input type="search" id="ec-os-busca" oninput="_expFiltrarOS()" placeholder="Buscar pelo número da OS ou modelo…" style="margin-bottom:6px;" autocomplete="off">`}
        <select id="ec-os" onchange="_expAtualizarSugestaoVolumes()">
          <option value="">— selecione —</option>
          ${ensacadas.length ? `<optgroup label="Ensacadas (prontas)">${optOS(ensacadas)}</optgroup>` : ''}
          ${outras.length ? `<optgroup label="Outras OS">${optOS(outras)}</optgroup>` : ''}
        </select>
        ${cargaEdit ? '' : '<div class="field-hint" id="ec-os-vazio" style="display:none;color:var(--alert);">Nenhuma OS encontrada para essa busca.</div>'}
      </div>
      ${_expCampoNum('ec-volumes', 'Volumes (sacos / caixas) *',
        cargaEdit ? (cargaEdit.volumes || '') : (osPre ? _expSugestaoVolumes(osPre) : ''),
        'É este número que conta contra o mínimo e o máximo da carga.')}
      <div class="field"><label>Observação</label><input type="text" id="ec-obs" value="${esc(cargaEdit ? (cargaEdit.obs || '') : '')}" placeholder="Ex.: vai junto com a grade de mostruário"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;" id="ec-info">Selecione a OS para ver as peças.</div>`;
  _expAtualizarSugestaoVolumes();
  openModal('modal-exp');
}

// Filtra o select de OS pela busca (número ou modelo). Esconde as options que
// não batem e os optgroups que ficaram sem nenhuma visível. Se sobrar exatamente
// uma, já a seleciona — o caso comum de digitar o número inteiro da OS.
function _expFiltrarOS() {
  const sel = document.getElementById('ec-os');
  const busca = document.getElementById('ec-os-busca');
  if (!sel || !busca) return;
  const q = _normNome(busca.value);
  let visiveis = 0, unica = null;
  sel.querySelectorAll('option').forEach(opt => {
    if (!opt.value) return; // "— selecione —" sempre fica
    const bate = !q || _normNome(opt.textContent).includes(q);
    opt.hidden = !bate;
    if (bate) { visiveis++; unica = opt; }
  });
  // Some o rótulo do grupo que ficou vazio.
  sel.querySelectorAll('optgroup').forEach(g => {
    const temVisivel = Array.from(g.querySelectorAll('option')).some(o => !o.hidden);
    g.hidden = !temVisivel;
  });
  // Se a OS escolhida sumiu do filtro, limpa a seleção pra não salvar às cegas.
  if (sel.selectedOptions[0] && sel.selectedOptions[0].hidden) sel.value = '';
  // Uma OS só sobrando: seleciona direto.
  if (q && visiveis === 1 && unica) sel.value = unica.value;
  const aviso = document.getElementById('ec-os-vazio');
  if (aviso) aviso.style.display = (q && visiveis === 0) ? 'block' : 'none';
  _expAtualizarSugestaoVolumes();
}

// Mostra as peças da OS e sugere os volumes = nº de tamanhos da grade + 1
// (reposição). Só preenche campo vazio, nunca sobrescreve digitação.
function _expAtualizarSugestaoVolumes() {
  const osId = document.getElementById('ec-os')?.value || '';
  const info = document.getElementById('ec-info');
  const campo = document.getElementById('ec-volumes');
  const o = osId ? (STATE.ordens || []).find(x => x.id === osId) : null;
  if (!o) { if (info) info.textContent = 'Selecione a OS para ver as peças.'; return; }
  const pecas = _expPecasOS(o);
  const nTam = _expTotalTamanhosGrade(o);
  const sug = _expSugestaoVolumes(o);
  if (campo && !campo.value && sug) campo.value = sug;
  if (info) {
    info.innerHTML = `OS <b>${esc(o.os || '—')}</b> · ${esc(o.modeloNome || 'sem modelo')} · <b>${pecas.toLocaleString('pt-BR')} peças</b>.`
      + (nTam > 0
        ? ` Grade com <b>${nTam} tamanho(s)</b> → sugestão de <b>${esc(sug)} volumes</b> (1 pacote por tamanho, por tonalidade, + 1 de reposição).`
        : ' Sem grade com tamanhos definidos — não dá pra sugerir os volumes.');
  }
}

// Preencher a VOLTA a partir de uma OE já montada. A volta quase nunca é uma
// carga nova: é o retorno do que uma ida levou. Montá-la OS por OS repetia à mão
// uma lista que já existe — e qualquer esquecimento vira peça largada na outra
// unidade. Aqui se escolhe a expedição de origem e as OSs dela vêm junto, com os
// mesmos volumes.
function abrirModalExpVolta(janelaId, dataOrig) {
  if (!exigirAdmin('alocar OS na expedição')) return;
  _expModalCtx = { tipo: 'volta', janelaId, dataOrig };

  // Candidatas: qualquer perna de qualquer ocorrência que TENHA carga, exceto a
  // própria volta que está sendo preenchida. Olha 180 dias para trás e 90 para
  // frente — a ida que se quer trazer costuma ser a da véspera ou a da manhã.
  const hoje = _expHoje();
  const ocs = ocorrenciasExpedicao(_expAddDias(hoje, -180), _expAddDias(hoje, 90));
  const origens = [];
  ocs.forEach(oc => {
    ['ida', 'volta'].forEach(p => {
      if (oc.janela.id === janelaId && oc.dataOrig === dataOrig && p === 'volta') return;
      const r = resumoPernaExpedicao(oc, p);
      if (!r.itens.length) return;
      origens.push({
        val: `${oc.janela.id}|${oc.dataOrig}|${p}`,
        data: oc.data,
        label: `${_EXP_DIAS_CURTO[_expData(oc.data).getDay()]} ${formatDate(oc.data)} · ${p === 'ida' ? 'IDA' : 'VOLTA'} · ${oc.janela.nome || 'sem nome'} — ${r.itens.length} OS · ${r.volumes} vol`
      });
    });
  });
  // Mais recente primeiro: a ida a trazer de volta é quase sempre a última.
  origens.sort((a, b) => String(b.data).localeCompare(String(a.data)));

  document.getElementById('modal-exp-title').textContent = 'Preencher a volta a partir de uma OE';
  if (!origens.length) {
    document.getElementById('modal-exp-fields').innerHTML =
      '<div class="info-box">Nenhuma outra expedição com OS alocada para trazer. Monte uma ida primeiro, ou use <b>+ Alocar OS</b> para preencher esta volta manualmente.</div>';
    openModal('modal-exp');
    return;
  }
  document.getElementById('modal-exp-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field full">
        <label>Expedição de origem *</label>
        <select id="ev-origem" onchange="_expVoltaListarOS()">
          ${origens.map((o, i) => `<option value="${esc(o.val)}" ${i === 0 ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}
        </select>
        <div class="field-hint">Traz as OSs desta expedição para a volta em edição, com os mesmos volumes.</div>
      </div>
      <div class="field full">
        <label>OSs a trazer</label>
        <div id="ev-lista" class="ev-lista"></div>
      </div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;" id="ev-info"></div>`;
  _expVoltaListarOS();
  openModal('modal-exp');
}

// Lista as OSs da origem escolhida com caixinha de seleção. As que já estão na
// volta de destino vêm desmarcadas e marcadas como repetidas — trazer de novo
// só duplicaria a linha.
function _expVoltaListarOS() {
  const ctx = _expModalCtx;
  const box = document.getElementById('ev-lista');
  const info = document.getElementById('ev-info');
  if (!ctx || !box) return;
  const [jId, dOrig, perna] = (document.getElementById('ev-origem')?.value || '').split('|');
  // Janela folgada e busca pela data de ORIGEM: uma ocorrência remarcada acontece
  // em outra data, e procurar só pelo dia original a deixaria de fora.
  const oc = ocorrenciasExpedicao(_expAddDias(dOrig, -90), _expAddDias(dOrig, 90))
    .find(o => o.janela.id === jId && o.dataOrig === dOrig);
  const r = oc ? resumoPernaExpedicao(oc, perna) : { itens: [] };
  const destino = (STATE.expedicaoCargas || []).filter(c =>
    c.janelaId === ctx.janelaId && c.data === ctx.dataOrig && c.perna === 'volta');
  const jaLa = new Set(destino.map(c => c.osId));

  box.innerHTML = r.itens.length ? r.itens.map(i => {
    const rep = jaLa.has(i.carga.osId);
    return `
      <label class="ev-item ${rep ? 'rep' : ''}">
        <input type="checkbox" class="ev-os" value="${esc(i.carga.osId)}" data-vol="${i.volumes}" ${rep ? '' : 'checked'}>
        <span class="n">${esc(i.osNumero)}</span>
        <span class="m">${esc(i.modelo) || '—'}</span>
        <span class="v">${i.volumes > 0 ? i.volumes + ' vol' : '— vol'}</span>
        ${rep ? '<span class="exp-badge vazio">já está na volta</span>' : ''}
      </label>`;
  }).join('') : '<div class="exp-vazio">Esta expedição não tem OS alocada.</div>';

  if (info) {
    const novas = r.itens.filter(i => !jaLa.has(i.carga.osId));
    const vol = novas.reduce((s, i) => s + (Number(i.volumes) || 0), 0);
    info.innerHTML = novas.length
      ? `Serão criadas <b>${novas.length}</b> alocação(ões) na volta, somando <b>${vol}</b> volume(s). Desmarque o que não voltar nesta viagem.`
      : 'Todas as OSs desta expedição já estão na volta em edição.';
  }
}

function abrirModalExpConfig() {
  if (!exigirAdmin('configurar a expedição')) return;
  _expModalCtx = { tipo: 'config' };
  const cfg = expCfg();
  document.getElementById('modal-exp-title').textContent = 'Unidades e carga de transporte';
  document.getElementById('modal-exp-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Unidade A (origem da ida) *</label><input type="text" id="ex-uni-a" value="${esc(cfg.unidadeA)}" placeholder="Ex.: Fábrica"></div>
      <div class="field"><label>Unidade B (destino da ida) *</label><input type="text" id="ex-uni-b" value="${esc(cfg.unidadeB)}" placeholder="Ex.: Loja / Depósito"></div>
      ${_expCampoNum('ex-vol-min', 'Volume mínimo padrão', _expNum(cfg.volMin, 0) || '', 'Carga planejada abaixo disso é sinalizada. 0 ou vazio = sem mínimo.')}
      ${_expCampoNum('ex-vol-max', 'Volume máximo padrão', _expNum(cfg.volMax, 0) || '', 'Capacidade do transporte. Acima disso a carga é sinalizada. 0 ou vazio = sem máximo.')}
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">O <b>volume</b> de cada OS é calculado pela grade: <b>1 pacote por tamanho, por tonalidade, + 1 de reposição</b> — cada tonalidade é ensacada separada. Ex.: grade de 7 tamanhos em 1 tom = 8 volumes; a mesma grade em 2 tons = 15.</div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">Os limites valem por <b>perna</b> (ida e volta contam separado) e podem ser sobrescritos em cada janela. A expedição é sempre interna, entre estas duas unidades.</div>`;
  openModal('modal-exp');
}

function abrirModalExpOcorrencia(janelaId, dataOrig) {
  if (!exigirAdmin('cancelar ou remarcar expedições')) return;
  const j = (STATE.expedicaoJanelas || []).find(x => x.id === janelaId);
  if (!j) return;
  _expModalCtx = { tipo: 'ocorrencia', janelaId, dataOrig };
  const exc = (STATE.expedicaoExcecoes || []).find(e => e.janelaId === janelaId && e.data === dataOrig);
  const situacao = exc ? exc.tipo : 'ativa';
  document.getElementById('modal-exp-title').textContent = 'Expedição de ' + formatDate(dataOrig);
  document.getElementById('modal-exp-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field full">
        <label>Situação desta expedição</label>
        <select id="eo-situacao" onchange="_expToggleSituacaoOcorrencia()">
          <option value="ativa" ${situacao === 'ativa' ? 'selected' : ''}>Acontece normalmente</option>
          <option value="cancelada" ${situacao === 'cancelada' ? 'selected' : ''}>Cancelada (não acontece neste dia)</option>
          <option value="remarcada" ${situacao === 'remarcada' ? 'selected' : ''}>Remarcada (muda a data e/ou os horários)</option>
        </select>
      </div>
      <div class="field" id="eo-wrap-data"><label>Nova data *</label><input type="date" id="eo-data" value="${esc((exc && exc.novaData) || dataOrig)}"></div>
      <div class="field" id="eo-wrap-horas">
        <label>Novos horários</label>
        <div style="display:flex;gap:6px;">
          <input type="time" id="eo-hora-ida" value="${esc((exc && exc.horaIda) || j.horaIda || '')}" title="Ida">
          <input type="time" id="eo-hora-volta" value="${esc((exc && exc.horaVolta) || j.horaVolta || '')}" title="Volta">
        </div>
        <div class="field-hint">Ida e volta. Em branco mantém o horário da janela.</div>
      </div>
      <div class="field full"><label>Motivo</label><input type="text" id="eo-motivo" value="${esc((exc && exc.motivo) || '')}" placeholder="Ex.: feriado / veículo em manutenção"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">Muda só <b>este dia</b> — a janela <b>${esc(j.nome) || 'sem nome'}</b> continua valendo nos demais. As OSs já alocadas acompanham a remarcação.</div>`;
  _expToggleSituacaoOcorrencia();
  openModal('modal-exp');
}

function _expToggleSituacaoOcorrencia() {
  const s = document.getElementById('eo-situacao')?.value || 'ativa';
  const remarcada = s === 'remarcada';
  document.getElementById('eo-wrap-data')?.classList.toggle('hidden', !remarcada);
  document.getElementById('eo-wrap-horas')?.classList.toggle('hidden', !remarcada);
}

async function salvarModalExpedicao() {
  if (!_expModalCtx) return;
  const v = id => document.getElementById(id)?.value || '';
  const ctx = _expModalCtx;

  if (ctx.tipo === 'janela') {
    if (!exigirAdmin('cadastrar janelas de expedição')) return;
    const nome = v('ej-nome').trim();
    if (!nome) return toast('Informe o nome da janela', 'err');
    const tipo = v('ej-tipo') === 'data' ? 'data' : 'semanal';
    const diasSemana = Array.from(document.querySelectorAll('.ej-dia:checked')).map(el => Number(el.value));
    if (tipo === 'semanal' && !diasSemana.length) return toast('Marque ao menos um dia da semana', 'err');
    const data = v('ej-data');
    if (tipo === 'data' && !data) return toast('Informe a data da expedição', 'err');
    const horaIda = v('ej-hora-ida'), horaVolta = v('ej-hora-volta');
    if (!horaIda || !horaVolta) return toast('Informe os horários de ida e de volta', 'err');
    const volMin = v('ej-vol-min') === '' ? '' : (parseInt(v('ej-vol-min')) || 0);
    const volMax = v('ej-vol-max') === '' ? '' : (parseInt(v('ej-vol-max')) || 0);
    if (volMin !== '' && volMax !== '' && volMax > 0 && volMin > volMax) {
      return toast('O volume mínimo não pode ser maior que o máximo', 'err');
    }
    const reg = {
      nome, tipo, diasSemana: tipo === 'semanal' ? diasSemana : [],
      data: tipo === 'data' ? data : '',
      horaIda, horaVolta, volMin, volMax,
      ativo: v('ej-ativo') !== '0',
      obs: v('ej-obs').trim()
    };
    if (!Array.isArray(STATE.expedicaoJanelas)) STATE.expedicaoJanelas = [];
    if (ctx.editId) {
      const i = STATE.expedicaoJanelas.findIndex(x => x.id === ctx.editId);
      if (i >= 0) STATE.expedicaoJanelas[i] = { ...STATE.expedicaoJanelas[i], ...reg };
    } else {
      STATE.expedicaoJanelas.push({ id: uid(), ...reg });
    }
    await saveState('expedicaoJanelas');
    toast(ctx.editId ? 'Janela atualizada' : 'Janela cadastrada', 'ok');

  } else if (ctx.tipo === 'carga') {
    if (!exigirAdmin('alocar OS na expedição')) return;
    const [janelaId, data, perna] = v('ec-ocorrencia').split('|');
    if (!janelaId || !data || !perna) return toast('Selecione a expedição', 'err');
    const osId = v('ec-os');
    if (!osId) return toast('Selecione a OS', 'err');
    const volumes = parseInt(v('ec-volumes')) || 0;
    if (!(volumes > 0)) return toast('Informe quantos volumes esta OS ocupa', 'err');
    if (!Array.isArray(STATE.expedicaoCargas)) STATE.expedicaoCargas = [];
    // Ao mover, a propria carga nao conta como duplicata dela mesma.
    const jaTem = STATE.expedicaoCargas.some(c => c.id !== ctx.editId
      && c.janelaId === janelaId && c.data === data && c.perna === perna && c.osId === osId);
    if (jaTem) return toast('Esta OS já está nesta carga', 'err');
    const campos = { janelaId, data, perna, osId, volumes, obs: v('ec-obs').trim() };
    if (ctx.editId) {
      const i = STATE.expedicaoCargas.findIndex(c => c.id === ctx.editId);
      if (i >= 0) STATE.expedicaoCargas[i] = { ...STATE.expedicaoCargas[i], ...campos };
    } else {
      STATE.expedicaoCargas.push({ id: uid(), ...campos });
    }
    await saveState('expedicaoCargas');
    toast(ctx.editId ? 'Expedição da OS alterada' : 'OS alocada na expedição', 'ok');

  } else if (ctx.tipo === 'volta') {
    if (!exigirAdmin('alocar OS na expedição')) return;
    const marcadas = Array.from(document.querySelectorAll('.ev-os:checked'));
    if (!marcadas.length) return toast('Marque ao menos uma OS para trazer', 'err');
    if (!Array.isArray(STATE.expedicaoCargas)) STATE.expedicaoCargas = [];
    const jaLa = new Set(STATE.expedicaoCargas
      .filter(c => c.janelaId === ctx.janelaId && c.data === ctx.dataOrig && c.perna === 'volta')
      .map(c => c.osId));
    let n = 0;
    marcadas.forEach(el => {
      const osId = el.value;
      if (!osId || jaLa.has(osId)) return;   // repetida: a checagem no salvar também vale
      jaLa.add(osId);
      STATE.expedicaoCargas.push({
        id: uid(), janelaId: ctx.janelaId, data: ctx.dataOrig, perna: 'volta',
        osId, volumes: parseInt(el.dataset.vol, 10) || 0, obs: ''
      });
      n++;
    });
    if (!n) return toast('Essas OSs já estão na volta', 'err');
    await saveState('expedicaoCargas');
    toast(`${n} OS trazida(s) para a volta`, 'ok');

  } else if (ctx.tipo === 'config') {
    if (!exigirAdmin('configurar a expedição')) return;
    const unidadeA = v('ex-uni-a').trim(), unidadeB = v('ex-uni-b').trim();
    if (!unidadeA || !unidadeB) return toast('Informe o nome das duas unidades', 'err');
    const volMin = parseInt(v('ex-vol-min')) || 0;
    const volMax = parseInt(v('ex-vol-max')) || 0;
    if (volMax > 0 && volMin > volMax) return toast('O volume mínimo não pode ser maior que o máximo', 'err');
    if (!STATE.meta || typeof STATE.meta !== 'object') STATE.meta = {};
    STATE.meta.expedicao = { ...(STATE.meta.expedicao || {}), unidadeA, unidadeB, volMin, volMax };
    await saveState('meta');
    toast('Configuração salva', 'ok');

  } else if (ctx.tipo === 'ocorrencia') {
    if (!exigirAdmin('cancelar ou remarcar expedições')) return;
    const situacao = v('eo-situacao');
    if (!Array.isArray(STATE.expedicaoExcecoes)) STATE.expedicaoExcecoes = [];
    STATE.expedicaoExcecoes = STATE.expedicaoExcecoes.filter(e => !(e.janelaId === ctx.janelaId && e.data === ctx.dataOrig));
    if (situacao === 'cancelada') {
      STATE.expedicaoExcecoes.push({ id: uid(), janelaId: ctx.janelaId, data: ctx.dataOrig, tipo: 'cancelada', motivo: v('eo-motivo').trim() });
    } else if (situacao === 'remarcada') {
      const novaData = v('eo-data');
      if (!novaData) return toast('Informe a nova data', 'err');
      STATE.expedicaoExcecoes.push({
        id: uid(), janelaId: ctx.janelaId, data: ctx.dataOrig, tipo: 'remarcada',
        novaData, horaIda: v('eo-hora-ida'), horaVolta: v('eo-hora-volta'), motivo: v('eo-motivo').trim()
      });
    }
    await saveState('expedicaoExcecoes');
    toast(situacao === 'ativa' ? 'Expedição restabelecida' : (situacao === 'cancelada' ? 'Expedição cancelada' : 'Expedição remarcada'), 'ok');
  }

  closeModal('modal-exp');
  _expModalCtx = null;
  renderExpedicaoPlano();
}

async function excluirCargaExp(id) {
  if (!exigirAdmin('remover OS da expedição')) return;
  if (!confirm('Tirar esta OS da carga?')) return;
  STATE.expedicaoCargas = (STATE.expedicaoCargas || []).filter(c => c.id !== id);
  await saveState('expedicaoCargas');
  toast('OS removida da carga', 'ok');
  renderExpedicaoPlano();
}

// Redefine os volumes das cargas FUTURAS pela regra da grade (tamanhos + 1).
// Corrige valores gravados por regras antigas (ex.: OS que ficou com 600).
// Não toca em expedição já realizada — aquilo é histórico do que saiu.
async function recalcularVolumesExpedicao() {
  if (!exigirAdmin('recalcular volumes')) return;
  const hoje = _expHoje();
  let n = 0;
  (STATE.expedicaoCargas || []).forEach(c => {
    if (_expDataEfetivaCarga(c) < hoje) return; // já realizada: não reescreve
    const o = (STATE.ordens || []).find(x => x.id === c.osId);
    const sug = Number(_expSugestaoVolumes(o)) || 0;
    if (sug > 0 && sug !== (Number(c.volumes) || 0)) { c.volumes = sug; n++; }
  });
  if (n) {
    await saveState('expedicaoCargas');
    toast(`${n} carga(s) recalculada(s) pela grade`, 'ok');
  } else {
    toast('Nada a recalcular — volumes já batem com a grade', '');
  }
  renderExpedicaoPlano();
}

async function excluirJanelaExp(id) {
  if (!exigirAdmin('excluir janelas de expedição')) return;
  const j = (STATE.expedicaoJanelas || []).find(x => x.id === id);
  if (!j) return;
  const cargas = (STATE.expedicaoCargas || []).filter(c => c.janelaId === id).length;
  const aviso = cargas
    ? `Excluir a janela "${j.nome || 'sem nome'}"?\n\n${cargas} alocação(ões) de OS serão perdidas junto — inclusive as de expedições já realizadas.`
    : `Excluir a janela "${j.nome || 'sem nome'}"?`;
  if (!confirm(aviso)) return;
  STATE.expedicaoJanelas = (STATE.expedicaoJanelas || []).filter(x => x.id !== id);
  STATE.expedicaoCargas = (STATE.expedicaoCargas || []).filter(c => c.janelaId !== id);
  STATE.expedicaoExcecoes = (STATE.expedicaoExcecoes || []).filter(e => e.janelaId !== id);
  await saveState('expedicaoJanelas');
  await saveState('expedicaoCargas');
  await saveState('expedicaoExcecoes');
  toast('Janela excluída', 'ok');
  renderExpedicaoPlano();
}

/* ========================================================= */
/*      PLANEJAMENTO DIÁRIO DE OPERAÇÕES (por função)        */
/* ========================================================= */
// O campo "Operações" planeja a JORNADA de cada posto de trabalho, não tarefa
// por tarefa. Uma operação aqui é o processo COMPLETO de uma função no dia —
// "o operador de enfestadeira começa 07:12 e leva 3h20" já engloba todas as
// etapas internas dele e o tempo total até concluir. Por isso não há vínculo
// obrigatório com OS: o que se planeja é o tempo do posto, e o pedido/lote
// entra só como referência em texto quando faz sentido.
//
// As funções correm em PARALELO: cada uma tem a sua faixa no dia, e a barra de
// tempo desenhada na janela comum do dia é o que deixa ver quem começa quando e
// onde os postos se cruzam.
//
// Reaproveita os helpers de data do planejamento de expedição (_expIso,
// _expData, _expHoje, _expAddDias, _expRange, _expLabelPeriodo).

const _OP_STATUS = {
  pendente:  { lbl: 'Pendente',     cls: 'baixo' },
  andamento: { lbl: 'Em andamento', cls: 'info' },
  feita:     { lbl: 'Feita',        cls: 'ok' }
};
// Clicar no status roda o ciclo pendente → andamento → feita → pendente.
const _OP_CICLO = { pendente: 'andamento', andamento: 'feita', feita: 'pendente' };

// Classificação da operação. "Eletiva" é o padrão — a operação programada, que
// é a maioria — e por isso não ganha selo na linha: poluir a agenda inteira com
// o rótulo do caso comum esconderia justamente o que precisa saltar aos olhos.
const _OP_PRIORIDADE = {
  urgente:   { lbl: 'Urgente' },
  emergente: { lbl: 'Emergente' },
  eletiva:   { lbl: 'Eletiva' }
};
function _opPrioridade(op) { return _OP_PRIORIDADE[op.prioridade] ? op.prioridade : 'eletiva'; }

let opPlanoModo = 'dia';           // o planejamento é DIÁRIO por natureza
let opPlanoAncora = _expHoje();
try {
  opPlanoModo = sessionStorage.getItem('gos:op:modo') || opPlanoModo;
  opPlanoAncora = sessionStorage.getItem('gos:op:ancora') || opPlanoAncora;
} catch (e) { /* sessionStorage indisponível, segue no padrão */ }

function opSetModo(modo) {
  opPlanoModo = modo;
  try { sessionStorage.setItem('gos:op:modo', modo); } catch (e) {}
  renderOperacoes();
}
function opNav(dir) {
  const passo = opPlanoModo === 'dia' ? 1 : (opPlanoModo === 'semana' ? 7 : 0);
  if (passo) opPlanoAncora = _expAddDias(opPlanoAncora, dir * passo);
  else {
    const d = _expData(opPlanoAncora);
    opPlanoAncora = _expIso(new Date(d.getFullYear(), d.getMonth() + dir, 1));
  }
  try { sessionStorage.setItem('gos:op:ancora', opPlanoAncora); } catch (e) {}
  renderOperacoes();
}
function opHoje() {
  opPlanoAncora = _expHoje();
  try { sessionStorage.setItem('gos:op:ancora', opPlanoAncora); } catch (e) {}
  renderOperacoes();
}

/* ---------------- tempo ---------------- */

// 'HH:MM' → minutos desde a meia-noite. Vazio/inválido vira null (a operação
// existe mesmo sem horário definido — fica listada como "sem horário").
function _opMin(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || '').trim());
  if (!m) return null;
  const h = Number(m[1]), mi = Number(m[2]);
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}
// Minutos → 'HH:MM'. Passando da meia-noite marca o dia seguinte: uma jornada
// que atravessa a virada é real (turno da noite) e não pode virar '01:00' seco.
function _opHHMM(min) {
  const v = Math.max(0, Math.round(Number(min) || 0));
  const dias = Math.floor(v / 1440);
  const t = v % 1440;
  return String(Math.floor(t / 60)).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0')
    + (dias ? ` (+${dias}d)` : '');
}
// Minutos → '3h20', '2h', '45min'. É a leitura de duração do chão de fábrica.
function _opDurTexto(min) {
  const v = Math.max(0, Math.round(Number(min) || 0));
  if (!v) return '—';
  const h = Math.floor(v / 60), m = v % 60;
  if (!h) return m + 'min';
  return h + 'h' + (m ? String(m).padStart(2, '0') : '');
}
function _opDuracao(op) { return Math.max(0, Math.round(Number(op.duracaoMin) || 0)); }
function _opInicioMin(op) { return _opMin(op.inicio); }
// Término = início + duração. null quando a operação não tem horário.
function _opFimMin(op) {
  const ini = _opInicioMin(op);
  return ini == null ? null : ini + _opDuracao(op);
}
// Janela da operação em texto: '07:12 → 10:32 · 3h20'.
function _opJanelaTexto(op) {
  const ini = _opInicioMin(op);
  const dur = _opDuracao(op);
  if (ini == null) return dur ? `sem horário · ${_opDurTexto(dur)}` : 'sem horário';
  return `${_opHHMM(ini)} → ${_opHHMM(ini + dur)}` + (dur ? ` · ${_opDurTexto(dur)}` : '');
}

/* ---------------- cadastros ligados ---------------- */

// Função cadastrada de uma operação. Cai no nome copiado no registro quando a
// função foi excluída do cadastro — o dia planejado não pode virar linha órfã.
function _opFuncaoNome(op) {
  const f = (STATE.funcoes || []).find(x => x.id === op.funcaoId);
  return f ? f.nome : (op.funcaoNome || '(função excluída)');
}
function _opResponsavelNome(op) {
  const p = (STATE.equipe || []).find(x => x.id === op.responsavelId);
  return p ? p.nome : (op.responsavelNome || '');
}
// Pessoas da equipe cuja função principal é esta. É o que faz o select de
// responsável mostrar antes quem realmente ocupa aquele posto.
function _opPessoasDaFuncao(funcaoNome) {
  const alvo = _normFuncaoNome(funcaoNome);
  const dentro = [], fora = [];
  (STATE.equipe || []).forEach(p => {
    (alvo && _normFuncaoNome(p.funcao) === alvo ? dentro : fora).push(p);
  });
  const ord = arr => arr.sort((a, b) => String(a.nome || '').localeCompare(String(b.nome || '')));
  return { dentro: ord(dentro), fora: ord(fora) };
}
// Sugestões de operação para uma função: as responsabilidades cadastradas nela
// (uma por linha em Funções) + as etapas de produção. O nome da operação é o do
// processo inteiro do posto — as sugestões só evitam redigitar tudo todo dia.
function _opSugestoesOperacao(funcaoId) {
  const f = (STATE.funcoes || []).find(x => x.id === funcaoId);
  const acoes = String(f && f.acoes || '').split('\n').map(s => s.trim()).filter(Boolean);
  const etapas = etapasOrdenadas().map(e => e.nome).filter(Boolean);
  return [...new Set([...acoes, ...etapas])];
}

function _opStatus(op) { return _OP_STATUS[op.status] ? op.status : 'pendente'; }

// Ordem de exibição DENTRO de um dia. A ordem manual (campo `ordem`, gravado
// quando o usuário move a operação) manda; quem nunca foi movido cai no
// horário de início e, sem horário, no nome da função. Assim o dia recém-criado
// já sai numa ordem sensata e continua reordenável à mão depois.
function _opCompararNoDia(a, b) {
  const oa = Number.isFinite(Number(a.ordem)) ? Number(a.ordem) : null;
  const ob = Number.isFinite(Number(b.ordem)) ? Number(b.ordem) : null;
  if (oa != null && ob != null && oa !== ob) return oa - ob;
  if (oa != null && ob == null) return -1;
  if (ob != null && oa == null) return 1;
  const ia = _opInicioMin(a), ib = _opInicioMin(b);
  if (ia == null && ib != null) return 1;
  if (ib == null && ia != null) return -1;
  if (ia != null && ib != null && ia !== ib) return ia - ib;
  return _opFuncaoNome(a).localeCompare(_opFuncaoNome(b));
}

// Operações do período, ordenadas por data e, dentro do dia, pela ordem de
// exibição acima.
function operacoesNoPeriodo(ini, fim) {
  return (STATE.operacoes || [])
    .filter(o => o.data && o.data >= ini && o.data <= fim)
    .sort((a, b) => String(a.data).localeCompare(String(b.data)) || _opCompararNoDia(a, b));
}

/* ---------------- ordem manual ---------------- */

// O dia visto como uma sequência de BLOCOS (um por função), cada um com os seus
// itens já na ordem de exibição. É a estrutura que a tela desenha e também a que
// as setas de mover manipulam — as duas leem o mesmo arranjo, então o que se vê
// é exatamente o que se move.
function _opBlocosDoDia(data) {
  const doDia = (STATE.operacoes || []).filter(o => o.data === data).sort(_opCompararNoDia);
  const blocos = [];
  const idx = new Map();
  doDia.forEach(op => {
    const nome = _opFuncaoNome(op);
    if (!idx.has(nome)) { idx.set(nome, blocos.length); blocos.push({ nome, itens: [] }); }
    blocos[idx.get(nome)].itens.push(op);
  });
  return blocos;
}

// Grava `ordem` 0..n na sequência dada (blocos achatados). Renumerar tudo a cada
// movimento mantém os blocos contíguos na numeração, que é o que permite mover
// um posto inteiro trocando dois trechos vizinhos.
function _opGravarOrdem(blocos) {
  let n = 0;
  blocos.forEach(b => b.itens.forEach(op => { op.ordem = n++; }));
}

// Move uma operação para cima/baixo DENTRO do seu posto.
async function moverOperacao(id, dir) {
  if (!exigirAdmin('reordenar operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  const blocos = _opBlocosDoDia(op.data);
  const bloco = blocos.find(b => b.itens.some(x => x.id === id));
  if (!bloco) return;
  const i = bloco.itens.findIndex(x => x.id === id);
  const j = i + dir;
  if (j < 0 || j >= bloco.itens.length) return;
  [bloco.itens[i], bloco.itens[j]] = [bloco.itens[j], bloco.itens[i]];
  _opGravarOrdem(blocos);
  await saveState('operacoes');
  renderOperacoes();
}

// Move um POSTO inteiro para cima/baixo no dia, levando junto as operações dele.
async function moverPostoOperacoes(data, funcaoNome, dir) {
  if (!exigirAdmin('reordenar operações')) return;
  const blocos = _opBlocosDoDia(data);
  const i = blocos.findIndex(b => b.nome === funcaoNome);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= blocos.length) return;
  [blocos[i], blocos[j]] = [blocos[j], blocos[i]];
  _opGravarOrdem(blocos);
  await saveState('operacoes');
  renderOperacoes();
}

// Operações da MESMA função que se sobrepõem no tempo. O posto é um só: duas
// jornadas cruzadas no mesmo operador é erro de planejamento, e é o aviso que
// mais importa numa agenda de tempo.
function _opConflitos(lista) {
  const ids = new Set();
  const porFuncao = new Map();
  lista.forEach(op => {
    if (_opInicioMin(op) == null || !_opDuracao(op)) return;
    const k = op.funcaoId || _opFuncaoNome(op);
    if (!porFuncao.has(k)) porFuncao.set(k, []);
    porFuncao.get(k).push(op);
  });
  porFuncao.forEach(arr => {
    arr.sort((a, b) => _opInicioMin(a) - _opInicioMin(b));
    for (let i = 1; i < arr.length; i++) {
      if (_opInicioMin(arr[i]) < _opFimMin(arr[i - 1])) { ids.add(arr[i].id); ids.add(arr[i - 1].id); }
    }
  });
  return ids;
}

/* ---------------- render da agenda ---------------- */

// Janela de tempo comum do dia (minutos), arredondada para horas cheias. É o
// eixo em que TODAS as faixas são desenhadas — é a base comum que deixa comparar
// os postos entre si.
function _opJanelaDoDia(ops) {
  let ini = null, fim = null;
  ops.forEach(op => {
    const i = _opInicioMin(op);
    if (i == null) return;
    const f = i + _opDuracao(op);
    if (ini == null || i < ini) ini = i;
    if (fim == null || f > fim) fim = f;
  });
  if (ini == null) return null;
  ini = Math.floor(ini / 60) * 60;
  fim = Math.ceil(fim / 60) * 60;
  if (fim - ini < 240) fim = ini + 240;   // no mínimo 4h de eixo, senão as barras ficam sem escala
  return { ini, fim };
}

function renderOperacoes() {
  const cont = document.getElementById('operacoes-painel');
  if (!cont) return;
  const { ini, fim } = _expRange(opPlanoModo, opPlanoAncora);
  const ops = operacoesNoPeriodo(ini, fim);
  const fmt = n => (Number(n) || 0).toLocaleString('pt-BR');

  const toolbar = `
    <div class="exp-toolbar no-print">
      <div class="exp-seg">
        <button class="${opPlanoModo === 'dia' ? 'active' : ''}" onclick="opSetModo('dia')">Diário</button>
        <button class="${opPlanoModo === 'semana' ? 'active' : ''}" onclick="opSetModo('semana')">Semanal</button>
        <button class="${opPlanoModo === 'mes' ? 'active' : ''}" onclick="opSetModo('mes')">Mensal</button>
      </div>
      <div style="display:flex;align-items:center;gap:6px;">
        <button class="btn" onclick="opNav(-1)" title="Período anterior">‹</button>
        <div class="exp-periodo">${esc(_expLabelPeriodo(opPlanoModo, opPlanoAncora))}</div>
        <button class="btn" onclick="opNav(1)" title="Próximo período">›</button>
        <button class="btn" onclick="opHoje()">Hoje</button>
      </div>
      <div class="admin-only" style="display:flex;gap:6px;flex-wrap:wrap;">
        <button class="btn primary" onclick="abrirModalOperacao()">+ Nova operação</button>
        ${opPlanoModo === 'dia' ? `<button class="btn" onclick="copiarOperacoesDoDiaAnterior()" title="Repete no dia mostrado a jornada do último dia planejado antes dele. Copia como pendente e não duplica o que já existe.">⧉ Repetir dia anterior</button>` : ''}
      </div>
    </div>`;

  const conflitos = _opConflitos(ops);
  let minutos = 0, pendentes = 0, feitas = 0, prioritarias = 0;
  const funcoesSet = new Set();
  ops.forEach(o => {
    minutos += _opDuracao(o);
    if (_opStatus(o) === 'feita') feitas++; else pendentes++;
    if (_opPrioridade(o) !== 'eletiva') prioritarias++;
    funcoesSet.add(_opFuncaoNome(o));
  });

  const resumo = `
    <div class="exp-resumo">
      <div class="item"><div class="num">${fmt(ops.length)}</div><div class="lbl">Operações no período</div></div>
      <div class="item"><div class="num">${fmt(funcoesSet.size)}</div><div class="lbl">Postos / funções</div></div>
      <div class="item"><div class="num">${esc(_opDurTexto(minutos))}</div><div class="lbl">Tempo planejado</div></div>
      <div class="item ${pendentes ? 'alerta' : ''}"><div class="num">${fmt(pendentes)}</div><div class="lbl">A executar</div></div>
      <div class="item ${prioritarias ? 'alerta' : ''}"><div class="num">${fmt(prioritarias)}</div><div class="lbl">Urgentes / emergentes</div></div>
      <div class="item"><div class="num">${fmt(feitas)}</div><div class="lbl">Concluídas</div></div>
      <div class="item ${conflitos.size ? 'alerta' : ''}"><div class="num">${fmt(conflitos.size)}</div><div class="lbl">Em sobreposição</div></div>
    </div>`;

  // Régua de horas do dia: o eixo em que as barras são lidas. O passo cresce
  // junto com a janela para os rótulos nunca se encavalarem.
  const reguaHtml = jan => {
    const larg = jan.fim - jan.ini;
    const passo = larg <= 480 ? 60 : (larg <= 960 ? 120 : 180);
    const ticks = [];
    for (let m = jan.ini; m <= jan.fim; m += passo) {
      const pos = (m - jan.ini) / larg * 100;
      // Os rótulos das pontas encostam na borda: centralizá-los cortaria metade
      // do texto para fora da faixa.
      const ponta = pos < 1 ? 'ini' : (pos > 99 ? 'fim' : '');
      const anc = ponta === 'ini' ? 'translateX(0)' : (ponta === 'fim' ? 'translateX(-100%)' : 'translateX(-50%)');
      const mk = ponta === 'ini' ? '0' : (ponta === 'fim' ? '100%' : '50%');
      ticks.push(`<span class="op-tick" style="left:${pos.toFixed(3)}%;transform:${anc};--mk:${mk}">${esc(_opHHMM(m))}</span>`);
    }
    return `<div class="op-regua"><div class="op-regua-lbl">Horário do dia</div><div class="op-regua-eixo">${ticks.join('')}</div></div>`;
  };

  // Barra da operação dentro da janela do dia. É onde "07:12 por 3h20" vira
  // uma coisa que se enxerga ao lado dos outros postos.
  const barraHtml = (op, jan) => {
    const i = _opInicioMin(op), dur = _opDuracao(op);
    if (i == null || !dur) return '';
    const larg = jan.fim - jan.ini;
    const left = (i - jan.ini) / larg * 100;
    const width = Math.max(1.2, dur / larg * 100);
    const st = _opStatus(op);
    return `<div class="op-bar ${st}" style="left:${left.toFixed(3)}%;width:${width.toFixed(3)}%"
      title="${esc(op.operacao)} · ${esc(_opJanelaTexto(op))}"><span>${esc(op.operacao)}</span></div>`;
  };

  const linhaHtml = (op, pos, qtd) => {
    const st = _opStatus(op);
    const pr = _opPrioridade(op);
    const resp = _opResponsavelNome(op);
    const conflito = conflitos.has(op.id);
    // O selo distingue as duas naturezas: processo inteiro do posto (o padrão,
    // sem selo) e etapa avulsa planejada à parte.
    const selo = op.escopo === 'etapa'
      ? ` <span class="exp-badge info" title="Planejada como etapa isolada da função${op.etapa && op.etapa !== op.operacao ? ': ' + op.etapa : ''}">etapa</span>`
      : '';
    const selopr = pr === 'eletiva' ? '' : ` <span class="op-prio ${pr}">${esc(_OP_PRIORIDADE[pr].lbl)}</span>`;
    return `
      <div class="op-row prio-${pr} ${st === 'feita' ? 'feita' : ''}">
        <span class="admin-only op-mover">
          <button title="Subir esta operação no posto" onclick="moverOperacao('${esc(op.id)}',-1)" ${pos === 0 ? 'disabled' : ''}>▲</button>
          <button title="Descer esta operação no posto" onclick="moverOperacao('${esc(op.id)}',1)" ${pos === qtd - 1 ? 'disabled' : ''}>▼</button>
        </span>
        <span class="janela">${esc(_opJanelaTexto(op))}</span>
        <span class="oper">${esc(op.operacao) || '(sem descrição)'}${selopr}${selo}${conflito ? ' <span class="exp-badge alto" title="Este posto tem outra operação no mesmo horário">sobreposta</span>' : ''}${op.obs ? ` <span class="obs">· ${esc(op.obs)}</span>` : ''}</span>
        <span class="resp">${esc(resp) || '<span class="obs">a definir</span>'}</span>
        <span class="ref">${esc(op.referencia) || ''}</span>
        <button type="button" class="exp-badge ${_OP_STATUS[st].cls} op-status" onclick="alternarStatusOperacao('${esc(op.id)}')" title="Clique para mudar: pendente → em andamento → feita">${esc(_OP_STATUS[st].lbl)}</button>
        <span class="admin-only op-acoes">
          <button title="Editar esta operação" onclick="abrirModalOperacao('${esc(op.id)}')">✎</button>
          <button title="Excluir esta operação" onclick="excluirOperacao('${esc(op.id)}')">×</button>
        </span>
      </div>`;
  };

  // Um bloco por FUNÇÃO dentro do dia: as funções correm em paralelo, então
  // cada uma tem a sua faixa própria no mesmo eixo de horas.
  const diaHtml = (data, doDia) => {
    const jan = _opJanelaDoDia(doDia);
    // A mesma estrutura que as setas de mover manipulam — desenhar a partir dela
    // garante que a ordem vista é a ordem gravada.
    const grupos = _opBlocosDoDia(data);

    const blocos = grupos.map((g, gi) => {
      const minutos = g.itens.reduce((s, o) => s + _opDuracao(o), 0);
      const pend = g.itens.filter(o => _opStatus(o) !== 'feita').length;
      const comHora = g.itens.filter(o => _opInicioMin(o) != null);
      const jIni = comHora.length ? Math.min(...comHora.map(_opInicioMin)) : null;
      const jFim = comHora.length ? Math.max(...comHora.map(_opFimMin)) : null;
      return `
        <div class="op-func">
          <div class="op-func-head">
            <div class="op-func-nome">
              <span class="admin-only op-mover">
                <button title="Subir este posto no dia" onclick="moverPostoOperacoes('${esc(data)}','${esc(g.nome).replace(/'/g, '&#39;')}',-1)" ${gi === 0 ? 'disabled' : ''}>▲</button>
                <button title="Descer este posto no dia" onclick="moverPostoOperacoes('${esc(data)}','${esc(g.nome).replace(/'/g, '&#39;')}',1)" ${gi === grupos.length - 1 ? 'disabled' : ''}>▼</button>
              </span>
              ${esc(g.nome)}
            </div>
            <div class="op-func-tot">
              ${jIni != null ? `<b>${esc(_opHHMM(jIni))} → ${esc(_opHHMM(jFim))}</b> · ` : ''}${esc(_opDurTexto(minutos))} de operação
              ${g.itens.length > 1 ? ` · ${g.itens.length} blocos` : ''}
              ${pend ? ` · <span class="exp-badge baixo">${pend} a fazer</span>` : ' · <span class="exp-badge ok">tudo feito</span>'}
            </div>
          </div>
          ${jan ? `<div class="op-faixa"><div class="op-faixa-eixo">${g.itens.map(op => barraHtml(op, jan)).join('')}</div></div>` : ''}
          ${g.itens.map((op, i) => linhaHtml(op, i, g.itens.length)).join('')}
        </div>`;
    }).join('');

    const totMin = doDia.reduce((s, o) => s + _opDuracao(o), 0);
    const comHora = doDia.filter(o => _opInicioMin(o) != null);
    const abre = comHora.length ? Math.min(...comHora.map(_opInicioMin)) : null;
    const fecha = comHora.length ? Math.max(...comHora.map(_opFimMin)) : null;
    const prioridades = ['emergente', 'urgente']
      .map(p => ({ p, n: doDia.filter(o => _opPrioridade(o) === p).length }))
      .filter(x => x.n)
      .map(x => ` · <span class="op-prio ${x.p}">${x.n} ${esc(_OP_PRIORIDADE[x.p].lbl.toLowerCase())}${x.n > 1 ? 's' : ''}</span>`).join('');
    return `
      <div class="card exp-ocor">
        <div class="exp-ocor-head">
          <div>
            <div class="exp-ocor-data">${_EXP_DIAS_CURTO[_expData(data).getDay()]} · ${esc(formatDate(data))}${data === _expHoje() ? ' <span class="exp-badge info">hoje</span>' : ''}</div>
            <div class="exp-ocor-nome">
              ${abre != null ? `Jornada <b>${esc(_opHHMM(abre))} → ${esc(_opHHMM(fecha))}</b> · ` : ''}${grupos.length} ${grupos.length === 1 ? 'posto' : 'postos'} em paralelo · ${esc(_opDurTexto(totMin))} de operação somados${prioridades}
            </div>
          </div>
          <div class="admin-only">
            <button class="btn" onclick="abrirModalOperacao('','${esc(data)}')">+ Operação neste dia</button>
          </div>
        </div>
        ${jan ? reguaHtml(jan) : ''}
        ${blocos}
      </div>`;
  };

  const porDia = new Map();
  ops.forEach(op => {
    if (!porDia.has(op.data)) porDia.set(op.data, []);
    porDia.get(op.data).push(op);
  });
  const cards = Array.from(porDia.keys()).sort().map(d => diaHtml(d, porDia.get(d))).join('');

  const semFuncoes = !(STATE.funcoes || []).length;
  const vazio = `
    <div class="card">
      <div class="empty" style="padding:24px 0;text-align:center;">
        ${semFuncoes
          ? 'Nenhuma <b>função</b> cadastrada ainda. O planejamento é feito por posto de trabalho — cadastre as funções em <a href="#" onclick="goto(\'cad-funcoes\'); return false;">Funções</a> antes de começar.'
          : 'Nenhuma operação planejada neste período. Use <b>+ Nova operação</b> para montar o dia.'}
      </div>
    </div>`;

  const comoFunciona = `
    <div class="info-box no-print" style="font-size:12px;">
      Planeje a <b>jornada de cada posto</b>: a operação é o processo completo daquela função —
      informar que a enfestadeira começa às <b>07:12</b> e leva <b>3h20</b> já engloba todas as etapas
      internas e o tempo total até concluir. As funções correm <b>em paralelo</b>, cada uma na sua faixa
      do mesmo eixo de horas. Clique no status para ir de <b>pendente</b> a <b>em andamento</b> e a <b>feita</b>.
    </div>`;

  cont.innerHTML = toolbar + comoFunciona + resumo + (cards || vazio);
}

/* ---------------- modal da operação ---------------- */

let _opModalCtx = null;

function abrirModalOperacao(opId = '', dataPre = '', funcaoIdPre = '') {
  if (!exigirAdmin('planejar operações')) return;
  if (!(STATE.funcoes || []).length) {
    return toast('Cadastre ao menos uma função antes de planejar operações', 'err');
  }
  const op = opId ? (STATE.operacoes || []).find(x => x.id === opId) : null;
  _opModalCtx = { editId: op ? opId : '' };

  const data = op ? (op.data || '') : (dataPre || opPlanoAncora || _expHoje());
  const funcaoSel = op ? (op.funcaoId || '') : funcaoIdPre;
  const funcoesOpts = (STATE.funcoes || [])
    .slice().sort((a, b) => String(a.nome || '').localeCompare(String(b.nome || '')))
    .map(f => `<option value="${esc(f.id)}" ${f.id === funcaoSel ? 'selected' : ''}>${esc(f.nome)}</option>`).join('');

  const dur = op ? _opDuracao(op) : 0;
  const durH = op ? Math.floor(dur / 60) : '';
  const durM = op ? (dur % 60) : '';
  const escopo = op && op.escopo === 'etapa' ? 'etapa' : 'completa';

  const statusOpts = Object.entries(_OP_STATUS).map(([k, v]) =>
    `<option value="${k}" ${op && _opStatus(op) === k ? 'selected' : ''}>${esc(v.lbl)}</option>`).join('');
  const prioAtual = op ? _opPrioridade(op) : 'eletiva';
  const prioridadeOpts = Object.entries(_OP_PRIORIDADE).map(([k, v]) =>
    `<option value="${k}" ${prioAtual === k ? 'selected' : ''}>${esc(v.lbl)}</option>`).join('');

  // Sugestões de referência: números das OS que estão no fluxo. É só atalho de
  // digitação — o campo é livre e aceita lote, coleção, o que o dia pedir.
  const refs = (STATE.ordens || [])
    .filter(o => faseAtualOS(o) >= 0 && (o.os || '').toString().trim())
    .map(o => `${o.os}${o.modeloNome ? ' · ' + o.modeloNome : ''}`)
    .sort((a, b) => String(b).localeCompare(String(a), undefined, { numeric: true }))
    .slice(0, 60);

  document.getElementById('modal-op-title').textContent = op ? 'Editar operação do dia' : 'Nova operação do dia';
  document.getElementById('modal-op-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Data *</label><input type="date" id="op-data" value="${esc(data)}"></div>
      <div class="field">
        <label>Função / posto *</label>
        <select id="op-funcao" onchange="_opTrocouFuncao()">
          <option value="">— selecione —</option>
          ${funcoesOpts}
        </select>
        <div class="field-hint">Cadastre em <a href="#" onclick="closeModal('modal-op'); goto('cad-funcoes'); return false;">Funções</a>. As responsabilidades da função viram sugestões de operação.</div>
      </div>
      <div class="field">
        <label>Abrangência *</label>
        <select id="op-escopo" onchange="_opTrocouEscopo()">
          <option value="completa" ${escopo !== 'etapa' ? 'selected' : ''}>Processo completo do posto</option>
          <option value="etapa" ${escopo === 'etapa' ? 'selected' : ''}>Uma etapa só</option>
        </select>
        <div class="field-hint">O padrão engloba todas as etapas da função. Escolha <b>uma etapa só</b> quando o posto for planejado por partes.</div>
      </div>
      <div class="field hidden" id="op-wrap-etapa">
        <label>Etapa *</label>
        <select id="op-etapa" onchange="_opEscolheuEtapa()"></select>
        <div class="field-hint">Responsabilidades da função e etapas de produção cadastradas.</div>
      </div>
      <div class="field full">
        <label>Operação *</label>
        <input type="text" id="op-operacao" list="op-sugestoes" value="${esc(op ? (op.operacao || '') : '')}" placeholder="Ex.: Enfesto e corte do dia" autocomplete="off">
        <datalist id="op-sugestoes"></datalist>
        <div class="field-hint" id="op-sug-hint">Descreva a operação inteira do posto — as etapas internas ficam subentendidas.</div>
      </div>
      <div class="field">
        <label>Início *</label>
        <input type="time" id="op-inicio" value="${esc(op ? (op.inicio || '') : '07:00')}" oninput="_opAtualizarJanela()">
        <div class="field-hint">Hora planejada para o posto começar.</div>
      </div>
      <div class="field">
        <label>Duração total *</label>
        <div style="display:flex;gap:6px;align-items:center;">
          <input type="number" min="0" step="1" id="op-dur-h" value="${esc(durH)}" placeholder="0" oninput="_opAtualizarJanela()" style="width:70px;">
          <span style="font-size:12px;color:var(--ink-3);">h</span>
          <input type="number" min="0" max="59" step="5" id="op-dur-m" value="${esc(durM)}" placeholder="0" oninput="_opAtualizarJanela()" style="width:70px;">
          <span style="font-size:12px;color:var(--ink-3);">min</span>
        </div>
        <div class="field-hint">Tempo total até concluir, com todas as etapas do posto incluídas.</div>
      </div>
      <div class="field">
        <label>Responsável</label>
        <select id="op-responsavel"></select>
        <div class="field-hint">Pessoas da <b>Equipe</b> com esta função aparecem primeiro.</div>
      </div>
      <div class="field">
        <label>Referência (opcional)</label>
        <input type="text" id="op-referencia" list="op-refs" value="${esc(op ? (op.referencia || '') : '')}" placeholder="Ex.: lote inverno, OS 1042/1051" autocomplete="off">
        <datalist id="op-refs">${refs.map(r => `<option value="${esc(r)}"></option>`).join('')}</datalist>
        <div class="field-hint">Texto livre — lote, coleção, OSs do dia. Só para situar, não amarra o plano.</div>
      </div>
      <div class="field">
        <label>Classificação *</label>
        <select id="op-prioridade">${prioridadeOpts}</select>
        <div class="field-hint"><b>Eletiva</b> é a operação programada — o caso comum, sem selo na agenda. <b>Urgente</b> e <b>Emergente</b> ganham destaque na linha e são contadas no cabeçalho do dia.</div>
      </div>
      <div class="field"><label>Status</label><select id="op-status">${statusOpts}</select></div>
      <div class="field full"><label>Observação</label><input type="text" id="op-obs" value="${esc(op ? (op.obs || '') : '')}" placeholder="Ex.: depende da entrega do tecido"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;" id="op-info">Informe o início e a duração para ver o término.</div>`;

  _opTrocouFuncao(op ? (op.responsavelId || '') : '', op ? (op.etapa || '') : '');
  const selE = document.getElementById('op-etapa');
  if (selE) selE.dataset.anterior = op ? (op.etapa || '') : '';
  _opAtualizarJanela();
  openModal('modal-op');
}

// Função escolhida muda três coisas: as sugestões de operação, as etapas
// oferecidas quando o plano é de uma etapa só, e a lista de responsáveis.
// Mantém a pessoa e a etapa já selecionadas quando continuam válidas.
function _opTrocouFuncao(responsavelPre = null, etapaPre = null) {
  const selF = document.getElementById('op-funcao');
  const selR = document.getElementById('op-responsavel');
  const dl = document.getElementById('op-sugestoes');
  const hint = document.getElementById('op-sug-hint');
  if (!selF || !selR || !dl) return;
  const funcaoId = selF.value;
  const funcao = (STATE.funcoes || []).find(f => f.id === funcaoId);
  const manter = responsavelPre != null ? responsavelPre : selR.value;

  const sugs = _opSugestoesOperacao(funcaoId);
  dl.innerHTML = sugs.map(s => `<option value="${esc(s)}"></option>`).join('');
  if (hint) {
    const nAcoes = String(funcao && funcao.acoes || '').split('\n').filter(s => s.trim()).length;
    hint.innerHTML = 'Descreva a operação inteira do posto — as etapas internas ficam subentendidas. '
      + (!funcaoId
        ? 'Escolha a função para ver as sugestões cadastradas nela.'
        : (nAcoes
          ? `${nAcoes} responsabilidade(s) de <b>${esc(funcao.nome)}</b> disponíveis como sugestão.`
          : `<b>${esc(funcao.nome)}</b> não tem responsabilidades cadastradas; as sugestões são as etapas de produção.`));
  }

  // Etapas ofertadas quando o plano é de uma etapa só: as mesmas sugestões, mas
  // como lista fechada — aqui a escolha precisa ser uma etapa nomeada, não texto.
  const selE = document.getElementById('op-etapa');
  if (selE) {
    const etapaManter = etapaPre != null ? etapaPre : selE.value;
    selE.innerHTML = '<option value="">— selecione —</option>'
      + sugs.map(s => `<option value="${esc(s)}" ${s === etapaManter ? 'selected' : ''}>${esc(s)}</option>`).join('');
  }

  const { dentro, fora } = _opPessoasDaFuncao(funcao ? funcao.nome : '');
  const opt = p => `<option value="${esc(p.id)}" ${p.id === manter ? 'selected' : ''}>${esc(p.nome)}</option>`;
  selR.innerHTML = '<option value="">— a definir —</option>'
    + (dentro.length ? `<optgroup label="${esc(funcao ? funcao.nome : 'Da função')}">${dentro.map(opt).join('')}</optgroup>` : '')
    + (fora.length ? `<optgroup label="Outras pessoas">${fora.map(opt).join('')}</optgroup>` : '');

  _opTrocouEscopo();
}

// Mostra/esconde o seletor de etapa conforme a abrangência escolhida.
function _opTrocouEscopo() {
  const escopo = document.getElementById('op-escopo')?.value || 'completa';
  const wrap = document.getElementById('op-wrap-etapa');
  const hint = document.getElementById('op-sug-hint');
  if (wrap) wrap.classList.toggle('hidden', escopo !== 'etapa');
  if (hint && escopo === 'etapa') {
    hint.innerHTML = 'Planejando <b>uma etapa só</b>: escolha a etapa ao lado. O nome dela vem para cá e pode ser detalhado.';
  } else if (escopo !== 'etapa') {
    _opTrocouFuncaoHint();
  }
}
// Restaura o texto de ajuda do campo Operação para o modo "processo completo".
function _opTrocouFuncaoHint() {
  const hint = document.getElementById('op-sug-hint');
  const funcaoId = document.getElementById('op-funcao')?.value || '';
  const funcao = (STATE.funcoes || []).find(f => f.id === funcaoId);
  if (!hint) return;
  const nAcoes = String(funcao && funcao.acoes || '').split('\n').filter(s => s.trim()).length;
  hint.innerHTML = 'Descreva a operação inteira do posto — as etapas internas ficam subentendidas. '
    + (!funcaoId
      ? 'Escolha a função para ver as sugestões cadastradas nela.'
      : (nAcoes
        ? `${nAcoes} responsabilidade(s) de <b>${esc(funcao.nome)}</b> disponíveis como sugestão.`
        : `<b>${esc(funcao.nome)}</b> não tem responsabilidades cadastradas; as sugestões são as etapas de produção.`));
}

// Escolher a etapa preenche o nome da operação. Só sobrescreve campo vazio ou
// que ainda tem o nome da etapa anterior — texto digitado à mão fica de pé.
function _opEscolheuEtapa() {
  const selE = document.getElementById('op-etapa');
  const campo = document.getElementById('op-operacao');
  if (!selE || !campo) return;
  const anterior = selE.dataset.anterior || '';
  if (!campo.value.trim() || campo.value.trim() === anterior) campo.value = selE.value;
  selE.dataset.anterior = selE.value;
}

// Mostra ao vivo o término calculado — é o número que o planejador confere.
function _opAtualizarJanela() {
  const info = document.getElementById('op-info');
  if (!info) return;
  const ini = _opMin(document.getElementById('op-inicio')?.value);
  const dur = _opDuracaoDoForm();
  if (ini == null) { info.textContent = 'Informe a hora de início para ver o término.'; return; }
  if (!dur) { info.innerHTML = `Começa às <b>${esc(_opHHMM(ini))}</b>. Informe a duração total para calcular o término.`; return; }
  info.innerHTML = `Começa às <b>${esc(_opHHMM(ini))}</b>, leva <b>${esc(_opDurTexto(dur))}</b> e conclui às <b>${esc(_opHHMM(ini + dur))}</b>.`
    + (ini + dur > 1440 ? ' <span class="exp-badge baixo">atravessa a meia-noite</span>' : '');
}
function _opDuracaoDoForm() {
  const h = parseInt(document.getElementById('op-dur-h')?.value, 10) || 0;
  const m = parseInt(document.getElementById('op-dur-m')?.value, 10) || 0;
  return Math.max(0, h) * 60 + Math.max(0, m);
}

async function salvarModalOperacao() {
  if (!_opModalCtx) return;
  if (!exigirAdmin('planejar operações')) return;
  const v = id => document.getElementById(id)?.value || '';

  const data = v('op-data');
  if (!data) return toast('Informe a data da operação', 'err');
  const funcaoId = v('op-funcao');
  const funcao = (STATE.funcoes || []).find(f => f.id === funcaoId);
  if (!funcao) return toast('Escolha a função / posto', 'err');
  const escopo = v('op-escopo') === 'etapa' ? 'etapa' : 'completa';
  const etapa = escopo === 'etapa' ? v('op-etapa').trim() : '';
  if (escopo === 'etapa' && !etapa) return toast('Escolha a etapa que será executada', 'err');
  // Etapa escolhida e nome livre em branco: o nome da etapa já descreve a
  // operação — não faz sentido exigir que o usuário redigite o mesmo texto.
  const operacao = v('op-operacao').trim() || etapa;
  if (!operacao) return toast('Descreva a operação', 'err');
  const inicio = v('op-inicio');
  if (_opMin(inicio) == null) return toast('Informe a hora de início', 'err');
  const duracaoMin = _opDuracaoDoForm();
  if (!duracaoMin) return toast('Informe a duração total da operação', 'err');

  const responsavelId = v('op-responsavel');
  const pessoa = (STATE.equipe || []).find(p => p.id === responsavelId);
  const status = _OP_STATUS[v('op-status')] ? v('op-status') : 'pendente';
  const prioridade = _OP_PRIORIDADE[v('op-prioridade')] ? v('op-prioridade') : 'eletiva';

  const campos = {
    data,
    funcaoId, funcaoNome: funcao.nome,
    operacao, escopo, etapa,
    inicio, duracaoMin,
    responsavelId: pessoa ? pessoa.id : '',
    responsavelNome: pessoa ? pessoa.nome : '',
    referencia: v('op-referencia').trim(),
    status, prioridade,
    obs: v('op-obs').trim()
  };

  if (!Array.isArray(STATE.operacoes)) STATE.operacoes = [];
  if (_opModalCtx.editId) {
    const i = STATE.operacoes.findIndex(x => x.id === _opModalCtx.editId);
    if (i >= 0) STATE.operacoes[i] = { ...STATE.operacoes[i], ...campos };
  } else {
    STATE.operacoes.push({ id: uid(), ...campos });
  }
  await saveState('operacoes');
  closeModal('modal-op');
  toast(_opModalCtx.editId ? 'Operação atualizada' : 'Operação planejada', 'ok');
  _opModalCtx = null;
  // O dia salvo pode estar fora do período visível — leva a agenda até ele.
  const { ini, fim } = _expRange(opPlanoModo, opPlanoAncora);
  if (data < ini || data > fim) {
    opPlanoAncora = data;
    try { sessionStorage.setItem('gos:op:ancora', opPlanoAncora); } catch (e) {}
  }
  renderOperacoes();
}

async function alternarStatusOperacao(id) {
  if (!exigirAdmin('mudar o status das operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  op.status = _OP_CICLO[_opStatus(op)];
  await saveState('operacoes');
  renderOperacoes();
}

async function excluirOperacao(id) {
  if (!exigirAdmin('excluir operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  if (!confirm(`Excluir a operação "${op.operacao || 'sem descrição'}" de ${formatDate(op.data)}?`)) return;
  STATE.operacoes = (STATE.operacoes || []).filter(x => x.id !== id);
  await saveState('operacoes');
  toast('Operação excluída', 'ok');
  renderOperacoes();
}

// Repete no dia mostrado a jornada do último dia planejado antes dele. A jornada
// dos postos é estável — o dia seguinte quase sempre começa igual. Copia sempre
// como PENDENTE (é plano novo, não histórico) e pula o que já existe no destino,
// então clicar duas vezes não duplica a agenda.
async function copiarOperacoesDoDiaAnterior() {
  if (!exigirAdmin('planejar operações')) return;
  if (opPlanoModo !== 'dia') return toast('Mude para o modo Diário para repetir um dia', 'err');
  const destino = opPlanoAncora;
  const anteriores = (STATE.operacoes || []).filter(o => o.data && o.data < destino).map(o => o.data).sort();
  const origem = anteriores.length ? anteriores[anteriores.length - 1] : '';
  if (!origem) return toast('Não há dia anterior com operações planejadas', 'err');

  const chave = o => [o.funcaoId, _normNome(o.operacao), o.inicio || '', _opDuracao(o)].join('|');
  const jaTem = new Set((STATE.operacoes || []).filter(o => o.data === destino).map(chave));
  const novas = (STATE.operacoes || [])
    .filter(o => o.data === origem && !jaTem.has(chave(o)))
    .map(o => ({ ...o, id: uid(), data: destino, status: 'pendente' }));
  if (!novas.length) return toast(`Nada novo a copiar de ${formatDate(origem)}`, 'err');

  STATE.operacoes.push(...novas);
  await saveState('operacoes');
  toast(`${novas.length} operação(ões) copiada(s) de ${formatDate(origem)}`, 'ok');
  renderOperacoes();
}

// Operações gravadas no primeiro desenho do campo (uma linha por OS, com peças
// e sem horário) viram o formato de jornada: a OS passa a ser só referência e a
// operação fica sem horário até alguém definir início e duração.
function migrarOperacoesParaJornada() {
  const lista = STATE.operacoes;
  if (!Array.isArray(lista) || !lista.length) return false;
  let mudou = false;
  lista.forEach(op => {
    if (op.prioridade == null) { op.prioridade = 'eletiva'; mudou = true; }
    if (op.escopo == null) { op.escopo = 'completa'; op.etapa = ''; mudou = true; }
    if (op.duracaoMin == null) { op.duracaoMin = 0; mudou = true; }
    if (op.inicio == null) { op.inicio = op.hora || ''; mudou = true; }
    if (op.referencia == null) {
      const partes = [];
      if (op.osNumero) partes.push('OS ' + op.osNumero);
      if (op.pecas) partes.push(op.pecas + ' pç');
      op.referencia = partes.join(' · ');
      mudou = true;
    }
    if ('hora' in op) { delete op.hora; mudou = true; }
    if ('osId' in op) { delete op.osId; mudou = true; }
    if ('osNumero' in op) { delete op.osNumero; mudou = true; }
    if ('pecas' in op) { delete op.pecas; mudou = true; }
  });
  return mudou;
}

/* ---------------- folha impressa do plano ---------------- */

// Campos de assinatura do rodapé da OE, por FUNÇÃO. Quem assina a expedição não
// é "o responsável" genérico: são os três postos que a carga atravessa — quem
// separa e despacha, quem entrega as peças da costura e quem leva no veículo.
const _EXP_ASSINATURAS = ['Auxiliar de expedição', 'Auxiliar de costura', 'Motorista'];

// Segunda folha do programa. Ao contrário da folha de OS, esta pode ocupar
// várias A4 (um mês de janelas não cabe em uma) — o que não pode partir no
// meio é o bloco de cada expedição, garantido no CSS.
function renderPrintPlanoExpedicao() {
  const sheet = document.getElementById('print-sheet-exp');
  if (!sheet) return;
  const cfg = expCfg();
  const { ini, fim } = _expRange(expPlanoModo, expPlanoAncora);
  // A folha impressa só traz data com OE PRODUZIDA: precisa ter carga/OS alocada
  // em alguma das pernas. Dia agendado e vazio não vira papel. Cancelada também
  // sai, mesmo que tenha carga alocada antes do cancelamento — sem as pernas
  // (que o bloco não imprime) ela seria só um cabeçalho com a data, ocupando
  // espaço sem dizer nada.
  // Vale nos TRÊS modos (diário, semanal, mensal); antes era só no mensal, então
  // o semanal — que é o modo padrão — continuava imprimindo dia vazio.
  // A tela de planejamento segue mostrando vazios e cancelados: lá eles servem.
  const ocs = ocorrenciasExpedicao(ini, fim).filter(oc =>
    !oc.cancelada &&
    resumoPernaExpedicao(oc, 'ida').itens.length +
    resumoPernaExpedicao(oc, 'volta').itens.length > 0);
  const fmt = n => (Number(n) || 0).toLocaleString('pt-BR');

  let volIda = 0, volVolta = 0, pecasTot = 0, ativas = 0;
  const osTot = new Set();
  ocs.forEach(oc => {
    ativas++;
    ['ida', 'volta'].forEach(p => {
      const r = resumoPernaExpedicao(oc, p);
      if (p === 'ida') volIda += r.volumes; else volVolta += r.volumes;
      pecasTot += r.pecas;
      r.itens.forEach(i => { if (i.os) osTot.add(i.os.id); });
    });
  });

  // Cada OS da folha de OE é um QUADRO fechado: cabeçalho (nº, modelo, peças,
  // volumes), a conta dos volumes e a tabela de quantidades. Sem a moldura, a
  // tabela de uma OS encostava na linha da OS seguinte e as duas se liam como
  // um bloco só — quem confere a carga não achava onde uma acaba e a outra começa.
  //
  // Os VOLUMES são mostrados como o "Total por tamanho" da OS os define: um
  // pacote por tamanho de cada tonalidade, mais um de reposição. A tabela
  // detalha quantas peças vão em cada pacote — é o que a pessoa que ensaca lê.
  // Os números saem de totaisPorTamanhoTomOS, a mesma fonte da folha de OS.
  const TAM_LABEL = { p:'P', m:'M', g:'G', gg:'GG', g1:'G1', g2:'G2', g3:'G3' };

  // O checklist de COSTURA da própria OS, repetido no quadro dela na OE. Quem
  // recebe a carga do outro lado precisa saber o que da costura já foi feito e o
  // que falta naquela OS — sem isso a folha diz o que chegou, mas não em que pé
  // o trabalho está.
  //
  // O desenho é o MESMO da folha de OS: a etapa com a caixa dela e o nome em
  // negrito, e abaixo, indentadas, uma linha por etapa interna — as marcadas com
  // ✓, as vazias com o quadrinho em branco. Duas folhas do mesmo trabalho com
  // layouts diferentes obrigam a reaprender a leitura em cada uma, e a caixa
  // vazia também é onde se dá baixa à caneta na conferência.
  const costuraPrint = (o) => {
    if (!o) return '';
    const nomeEtapa = (o.etapas || []).find(n => /costura/i.test(n));
    if (!nomeEtapa) return '';
    const prog = o.progresso || {};
    const cad = (STATE.etapas || []).find(e => e.nome === nomeEtapa);
    const feitaEtapa = !!(prog.etapasCheck || {})[nomeEtapa];
    const marcadas = (prog.tarefasCheck || {})[nomeEtapa] || {};
    const tarefas = cad ? tarefasDaEtapa(cad).map(t => t.nome).filter(Boolean) : [];
    // Mesmo resgate da folha de OS: marca gravada por nome de tarefa que saiu do
    // cadastro continua na OS, e some da folha se a lista vier só do cadastro.
    Object.keys(marcadas).forEach(t => {
      if (marcadas[t] && t && !tarefas.includes(t)) tarefas.push(t);
    });
    const feitas = tarefas.filter(t => !!marcadas[t]).length;
    const cx = ok => `<span class="exp-print-box${ok ? ' ok' : ''}"></span>`;
    return `
      <div class="cost">
        <ul class="cl">
          <li class="et">${cx(feitaEtapa)}<strong>${esc(nomeEtapa)}</strong>${
            tarefas.length ? `<span class="cn">${feitas}/${tarefas.length}</span>` : ''}</li>
          ${tarefas.length
            ? tarefas.map(t => `<li class="tf">${cx(!!marcadas[t])}<span>${esc(t)}</span></li>`).join('')
            : '<li class="tf cv">sem etapas cadastradas nesta costura</li>'}
        </ul>
      </div>`;
  };

  const osPrint = (i) => {
    const o = i.os;
    const cab = `
      <div class="cab">
        <span class="exp-print-box"></span>
        <span class="n">${esc(i.osNumero)}</span>
        <span class="m">${esc(i.modelo)}</span>
        <span class="q">${fmt(i.pecas)} pç</span>
        <span class="v">${i.volumes > 0 ? fmt(i.volumes) + ' vol' : '— vol'}</span>
      </div>`;
    const costura = costuraPrint(o);
    const TT = o ? totaisPorTamanhoTomOS(o) : null;
    if (!TT || !TT.tamanhos.length) return `<div class="exp-print-os">${cab}${costura}</div>`;

    // Cor predominante = Cor 1 da 1ª variante (a mesma do banner da folha de OS),
    // caindo na cor da 1ª fase do enfesto. Sem o tecido no nome.
    const corPred = corNomeCurto(
      (o.variantes || []).map(v => v.cor1Nome).find(c => c && c !== '—')
      || (o.fases || [])[0]?.corNome
      || (o.tecidos || [])[0]?.corNome
      || ''
    );

    // A conta do volume, escrita por extenso: é a mesma regra do planejamento
    // (nº de tamanhos × tonalidades + 1 de reposição). Divergência contra o que
    // está alocado na carga fica à vista em vez de virar surpresa na doca.
    const nTam = _expTotalTamanhosGrade(o);
    const nTons = Math.max(1, TT.tons.length);
    const volCalc = nTam > 0 ? nTam * nTons + 1 : 0;
    const diverge = volCalc > 0 && i.volumes > 0 && volCalc !== i.volumes;
    // O volume extra não é só reposição: é o pacote que leva junto a ribana.
    // Escrito por extenso porque quem confere precisa saber o que procurar nele.
    const contaVol = volCalc > 0
      ? `<b>${fmt(volCalc)} volume${volCalc === 1 ? '' : 's'}</b> = ${fmt(nTam)} tamanho${nTam === 1 ? '' : 's'} × ${nTons} tonalidade${nTons === 1 ? '' : 's'} + 1 reposição e ribana`
      : `${fmt(i.volumes)} volume${i.volumes === 1 ? '' : 's'}`;

    // Linhas por tonalidade. Com DUAS OU MAIS tonalidades a repartição é o
    // detalhe que interessa; com uma tonalidade só (ou nenhuma marcada) ela
    // levaria a coluna inteira e a linha sairia idêntica à do total — repetir o
    // mesmo número duas vezes só faz o conferente procurar uma diferença que não
    // existe. Nesse caso a tonalidade é dita no cabeçalho e a tabela fica com uma
    // linha só. Com duas ou mais e nada repartido na OS, a divisão é declarada
    // indefinida em vez de inventada.
    const umTomSo = TT.tons.length <= 1;
    const indef = TT.semDigitacao && !umTomSo;
    const linhas = umTomSo ? [] : TT.linhas.map(L => ({
      rot: 'Tom ' + L.tom,
      cels: TT.tamanhos.map(k => L.cels[k]),
      total: L.total
    }));
    // Com uma tonalidade só, a linha única é a do total e é ela que carrega o
    // nome do tom — assim a folha continua dizendo por tamanho E por tom.
    const rotTotal = umTomSo ? (TT.tons.length ? 'Tom ' + TT.tons[0] : 'Tom único') : 'Total';

    const th = 'padding:0 2px;font-weight:700;border-bottom:.5pt solid #999;';
    const td = 'padding:0 2px;text-align:center;font-family:\'IBM Plex Mono\',monospace;';
    const cabec = TT.tamanhos.map(k => `<th style="${th}text-align:center;">${TAM_LABEL[k]}</th>`).join('');
    const linhaTotalTam = TT.tamanhos.map(k => `<td style="${td}font-weight:700;">${fmt(TT.colTotal(k))}</td>`).join('');
    const linhasTom = linhas.map(L => `
      <tr>
        <td style="${td}text-align:left;white-space:nowrap;">${esc(L.rot)}</td>
        ${L.cels.map(v => `<td style="${td}">${indef ? '—' : (v > 0 ? fmt(v) : '')}</td>`).join('')}
        <td style="${td}font-weight:700;background:#eef3ee;">${indef ? '—' : (L.total > 0 ? fmt(L.total) : '')}</td>
      </tr>`).join('');

    return `
      <div class="exp-print-os">
        ${cab}
        <div class="sub">
          <b>${esc(corPred) || '—'}</b> · ${contaVol}${diverge ? ` · <b>carga alocada com ${fmt(i.volumes)}</b>` : ''}
        </div>
        <table>
          <thead>
            <tr>
              <th style="${th}text-align:left;width:26pt;">Pacotes</th>
              ${cabec}
              <th style="${th}text-align:center;background:#eef3ee;">Total</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="${td}text-align:left;white-space:nowrap;font-weight:700;">${esc(rotTotal)}</td>
              ${linhaTotalTam}
              <td style="${td}font-weight:700;background:#eef3ee;">${TT.totalGeral > 0 ? fmt(TT.totalGeral) : ''}</td>
            </tr>
            ${linhasTom}
          </tbody>
        </table>
        <div class="pe">Cada célula é um pacote: peças daquele tamanho, naquela tonalidade.${indef ? ' A divisão entre as tonalidades ainda não foi repartida na OS.' : ''}</div>
        ${costura}
      </div>`;
  };

  const pernaPrint = (oc, perna) => {
    const r = resumoPernaExpedicao(oc, perna);
    const hora = perna === 'ida' ? oc.horaIda : oc.horaVolta;
    const linhas = r.itens.length
      ? r.itens.map(osPrint).join('')
      : '<div class="vazia">Sem OS alocada.</div>';
    return `
      <div class="exp-print-perna">
        <div class="ph">
          <div>
            <span class="t">${perna === 'ida' ? 'IDA' : 'VOLTA'}</span>
            <span class="r"> ${esc(_expRotaTexto(perna))}</span>
          </div>
          <span class="h">${esc(hora) || '—'}</span>
        </div>
        ${linhas}
        <div class="tot">
          <span>${fmt(r.volumes)} vol · ${fmt(r.pecas)} pç</span>
          <span>${esc(_expLimitesTexto(r.volMin, r.volMax))}${r.situacao === 'baixo' ? ' · ABAIXO' : (r.situacao === 'alto' ? ' · ACIMA' : '')}</span>
        </div>
      </div>`;
  };

  // Sem tratamento de cancelada aqui: o filtro acima já as tirou da folha.
  const blocos = ocs.map(oc => `
    <div class="exp-print-bloco">
      <div class="cab">
        <span class="d">${_EXP_DIAS_CURTO[_expData(oc.data).getDay()]} ${esc(formatDate(oc.data))}</span>
        <span class="j">
          ${esc(oc.janela.nome) || 'Janela sem nome'}
          ${oc.remarcada ? ` · remarcada de ${esc(formatDate(oc.dataOrig))}` : ''}
        </span>
      </div>
      <div class="exp-print-pernas">${pernaPrint(oc, 'ida')}${pernaPrint(oc, 'volta')}</div>
    </div>`).join('');

  const emissao = new Date();
  const emissaoTxt = formatDate(_expIso(emissao)) + ' ' + String(emissao.getHours()).padStart(2, '0') + ':' + String(emissao.getMinutes()).padStart(2, '0');

  sheet.innerHTML = `
    <div class="exp-print-head">
      <div>
        <div class="tit">ORDEM DE EXPEDIÇÃO OE</div>
        <div class="sub">Plano ${esc(_expNomeModo(expPlanoModo))} · ${esc(formatDate(ini))} a ${esc(formatDate(fim))} · expedição interna, ida e volta</div>
      </div>
      <div class="meta">
        <div><b>${esc(cfg.unidadeA)}</b> ⇄ <b>${esc(cfg.unidadeB)}</b></div>
        <div>Limite por perna: ${esc(_expLimitesTexto(_expNum(cfg.volMin, 0), _expNum(cfg.volMax, 0)))}</div>
        <div>Emitido em ${esc(emissaoTxt)}</div>
      </div>
    </div>
    <div class="exp-print-resumo">
      <div class="item"><div class="n">${fmt(ativas)}</div><div class="l">Expedições</div></div>
      <div class="item"><div class="n">${fmt(volIda)}</div><div class="l">Volumes ida</div></div>
      <div class="item"><div class="n">${fmt(volVolta)}</div><div class="l">Volumes volta</div></div>
      <div class="item"><div class="n">${fmt(volIda + volVolta)}</div><div class="l">Volumes total</div></div>
      <div class="item"><div class="n">${fmt(pecasTot)}</div><div class="l">Peças</div></div>
      <div class="item"><div class="n">${fmt(osTot.size)}</div><div class="l">OS alocadas</div></div>
    </div>
    ${ocs.length ? blocos : `<div style="padding:20px 0;text-align:center;font-size:9pt;font-style:italic;">Nenhuma Ordem de Expedição produzida ${esc(_EXP_VAZIO_PERIODO[expPlanoModo] || 'neste período')}.</div>`}
    <div class="exp-print-rodape">
      ${_EXP_ASSINATURAS.map(f => `<div class="ass"><div class="linha"></div><div class="lbl">${esc(f)}</div></div>`).join('')}
    </div>`;
}

// SKU(s) do produto acabado de uma OS = Linha de SKU do modelo + Sigla SKU de
// cada cor (variante). Override em o.skuOverride tem prioridade. Usado no
// cabeçalho da folha impressa e no snapshot para a Contabilidade/Estoque.
function skusDaOS(o) {
  // Valor base: override da OS > SKU do desenho técnico > SKU do modelo.
  const desenhoObj = (STATE.desenhos || []).find(d => d.id === o.desenhoId);
  const modeloObj = (STATE.modelos || []).find(m => m.id === o.modeloId);
  const base = ((o.skuOverride || (desenhoObj && desenhoObj.skuLinha) || (modeloObj && modeloObj.skuLinha)) || '').trim().toUpperCase();
  if (!base) return [];
  // Regra do traço: SKU COMPLETO (ex.: CM.LISA-PRE) tem "-" → usa direto.
  // LINHA (ex.: CM.LISA) não tem "-" → compõe com a Sigla da cor de cada variante.
  if (base.includes('-')) return [base];
  const cores = [...new Set((o.variantes || []).map(v => v.cor1Nome).filter(c => c && c !== '—'))];
  const out = [];
  cores.forEach(corNome => {
    const corObj = (STATE.cores || []).find(c => _normNome(c.nome) === _normNome(corNome));
    const sigla = ((corObj && corObj.siglaSku) || '').trim().toUpperCase();
    if (sigla) out.push(base + '-' + sigla);
  });
  return [...new Set(out)];
}

// Datalist com os SKUs COMPLETOS do catálogo do Estoque-Confeccao, para o
// dropdown dos campos de SKU nos cadastros de Desenho e Modelo.
function datalistSkusHtml() {
  const opts = (catalogoSkus || [])
    .map(s => `<option value="${esc(s.item)}">${esc(s.descricao || s.item)}</option>`)
    .join('');
  return `<datalist id="dl-skus">${opts}</datalist>`;
}

/* ========================================================= */
/*   SNAPSHOT PARA A CONTABILIDADE (quantidades p/ valorar)   */
/* ========================================================= */
// O programa de Contabilidade-Tributação declara os estoques lendo este
// snapshot do Supabase (chave 'contabSnapshot' no shared_data). Aqui só
// publicamos QUANTIDADES (a Contabilidade aplica os valores em R$: custo/kg
// das compras + R$/peça da mão de obra). Divisão: Gerador-OS = quantidades;
// Contabilidade = valores. Reescreve a chave a cada save relevante.
//   materiaPrima       = tecido disponível (entrada − reservado − saída), em kg.
//   produtosElaboracao = OSs cortadas e NÃO costuradas (work-in-progress):
//                        kg de tecido consumido + nº de peças, por tecido+cor.
//   ordens             = uma linha por OS com produção: data, camisetas produzidas
//                        (total da grade × camadas × multiplicador), se já costurada
//                        e o consumo de tecido por tecido+cor. A Contabilidade usa
//                        isto para ratear as despesas operacionais por peça/OS.
function construirContabSnapshot() {
  const r3 = n => Math.round((Number(n) || 0) * 1000) / 1000;
  const materiaPrima = (calcularSaldosEstoque().detalhe || [])
    .filter(d => Math.abs(d.disponivel) > 1e-9)
    .map(d => ({ tecido: d.tecidoNome || '', cor: d.corNome || '', kg: r3(d.disponivel) }));

  // WIP: agrega por tecido+cor as OSs cortadas que ainda não foram costuradas.
  const wip = new Map();
  (STATE.ordens || []).forEach(o => {
    if (osCosturaMarcada(o)) return;
    const peca = (componentesPorTecidoCorOS(o) || []);
    const kgs = (consumoAgregadoPorTecidoCor(o) || []);
    const pegar = (tNome, cNome) => {
      const k = _normNome(tNome) + '||' + _normNome(cNome);
      let cur = wip.get(k);
      if (!cur) { cur = { tecido: tNome || '', cor: cNome || '', kg: 0, pecas: 0 }; wip.set(k, cur); }
      return cur;
    };
    peca.forEach(it => { pegar(it.tecidoNome, it.corNome).pecas += (Number(it.qtd) || 0); });
    kgs.forEach(it => { pegar(it.tecidoNome, it.corNome).kg += (Number(it.kg) || 0); });
  });
  const produtosElaboracao = Array.from(wip.values())
    .filter(w => w.pecas > 0 || w.kg > 1e-9)
    .map(w => ({ tecido: w.tecido, cor: w.cor, kg: r3(w.kg), pecas: Math.round(w.pecas) }));

  // Por OS: produção (camisetas + por tamanho), material, modelo/cor e fase.
  // O Estoque-Confeccao usa `estoque` (etapa terminal "Estoque" marcada) como
  // gatilho para lançar a entrada de produtos acabados, casando pelo SKU da OS.
  const TAMS = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
  const ordens = (STATE.ordens || []).map(o => {
    const tamanhos = {};
    TAMS.forEach(t => { const q = Math.round(calcularColTotalAlvoImpressao(o, t) || 0); if (q > 0) tamanhos[t] = q; });
    // Cor: só casa direto quando a OS tem uma única cor (variante). Multicor
    // fica sem cor (vai para "a identificar" no Estoque-Confeccao).
    const coresV = [...new Set((o.variantes || []).map(v => v.cor1Nome).filter(c => c && c !== '—'))];
    const corPrincipal = coresV.length === 1 ? coresV[0] : '';
    // SKU para a entrada: único quando skusDaOS resolve exatamente 1 (cor única,
    // SKU completo definido, ou override). Multicor (vários) fica vazio → manual.
    const _skus = skusDaOS(o);
    const sku = _skus.length === 1 ? _skus[0] : '';
    return {
      os: o.os || '',
      data: (o.data || '').slice(0, 10),
      modelo: o.modeloNome || '',
      cor: corPrincipal,
      sku,
      multicor: coresV.length > 1,
      camisetas: Math.round(calcularTotalGeralAlvoImpressao(o) || 0),
      tamanhos,
      componentes: Math.round((componentesPorTecidoCorOS(o) || []).reduce((s, x) => s + (Number(x.qtd) || 0), 0)),
      costura: osCosturaMarcada(o),
      fios: osFiosMarcada(o),
      // Etapa terminal "Estoque" marcada = OS virou produto acabado. É o gatilho
      // (no Estoque-Confeccao) da entrada automática de produtos acabados por SKU.
      estoque: osEtapaMarcada(o, TERMINAL_ETAPA_RE),
      material: (consumoAgregadoPorTecidoCor(o) || [])
        .filter(x => (Number(x.kg) || 0) > 1e-9)
        .map(x => ({ tecido: x.tecidoNome || '', cor: x.corNome || '', kg: r3(x.kg) })),
    };
  }).filter(x => x.camisetas > 0 || x.componentes > 0 || x.material.length);

  return { geradoEm: new Date().toISOString(), materiaPrima, produtosElaboracao, ordens };
}

// Recalcula e grava o snapshot no blob (sem entrar no STATE/loadState — é só
// para consumo externo da Contabilidade). Best-effort: nunca quebra o save.
async function atualizarContabSnapshot() {
  try {
    await DB.set('contabSnapshot', JSON.stringify(construirContabSnapshot()));
  } catch (e) { console.warn('atualizarContabSnapshot', e); }
}

function renderModelos() {
  const tb = document.getElementById('tbl-modelos');
  if (!STATE.modelos.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhum modelo cadastrado.</td></tr>`; return; }
  const catLabel = { malha: 'Camiseta', moletom: 'Moletom', outro: 'Outro' };
  tb.innerHTML = STATE.modelos.map(m => `
    <tr><td><strong>${esc(m.nome)}</strong></td><td><span class="badge">${catLabel[m.categoria]||'—'}</span></td><td>${esc(m.linha)||'—'}</td>${acoesCell('modelo', m.id)}</tr>`).join('');
}
function renderColecoes() {
  const tb = document.getElementById('tbl-colecoes');
  if (!STATE.colecoes.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma coleção cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.colecoes.map(c => `
    <tr><td><strong>${esc(c.nome)}</strong></td><td>${esc(c.temporada)||'—'}</td>${acoesCell('colecao', c.id)}</tr>`).join('');
}
let pastasGradeExpandidas = new Set();

function toggleFolderGrade(path) {
  if (pastasGradeExpandidas.has(path)) pastasGradeExpandidas.delete(path);
  else pastasGradeExpandidas.add(path);
  renderGrades();
}

const TP_FIXOS = new Set(['camiseta', 'blusa_moletom', 'outro', '']);
const VR_FIXOS = new Set(['basica', 'bicolor', 'tricolor', '']);
const LABELS_TP_PADRAO = { camiseta: 'Camiseta', blusa_moletom: 'Blusa Moletom', outro: 'Outro', '': 'Sem categoria' };
const LABELS_VR_PADRAO = { basica: 'Básica', bicolor: 'Bicolor', tricolor: 'Tricolor', '': 'Sem variação' };

function _gfl() {
  STATE.gradeFolderLabels = STATE.gradeFolderLabels || { tp: {}, vr: {}, tpOrder: [], vrOrder: [] };
  STATE.gradeFolderLabels.tp = STATE.gradeFolderLabels.tp || {};
  STATE.gradeFolderLabels.vr = STATE.gradeFolderLabels.vr || {};
  STATE.gradeFolderLabels.tpOrder = STATE.gradeFolderLabels.tpOrder || [];
  STATE.gradeFolderLabels.vrOrder = STATE.gradeFolderLabels.vrOrder || [];
  return STATE.gradeFolderLabels;
}
function labelTp(tp) {
  const ov = _gfl().tp[tp];
  if (ov) return ov;
  return LABELS_TP_PADRAO[tp] !== undefined ? LABELS_TP_PADRAO[tp] : tp;
}
function labelVr(vr) {
  const ov = _gfl().vr[vr];
  if (ov) return ov;
  return LABELS_VR_PADRAO[vr] !== undefined ? LABELS_VR_PADRAO[vr] : vr;
}

async function renameGradeFolder(tpAtual) {
  const ehFixa = TP_FIXOS.has(tpAtual);
  const labelAtual = labelTp(tpAtual);
  const novo = (prompt('Novo nome da pasta:', labelAtual) || '').trim();
  if (!novo || novo === labelAtual) return;
  if (ehFixa) {
    // Renomeia só visualmente — a chave técnica continua sendo usada nos filtros da OS
    const gfl = _gfl();
    if (LABELS_TP_PADRAO[tpAtual] === novo) delete gfl.tp[tpAtual];
    else gfl.tp[tpAtual] = novo;
    await saveState('gradeFolderLabels');
  } else {
    if (TP_FIXOS.has(novo.toLowerCase())) { toast('Esse nome conflita com uma pasta fixa', 'err'); return; }
    let mexeu = 0;
    STATE.grades.forEach(g => { if ((g.tipoPeca || '') === tpAtual) { g.tipoPeca = novo; mexeu++; } });
    if (!mexeu) return;
    // Atualiza chaves de expansão e ordem
    const oldKey = 'tp:' + tpAtual;
    const newKey = 'tp:' + novo;
    if (pastasGradeExpandidas.has(oldKey)) { pastasGradeExpandidas.delete(oldKey); pastasGradeExpandidas.add(newKey); }
    const prefixOld = oldKey + '|var:';
    const prefixNew = newKey + '|var:';
    for (const k of [...pastasGradeExpandidas]) {
      if (k.startsWith(prefixOld)) {
        pastasGradeExpandidas.delete(k);
        pastasGradeExpandidas.add(prefixNew + k.slice(prefixOld.length));
      }
    }
    const gfl = _gfl();
    const idx = gfl.tpOrder.indexOf(tpAtual);
    if (idx >= 0) gfl.tpOrder[idx] = novo;
    await saveState('grades');
    await saveState('gradeFolderLabels');
  }
  renderGrades();
  toast('Pasta renomeada', 'ok');
}

async function renameGradeSubfolder(tp, vrAtual) {
  const ehFixa = VR_FIXOS.has(vrAtual);
  const labelAtual = labelVr(vrAtual);
  const novo = (prompt('Novo nome da subpasta:', labelAtual) || '').trim();
  if (!novo || novo === labelAtual) return;
  if (ehFixa) {
    const gfl = _gfl();
    if (LABELS_VR_PADRAO[vrAtual] === novo) delete gfl.vr[vrAtual];
    else gfl.vr[vrAtual] = novo;
    await saveState('gradeFolderLabels');
  } else {
    if (VR_FIXOS.has(novo.toLowerCase())) { toast('Esse nome conflita com uma subpasta fixa', 'err'); return; }
    let mexeu = 0;
    STATE.grades.forEach(g => {
      if ((g.tipoPeca || '') === tp && (g.variacao || '') === vrAtual) { g.variacao = novo; mexeu++; }
    });
    if (!mexeu) return;
    const oldKey = 'tp:' + tp + '|var:' + vrAtual;
    const newKey = 'tp:' + tp + '|var:' + novo;
    if (pastasGradeExpandidas.has(oldKey)) { pastasGradeExpandidas.delete(oldKey); pastasGradeExpandidas.add(newKey); }
    const gfl = _gfl();
    const idx = gfl.vrOrder.indexOf(vrAtual);
    if (idx >= 0) gfl.vrOrder[idx] = novo;
    await saveState('grades');
    await saveState('gradeFolderLabels');
  }
  renderGrades();
  toast('Subpasta renomeada', 'ok');
}

// Aplica ordem manual + fallback (fixos primeiro, depois custom alfabético)
function _ordenarPastas(chaves, ordemManual, fixosSet) {
  const presente = new Set(chaves);
  const naOrdem = ordemManual.filter(k => presente.has(k));
  const restantes = chaves.filter(k => !naOrdem.includes(k));
  const fixosOrdem = ['camiseta', 'blusa_moletom', 'outro', 'basica', 'bicolor', 'tricolor', ''];
  const fixosRest = restantes.filter(k => fixosSet.has(k)).sort((a,b) => fixosOrdem.indexOf(a) - fixosOrdem.indexOf(b));
  const customsRest = restantes.filter(k => !fixosSet.has(k)).sort((a,b) => labelTp(a).localeCompare(labelTp(b),'pt-BR'));
  return [...naOrdem, ...fixosRest, ...customsRest];
}

async function moveGradeFolder(tp, dir) {
  const gfl = _gfl();
  // Constrói a ordem corrente como aparece na tela e move o item
  const presentes = [...new Set(STATE.grades.map(g => g.tipoPeca || ''))];
  const ordem = _ordenarPastas(presentes, gfl.tpOrder, TP_FIXOS);
  const i = ordem.indexOf(tp);
  if (i < 0) return;
  const j = i + dir;
  if (j < 0 || j >= ordem.length) return;
  [ordem[i], ordem[j]] = [ordem[j], ordem[i]];
  gfl.tpOrder = ordem;
  await saveState('gradeFolderLabels');
  renderGrades();
}

async function moveGradeSubfolder(tp, vr, dir) {
  const gfl = _gfl();
  const presentes = [...new Set(STATE.grades.filter(g => (g.tipoPeca || '') === tp).map(g => g.variacao || ''))];
  const ordem = _ordenarPastas(presentes, gfl.vrOrder, VR_FIXOS);
  const i = ordem.indexOf(vr);
  if (i < 0) return;
  const j = i + dir;
  if (j < 0 || j >= ordem.length) return;
  [ordem[i], ordem[j]] = [ordem[j], ordem[i]];
  // vrOrder é compartilhado entre todas as pastas — é uma ordem geral de subpastas
  // (ex.: 'basica' antes de 'bicolor' globalmente). Atualiza só preservando outras chaves.
  const outros = gfl.vrOrder.filter(k => !ordem.includes(k));
  gfl.vrOrder = [...ordem, ...outros];
  await saveState('gradeFolderLabels');
  renderGrades();
}

function onSelectGradeFolder(sel, kind) {
  if (sel.value !== '__nova__') {
    sel.dataset.prev = sel.value;
    return;
  }
  const label = kind === 'pasta' ? 'Nome da nova pasta' : 'Nome da nova subpasta';
  const nome = (prompt(label + ':') || '').trim();
  if (!nome) { sel.value = sel.dataset.prev || ''; return; }
  const existente = Array.from(sel.options).find(o => o.value !== '__nova__' && o.value.toLowerCase() === nome.toLowerCase());
  if (existente) {
    sel.value = existente.value;
  } else {
    const opt = document.createElement('option');
    opt.value = nome;
    opt.textContent = nome;
    let grupo = Array.from(sel.querySelectorAll('optgroup')).find(g => g.label.startsWith(kind === 'pasta' ? 'Pastas' : 'Subpastas'));
    if (!grupo) {
      grupo = document.createElement('optgroup');
      grupo.label = kind === 'pasta' ? 'Pastas adicionais' : 'Subpastas adicionais';
      const novaOpt = Array.from(sel.options).find(o => o.value === '__nova__');
      sel.insertBefore(grupo, novaOpt);
    }
    grupo.appendChild(opt);
    sel.value = nome;
  }
  sel.dataset.prev = sel.value;
}

function renderGrades() {
  const tb = document.getElementById('tbl-grades');
  if (!STATE.grades.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhuma grade cadastrada.</td></tr>`; return; }

  const gfl = _gfl();

  // Agrupa por tipoPeca → variacao
  const grupos = {};
  for (const g of STATE.grades) {
    const tp = g.tipoPeca || '';
    const vr = g.variacao || '';
    grupos[tp] = grupos[tp] || {};
    grupos[tp][vr] = grupos[tp][vr] || [];
    grupos[tp][vr].push(g);
  }

  const ordemTipoPeca = _ordenarPastas(Object.keys(grupos), gfl.tpOrder, TP_FIXOS);

  const renderGradeRow = (g) => {
    const t = g.tamanhos || {};
    const dist = ['p','m','g','gg','g1','g2','g3']
      .filter(x => t[x] > 0).map(x => `${x.toUpperCase()}:${t[x]}`).join(' · ');
    const total = Object.values(t).reduce((a,b)=>a+(b||0),0);
    const nFases = Array.isArray(g.fases) ? g.fases.length : 0;
    const fasesBadge = nFases > 0 ? ` <span class="badge" style="background:#fff8e1">${nFases} fase${nFases>1?'s':''}</span>` : '';
    // Volume de expedição da grade, para UMA tonalidade: 1 pacote por tamanho
    // + 1 de reposição. Na OS o número é multiplicado pelo nº de tonalidades
    // marcadas (ver _expSugestaoVolumes) — aqui ainda não se sabe quantas são.
    const volBadge = total > 0 ? ` <span class="badge" title="Volume na expedição com 1 tonalidade: 1 pacote por tamanho + 1 de reposição. Com 2 tons dobra (${total * 2 + 1}), com 3 triplica (${total * 3 + 1}).">${total + 1} vol</span>` : '';
    return `<tr><td style="padding-left:48px;"><strong>${esc(g.nome)}</strong>${fasesBadge}${volBadge}</td>
      <td><code style="font-size:11px">${dist||'—'}</code></td>
      <td><span class="badge">${total}</span></td>${acoesCell('grade', g.id)}</tr>`;
  };

  const folderActions = (clickAttrs) => `<span class="folder-actions" onclick="event.stopPropagation()">${clickAttrs}</span>`;

  let html = '';
  for (let i = 0; i < ordemTipoPeca.length; i++) {
    const tp = ordemTipoPeca[i];
    if (!grupos[tp]) continue;
    const tpPath = 'tp:' + tp;
    const tpOpen = pastasGradeExpandidas.has(tpPath);
    const chevTop = tpOpen ? '▼' : '▶';
    const totalNoGrupo = Object.values(grupos[tp]).reduce((a, v) => a + v.length, 0);
    const tpJson = esc(JSON.stringify(tp));
    const upDis = i === 0 ? 'disabled' : '';
    const downDis = i === ordemTipoPeca.length - 1 ? 'disabled' : '';
    const acoesTp = folderActions(
      `<button type="button" class="folder-btn" title="Mover para cima" ${upDis} onclick="moveGradeFolder(${tpJson}, -1)">↑</button>`
      + `<button type="button" class="folder-btn" title="Mover para baixo" ${downDis} onclick="moveGradeFolder(${tpJson}, 1)">↓</button>`
      + `<button type="button" class="folder-btn" title="Renomear pasta" onclick="renameGradeFolder(${tpJson})">✎</button>`
    );
    html += `<tr class="grade-folder grade-folder-top" onclick="toggleFolderGrade('${esc(tpPath)}')"><td colspan="4">
      <span class="folder-chev">${chevTop}</span> 📁 ${esc(labelTp(tp))}
      <span class="folder-count">(${totalNoGrupo})</span>
      ${acoesTp}
    </td></tr>`;
    if (!tpOpen) continue;

    const ordemVariacao = _ordenarPastas(Object.keys(grupos[tp]), gfl.vrOrder, VR_FIXOS);
    for (let j = 0; j < ordemVariacao.length; j++) {
      const vr = ordemVariacao[j];
      const gs = grupos[tp][vr];
      if (!gs || !gs.length) continue;
      const vrPath = tpPath + '|var:' + vr;
      const vrOpen = pastasGradeExpandidas.has(vrPath);
      const chevSub = vrOpen ? '▼' : '▶';
      const vrJson = esc(JSON.stringify(vr));
      const upDisV = j === 0 ? 'disabled' : '';
      const downDisV = j === ordemVariacao.length - 1 ? 'disabled' : '';
      const acoesVr = folderActions(
        `<button type="button" class="folder-btn" title="Mover para cima" ${upDisV} onclick="moveGradeSubfolder(${tpJson}, ${vrJson}, -1)">↑</button>`
        + `<button type="button" class="folder-btn" title="Mover para baixo" ${downDisV} onclick="moveGradeSubfolder(${tpJson}, ${vrJson}, 1)">↓</button>`
        + `<button type="button" class="folder-btn" title="Renomear subpasta" onclick="renameGradeSubfolder(${tpJson}, ${vrJson})">✎</button>`
      );
      html += `<tr class="grade-folder grade-folder-sub" onclick="event.stopPropagation(); toggleFolderGrade('${esc(vrPath)}')"><td colspan="4">
        <span class="folder-chev">${chevSub}</span> ↳ ${esc(labelVr(vr))}
        <span class="folder-count">(${gs.length})</span>
        ${acoesVr}
      </td></tr>`;
      if (!vrOpen) continue;
      html += gs.map(renderGradeRow).join('');
    }
  }
  tb.innerHTML = html;
}
function renderDesenhos() {
  const tb = document.getElementById('tbl-desenhos');
  if (!STATE.desenhos.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhum desenho cadastrado.</td></tr>`; return; }
  const ordenados = [...STATE.desenhos].sort((a, b) =>
    (a.codigo || '').localeCompare(b.codigo || '', 'pt-BR', { numeric: true, sensitivity: 'base' })
  );
  tb.innerHTML = ordenados.map(d => `
    <tr>
      <td><div style="width:60px;height:45px;background:#f5f2ea;display:flex;align-items:center;justify-content:center;border:1px solid var(--line);overflow:hidden">
        ${d.img ? `<img src="${d.img}" style="max-width:100%;max-height:100%;object-fit:contain;">` : '—'}</div></td>
      <td><strong>${esc(d.codigo)}</strong></td><td>${esc(d.desc)||'—'}</td>${acoesCell('desenho', d.id)}</tr>`).join('');
}
function renderMarcas() {
  const tb = document.getElementById('tbl-marcas');
  if (!STATE.marcas.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma marca cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.marcas.map(m => `
    <tr><td><strong>${esc(m.nome)}</strong></td><td>${esc(m.desc)||'—'}</td>${acoesCell('marca', m.id)}</tr>`).join('');
}
function renderLinhas() {
  const tb = document.getElementById('tbl-linhas');
  if (!STATE.linhas.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma linha cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.linhas.map(l => `
    <tr><td><strong>${esc(l.nome)}</strong></td><td>${esc(l.desc)||'—'}</td>${acoesCell('linha', l.id)}</tr>`).join('');
}
function renderBases() {
  const tb = document.getElementById('tbl-bases');
  if (!STATE.bases.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma base cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.bases.map(b => `
    <tr><td><strong>${esc(b.nome)}</strong></td><td>${esc(b.desc)||'—'}</td>${acoesCell('base', b.id)}</tr>`).join('');
}
function renderBlocos() {
  const tb = document.getElementById('tbl-blocos');
  if (!STATE.blocos.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhum bloco cadastrado.</td></tr>`; return; }
  tb.innerHTML = STATE.blocos.map(b => `
    <tr><td><strong>${esc(b.nome)}</strong></td><td>${esc(b.desc)||'—'}</td>${acoesCell('bloco', b.id)}</tr>`).join('');
}
function renderEquipe() {
  const tb = document.getElementById('tbl-equipe');
  if (!STATE.equipe.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma pessoa cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.equipe.map(p => `
    <tr><td><strong>${esc(p.nome)}</strong></td><td><span class="badge">${esc(p.funcao)||'—'}</span></td>${acoesCell('equipe', p.id)}</tr>`).join('');
}
function etapasOrdenadas() {
  return [...STATE.etapas].sort((a,b) => (a.ordem||0) - (b.ordem||0));
}

function nomesFuncoesPorIds(ids) {
  if (!ids || !ids.length) return [];
  return ids
    .map(id => STATE.funcoes.find(f => f.id === id))
    .filter(Boolean)
    .map(f => f.nome);
}

function nomesTarefasPorIds(ids) {
  if (!ids || !ids.length) return [];
  return ids
    .map(id => STATE.tarefas.find(t => t.id === id))
    .filter(Boolean)
    .map(t => t.nome);
}

function renderComponentesCad() {
  const tb = document.getElementById('tbl-componentes');
  if (!STATE.componentes.length) { tb.innerHTML = `<tr><td colspan="6" class="empty">Nenhum componente cadastrado.</td></tr>`; return; }
  const labelTipoLegacy = { camiseta: 'Camiseta', blusa_moletom: 'Blusa Moletom', outro: 'Outro' };
  const labelVar = { basica: 'Básica', bicolor: 'Bicolor', tricolor: 'Tricolor' };
  const modeloById = new Map(STATE.modelos.map(m => [m.id, m]));
  const corById = new Map(STATE.cores.map(x => [x.id, x]));
  // Detecta nomes duplicados (case-insensitive, trim)
  const contagemNomes = new Map();
  STATE.componentes.forEach(c => {
    const k = (c.nome || '').trim().toLowerCase();
    if (!k) return;
    contagemNomes.set(k, (contagemNomes.get(k) || 0) + 1);
  });
  const duplicado = nome => contagemNomes.get((nome || '').trim().toLowerCase()) > 1;
  const corSwatch = (id) => {
    const c = corById.get(id);
    if (!c) return '';
    return `<span class="badge" style="display:inline-flex;align-items:center;gap:4px;margin-right:4px;">
      <span style="display:inline-block;width:10px;height:10px;border:1px solid var(--line);background:${esc(c.hex||'#fff')};"></span>
      ${esc(c.nome)}
    </span>`;
  };
  const tipoLabel = (v) => {
    if (!v) return '—';
    const m = modeloById.get(v);
    if (m) return `<span class="badge">${esc(m.nome)}</span>`;
    if (labelTipoLegacy[v]) return `<span class="badge">${esc(labelTipoLegacy[v])}</span>`;
    return `<span class="badge">${esc(v)}</span>`;
  };
  tb.innerHTML = STATE.componentes.map(c => {
    const cores = [c.cor1Id, c.cor2Id, c.cor3Id].filter(Boolean).map(corSwatch).join('') || '—';
    const dupBadge = duplicado(c.nome)
      ? ' <span class="badge" style="background:#fff3cd;color:#856404;border:1px solid #ffc107;" title="Existem múltiplos componentes com este nome — o auto-preenchimento de cor pode pegar o errado">⚠ Nome duplicado</span>'
      : '';
    return `
    <tr>
      <td><strong>${esc(c.nome)}</strong>${dupBadge}</td>
      <td>${tipoLabel(c.tipoPeca)}</td>
      <td>${c.variacao ? `<span class="badge">${esc(labelVar[c.variacao]||c.variacao)}</span>` : '—'}</td>
      <td>${cores}</td>
      <td>${esc(c.desc)||'—'}</td>
      ${acoesCell('componente', c.id)}
    </tr>`;
  }).join('');
}

// Tarefas de uma etapa: prioriza item.tarefas (estrutura nova, embutida na etapa);
// fallback p/ tarefasIds + STATE.tarefas (modelo antigo).
function tarefasDaEtapa(etapa) {
  if (Array.isArray(etapa?.tarefas) && etapa.tarefas.length) return etapa.tarefas;
  if (Array.isArray(etapa?.tarefasIds) && etapa.tarefasIds.length) {
    return etapa.tarefasIds
      .map(tid => (STATE.tarefas || []).find(t => t.id === tid))
      .filter(Boolean);
  }
  return [];
}

function renderEtapasCad() {
  const cont = document.getElementById('etapas-pastas');
  if (!cont) return;
  if (!STATE.etapas.length) {
    cont.innerHTML = `<div class="card" style="padding:20px;text-align:center;color:var(--ink-3);">Nenhuma etapa cadastrada. Use <strong>+ Nova etapa</strong> para começar.</div>`;
    return;
  }
  const acoesEtapa = (id) => `
    <span class="row-actions" style="display:inline-flex;gap:4px;">
      <button class="edit" onclick="openCadastroModal('etapa','${esc(id)}')">editar</button>
      <button class="edit" onclick="duplicarCadastro('etapa','${esc(id)}')">duplicar</button>
      <button class="del" onclick="excluirCadastro('etapa','${esc(id)}')">excluir</button>
    </span>`;

  const html = etapasOrdenadas().map(e => {
    const tarefas = tarefasDaEtapa(e);
    const sub = tarefas.length
      ? tarefas.map(t => `
          <li style="display:flex;align-items:center;gap:10px;padding:6px 8px;border-bottom:1px dotted var(--line);">
            <span style="font-size:14px;">📄</span>
            <strong style="flex:0 0 auto;">${esc(t.nome)}</strong>
            <span style="color:var(--ink-3);font-size:12px;flex:1;">${esc(t.desc) || ''}</span>
          </li>`).join('')
      : `<li style="padding:8px;color:var(--ink-3);font-style:italic;font-size:12px;">Nenhuma tarefa nesta etapa ainda — clique em <strong>editar</strong> para adicionar.</li>`;
    return `
      <div class="card etapa-pasta" style="margin-bottom:10px;padding:0;overflow:hidden;">
        <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;background:var(--line-2);border-bottom:1px solid var(--line);">
          <span style="font-size:18px;">📁</span>
          <span class="badge" title="Ordem">${e.ordem || 0}</span>
          <strong style="flex:1;font-size:14px;">${esc(e.nome)}</strong>
          <span style="color:var(--ink-3);font-size:11px;">${tarefas.length} tarefa${tarefas.length===1?'':'s'}</span>
          ${acoesEtapa(e.id)}
        </div>
        <ul style="list-style:none;margin:0;padding:6px 12px 10px 28px;">
          ${sub}
        </ul>
      </div>`;
  }).join('');
  cont.innerHTML = html;
}

function renderFuncoes() {
  const tb = document.getElementById('tbl-funcoes');
  if (!STATE.funcoes.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhuma função cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.funcoes.map(f => {
    const acoes = (f.acoes||'').trim();
    const acoesHtml = acoes ? acoes.split(/\r?\n/).filter(x=>x.trim()).map(a => `<span class="badge" style="margin-right:4px">${esc(a)}</span>`).join('') : '—';
    return `<tr><td><strong>${esc(f.nome)}</strong></td><td>${esc(f.desc)||'—'}</td><td>${acoesHtml}</td>${acoesCell('funcao', f.id)}</tr>`;
  }).join('');
}


function renderHome() {
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  set('stat-os', STATE.ordens.length);
  set('stat-tecidos', STATE.tecidos.length);
  set('stat-cores', STATE.cores.length);
  set('stat-materiais', STATE.materiais.length);
  set('stat-modelos', STATE.modelos.length);
  set('stat-colecoes', STATE.colecoes.length);
  set('stat-grades', STATE.grades.length);
  set('stat-desenhos', STATE.desenhos.length);
  set('stat-marcas', STATE.marcas.length);
  set('stat-linhas', STATE.linhas.length);
  set('stat-bases', STATE.bases.length);
  set('stat-blocos', STATE.blocos.length);
  set('stat-equipe', STATE.equipe.length);
}

/* ========================================================= */
/*                   FORMULÁRIO DA OS                        */
/* ========================================================= */
let osEditId = null;

function funcaoPorNome(nome) {
  if (!nome) return null;
  return STATE.funcoes.find(f => f.nome === nome) || null;
}

function renderResponsabilidadesBadges(f) {
  if (!f || !f.acoes) return '';
  const linhas = (f.acoes || '').split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  if (!linhas.length) return '';
  return linhas.map(a => `<span class="badge" style="margin:2px 3px 0 0;display:inline-block;">${esc(a)}</span>`).join('');
}

function mostrarResponsabilidadesFuncao() {
  const sel = document.getElementById('m-funcao');
  const resp = document.getElementById('m-funcao-resp');
  if (!sel || !resp) return;
  const f = funcaoPorNome(sel.value);
  const html = renderResponsabilidadesBadges(f);
  resp.innerHTML = html ? `Responsabilidades: ${html}` : '';
}

function atualizarResponsabilidadesOS() {
  const mapa = { 'f-designer': 'resp-designer', 'f-ftec': 'resp-ftec', 'f-coordenado': 'resp-coordenador' };
  const resumoItens = [];
  Object.entries(mapa).forEach(([selId, respId]) => {
    const sel = document.getElementById(selId);
    const respEl = document.getElementById(respId);
    if (!sel || !respEl) return;
    const pessoaNome = sel.value;
    const pessoa = STATE.equipe.find(p => p.nome === pessoaNome);
    if (!pessoa) { respEl.innerHTML = ''; return; }
    const f = funcaoPorNome(pessoa.funcao);
    const badges = renderResponsabilidadesBadges(f);
    respEl.innerHTML = badges ? badges : '';
    if (badges) {
      const rotulo = { 'f-designer': 'Designer', 'f-ftec': 'Ficha técnica', 'f-coordenado': 'Coordenador' }[selId];
      resumoItens.push(`<div style="margin-bottom:6px;"><strong>${rotulo}: ${esc(pessoaNome)}</strong> <span style="color:var(--ink-3);font-weight:normal;">(${esc(pessoa.funcao||'—')})</span><br>${badges}</div>`);
    }
  });
  const resumo = document.getElementById('resp-resumo');
  if (resumo) {
    resumo.innerHTML = resumoItens.length
      ? resumoItens.join('')
      : '<em style="color:var(--ink-3);">Selecione Designer, Ficha técnica ou Coordenador acima para ver responsabilidades da equipe.</em>';
  }
}

function atualizarDatalistCodigos() {
  const dl = document.getElementById('codigos-datalist');
  if (!dl) return;
  dl.innerHTML = STATE.desenhos.map(d =>
    `<option value="${esc(d.codigo)}">${esc(d.desc||'')}</option>`
  ).join('');
}

function sincCodigoDesenho(origem) {
  const codigoEl = document.getElementById('f-codigo');
  const desenhoEl = document.getElementById('f-desenho');
  if (!codigoEl || !desenhoEl) return;
  if (origem === 'desenho') {
    const id = desenhoEl.value;
    if (id) {
      const d = STATE.desenhos.find(x => x.id === id);
      if (d) codigoEl.value = d.codigo;
    } else {
      codigoEl.value = '';
    }
    preencherDropdownGradesOS();
    aplicarVinculosDesenho();
  } else {
    const typed = codigoEl.value.trim();
    if (!typed) {
      desenhoEl.value = '';
      previewDesenhoSelecionado();
      preencherDropdownGradesOS();
      return;
    }
    const d = STATE.desenhos.find(x => x.codigo.toLowerCase() === typed.toLowerCase());
    if (d && desenhoEl.value !== d.id) {
      desenhoEl.value = d.id;
      previewDesenhoSelecionado();
      preencherDropdownGradesOS();
      aplicarVinculosDesenho();
    }
  }
}

function aplicarVinculosDesenho() {
  const desenhoId = document.getElementById('f-desenho')?.value;
  if (!desenhoId) return;
  const d = STATE.desenhos.find(x => x.id === desenhoId);
  if (!d) return;
  aplicandoVinculosDesenho = true;
  try {
    // 1º aplica vínculos do modelo (base/designer/ftec/coord/marca/grade padrões)
    if (d.modeloId) {
      document.getElementById('f-modelo').value = d.modeloId;
      aplicarVinculosModelo();
    }
    // 2º sobrescreve com vínculos específicos do desenho (têm prioridade)
    const mapa = {
      modeloId: 'f-modelo', baseId: 'f-base', colecaoId: 'f-colecao',
      marcaId: 'f-griffe', linhaId: 'f-linha', blocoId: 'f-bloco',
      designerId: 'f-designer', coordId: 'f-coordenado'
    };
    let aplicou = false;
    Object.entries(mapa).forEach(([campo, selId]) => {
      const el = document.getElementById(selId);
      if (!el) return;
      if (d[campo]) { el.value = d[campo]; aplicou = true; }
    });
    aplicarFiltroTecidosPorModelo();
    atualizarResponsabilidadesOS();
    atualizarCalculosEnfesto();
    // Aplica componentes padrão do desenho — mesma fonte que o botão "Repor
    // componentes do desenho" usa (_componentesDoDesenho), pra não divergirem.
    const compsDesenho = _componentesDoDesenho(d);
    if (compsDesenho.length) {
      const cont = document.getElementById('componentes-rows');
      if (cont) {
        cont.innerHTML = '';
        compsDesenho.forEach(c => addComponenteRow(c));
        aplicou = true;
      }
    }
    // Aplica etapas padrão do desenho (marca + ordena)
    if (Array.isArray(d.etapasNomes) && d.etapasNomes.length) {
      document.querySelectorAll('#etapas-container .etapa-check').forEach(lbl => {
        const input = lbl.querySelector('input');
        const on = d.etapasNomes.includes(input.value);
        input.checked = on;
        lbl.classList.toggle('checked', on);
      });
      aplicarOrdemEtapas(d.etapasNomes);
      aplicou = true;
    }

    // Aplica aviamentos padrão do desenho — estrutura nova com qtd/peça + aplicação
    const avsDesenho = Array.isArray(d.aviamentos) && d.aviamentos.length
      ? d.aviamentos
      : (d.aviamentosIds || []).map(id => ({ materialId: id, qtdPorPeca: 1, aplicacao: '' }));
    if (avsDesenho.length) {
      const avCont = document.getElementById('aviamentos-rows');
      if (avCont) {
        avCont.innerHTML = '';
        avsDesenho.forEach(av => {
          if (STATE.materiais.find(x => x.id === av.materialId)) {
            addAviamentoRow({
              material: av.materialId,
              qtd: av.qtdPorPeca,
              app: av.aplicacao || ''
            });
          }
        });
        aplicou = true;
      }
    }
    // Aplica cor principal + secundária + terciária na Variante 1 (cria a row se não existir)
    if (d.corPrincipalId || d.corSecundariaId || d.corTerciariaId) {
      const varCont = document.getElementById('variantes-rows');
      if (varCont) {
        if (!varCont.querySelector('.variante-row')) addVarianteRow();
        const primeira = varCont.querySelector('.variante-row');
        if (primeira) {
          const c1 = primeira.querySelector('.var-c1');
          const c2 = primeira.querySelector('.var-c2');
          const c3 = primeira.querySelector('.var-c3');
          // Ordena pela sequência canônica do desenho (desc) pra a variante nascer
          // já na ordem certa — igual ao banner e às fases de enfesto.
          const coresOrd = ordenarCoresIdsPorDesc(
            [d.corPrincipalId, d.corSecundariaId, d.corTerciariaId].filter(Boolean), d);
          if (c1 && coresOrd[0]) c1.value = coresOrd[0];
          if (c2 && coresOrd[1]) c2.value = coresOrd[1];
          if (c3 && coresOrd[2]) c3.value = coresOrd[2];
          aplicou = true;
        }
      }
    }
    // Aplica tecido/cores do desenho nas linhas de Tecidos (1 linha por cor).
    // Na sequência canônica do desenho (desc) — mesma ordem do banner/enfesto.
    const coresDoDesenho = ordenarCoresIdsPorDesc(
      [d.corPrincipalId, d.corSecundariaId, d.corTerciariaId].filter(Boolean), d);
    if (d.tecidoPadraoId || coresDoDesenho.length) {
      const tecCont = document.getElementById('tecidos-rows');
      // Grade preset selecionado tem prioridade — ela preenche tecidos pelas fases.
      const gradePresetAtivo = !!document.getElementById('f-grade-preset')?.value;
      if (tecCont && !gradePresetAtivo) {
        tecCont.innerHTML = '';
        const n = Math.max(coresDoDesenho.length, 1);
        for (let i = 0; i < n; i++) {
          addTecidoRow({ tecidoId: d.tecidoPadraoId || '', corId: coresDoDesenho[i] || '' });
        }
        aplicou = true;
      }
    }

    // Ajusta quantidade de blocos de enfesto pela quantidade de cores do desenho
    // (se ainda não foram preenchidos pela grade)
    const enfestoCont = document.getElementById('f-enfestos-blocos');
    if (enfestoCont && coresDoDesenho.length) {
      const blocosAtuais = enfestoCont.querySelectorAll('.enfesto-bloco');
      const algumPreenchido = Array.from(blocosAtuais).some(b =>
        b.querySelector('.enf-comp').value || b.querySelector('.enf-larg').value);
      if (!algumPreenchido) {
        renderEnfestoBlocos(coresDoDesenho.length);
        aplicou = true;
      }
    }
    // Se já há grade preset ativa, re-aplica pra recalcular as cores das fases
    // com base no NOVO desenho (corPorFase usa desenhoAtual). Sem isso, cores
    // antigas do desenho anterior ficam congeladas em tecidos/enfesto.
    const gradePresetAtivo = !!document.getElementById('f-grade-preset')?.value;
    if (gradePresetAtivo) aplicarGradePreset();
    if (aplicou) toast('Campos vinculados preenchidos automaticamente', 'ok');
  } finally {
    aplicandoVinculosDesenho = false;
  }
}

function initOSForm() {
  // presence: marca o canal da OS sendo editada
  iniciarPresenceOS(osEditId || 'nova');

  // popula dropdowns
  fillSelect('f-colecao', STATE.colecoes, 'nome', '— selecione —');
  fillSelect('f-modelo', STATE.modelos, 'nome', '— selecione —');
  fillSelect('f-desenho', STATE.desenhos, 'codigo', '— selecione —', d => `${d.codigo}${d.desc ? ' · '+d.desc : ''}`);
  preencherDropdownGradesOS();
  atualizarDatalistCodigos();

  // novos selects do cabeçalho
  fillSelect('f-griffe', STATE.marcas, 'nome', '— selecione —');
  fillSelect('f-linha', STATE.linhas, 'nome', '— selecione —');
  fillSelect('f-base', STATE.bases, 'nome', '— selecione —');
  fillSelect('f-bloco', STATE.blocos, 'nome', '— selecione —');
  fillSelect('f-designer', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));
  fillSelect('f-ftec', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));
  // coordenador puxa da equipe (pessoa que coordena)
  fillSelect('f-coordenado', STATE.equipe, 'nome', '— selecione —', p => p.nome + (p.funcao ? ' ('+p.funcao+')' : ''));

  // se não estiver editando e os campos principais estão vazios, limpar e inicializar linhas
  if (!osEditId) {
    document.getElementById('os-form').reset();
    document.getElementById('f-id').value = '';
    document.getElementById('f-data').value = new Date().toISOString().slice(0,10);
    document.getElementById('os-form-title').textContent = 'Nova Ordem de Serviço';
    // número OS automático sequencial
    document.getElementById('f-os').value = proximoNumeroOS();
    // Peças-alvo já nasce em 160 (padrão da casa), como o número da OS e a data.
    // Só no formulário NOVO: editar uma OS existente carrega o valor salvo dela,
    // mais abaixo. Continua editável — é só o ponto de partida.
    document.getElementById('f-enf-target').value = PECAS_ALVO_PADRAO;
    // linhas iniciais
    document.getElementById('tecidos-rows').innerHTML = '';
    document.getElementById('variantes-rows').innerHTML = '';
    document.getElementById('componentes-rows').innerHTML = '';
    document.getElementById('aviamentos-rows').innerHTML = '';
    addTecidoRow(); addTecidoRow();
    addVarianteRow();
    addComponenteRow(); addComponenteRow();
    renderEnfestoBlocos(1);
    document.getElementById('f-desenho-preview').innerHTML = '<span>Nenhum desenho selecionado</span>';
  }

  renderEtapas();
  atualizarCalculosEnfesto();
  atualizarResponsabilidadesOS();
}

function fillSelect(id, items, labelField, placeholder, custom = null) {
  const el = document.getElementById(id);
  if (!el) return;
  const curVal = el.value;
  el.innerHTML = `<option value="">${placeholder}</option>` +
    items.map(it => `<option value="${esc(it.id)}">${esc(custom ? custom(it) : it[labelField])}</option>`).join('');
  if (curVal) el.value = curVal;
}

// Define o valor de um select — primeiro tenta pelo ID direto; se não achar, tenta casar por nome (fallback para OS antigas)
function setSelectByIdOrName(selectId, itemId, nameFallback, list) {
  const el = document.getElementById(selectId);
  if (!el) return;
  if (itemId) {
    const hasOpt = Array.from(el.options).some(o => o.value === itemId);
    if (hasOpt) { el.value = itemId; return; }
  }
  if (nameFallback && list?.length) {
    const match = list.find(x => (x.nome || '').toLowerCase() === nameFallback.toLowerCase());
    if (match) el.value = match.id;
  }
}

function renderEtapas() {
  const cont = document.getElementById('etapas-container');
  if (!cont) return;
  const checked = Array.from(cont.querySelectorAll('input:checked')).map(c => c.value);
  const fonte = STATE.etapas.length
    ? etapasOrdenadas().map(e => ({ nome: e.nome, tarefas: tarefasDaEtapa(e).map(t => t.nome) }))
    : STATE.etapasPadrao.map(nome => ({ nome, tarefas: [] }));
  cont.innerHTML = fonte.map(e => {
    const tarefasBadges = e.tarefas.length
      ? `<div style="font-size:10px;color:var(--ink-3);margin-top:3px;">${e.tarefas.map(t => `<span class="badge" style="margin-right:3px;font-size:10px;padding:1px 5px;">${esc(t)}</span>`).join('')}</div>`
      : '';
    return `<label class="etapa-check ${checked.includes(e.nome)?'checked':''}">
      <span class="etapa-reorder">
        <button type="button" class="etapa-move" onclick="event.preventDefault(); event.stopPropagation(); moverEtapaForm(this, -1)" title="Mover para cima">▲</button>
        <button type="button" class="etapa-move" onclick="event.preventDefault(); event.stopPropagation(); moverEtapaForm(this, 1)" title="Mover para baixo">▼</button>
      </span>
      <input type="checkbox" value="${esc(e.nome)}" ${checked.includes(e.nome)?'checked':''} onchange="this.parentElement.classList.toggle('checked', this.checked)">
      <span>${esc(e.nome)}${tarefasBadges}</span>
    </label>`;
  }).join('');
}

function moverEtapaDesenho(btn, dir) {
  const label = btn.closest('.etapa-check');
  if (!label) return;
  if (dir < 0) {
    const prev = label.previousElementSibling;
    if (prev && prev.classList.contains('etapa-check')) label.parentNode.insertBefore(label, prev);
  } else {
    const next = label.nextElementSibling;
    if (next && next.classList.contains('etapa-check')) label.parentNode.insertBefore(next, label);
  }
}

function moverEtapaForm(btn, dir) {
  const label = btn.closest('.etapa-check');
  if (!label) return;
  if (dir < 0) {
    const prev = label.previousElementSibling;
    if (prev && prev.classList.contains('etapa-check')) label.parentNode.insertBefore(label, prev);
  } else {
    const next = label.nextElementSibling;
    if (next && next.classList.contains('etapa-check')) label.parentNode.insertBefore(next, label);
  }
}

function aplicarOrdemEtapas(ordemNomes) {
  const cont = document.getElementById('etapas-container');
  if (!cont || !Array.isArray(ordemNomes) || !ordemNomes.length) return;
  const labels = Array.from(cont.querySelectorAll('.etapa-check'));
  const porNome = new Map(labels.map(l => [l.querySelector('input').value, l]));
  const usadas = new Set();
  ordemNomes.forEach(nome => {
    const l = porNome.get(nome);
    if (l) { cont.appendChild(l); usadas.add(nome); }
  });
  labels.forEach(l => {
    const nome = l.querySelector('input').value;
    if (!usadas.has(nome)) cont.appendChild(l);
  });
}

function addEtapaCustomizada() {
  openCadastroModal('etapa', null, 'os-form');
}

function renderEnfestoBlocos(n, prefills = []) {
  const cont = document.getElementById('f-enfestos-blocos');
  if (!cont) return;
  const qtd = Math.max(1, n || 1);
  cont.innerHTML = '';
  for (let i = 0; i < qtd; i++) {
    const p = prefills[i] || {};
    // Retrocompat: se tinha "Tecido · Cor" salvo em nomeTecido sem nomeCor, separa
    let nomeTecido = p.nomeTecido || '';
    let nomeCor = p.nomeCor || '';
    if (!nomeCor && nomeTecido.includes(' · ')) {
      const [t, ...rest] = nomeTecido.split(' · ');
      nomeTecido = t;
      nomeCor = rest.join(' · ');
    }
    const bloco = document.createElement('div');
    bloco.className = 'enfesto-bloco';
    bloco.dataset.nomeTecido = nomeTecido;
    bloco.dataset.nomeCor = nomeCor;
    bloco.style.cssText = 'margin-bottom:8px;padding:8px;border:1px solid var(--line);border-radius:2px;background:var(--line-2);';
    const labelDisplay = [nomeTecido, nomeCor].filter(Boolean).join(' · ');
    // Regra: fase Viés sempre tem 1 camada
    const ehVies = /vi[eé]s/i.test(nomeTecido);
    const camadasValue = ehVies ? '1' : (p.camadas || '');
    const camadasAttrs = ehVies ? 'readonly title="Fase Viés sempre tem 1 camada"' : '';
    bloco.innerHTML = `
      <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;color:var(--ink);margin-bottom:6px;letter-spacing:.08em;">
        ENFESTO ${i+1}${labelDisplay ? ` · <span style="color:var(--ink-2);font-weight:500;">${esc(labelDisplay)}</span>` : ''}
      </div>
      <div class="form-grid cols-3">
        <div class="field"><label>Comprimento (m)</label><input type="number" step="0.01" class="enf-comp" data-idx="${i}" value="${esc(p.comp||'')}" placeholder="Ex.: 6,50"></div>
        <div class="field"><label>Largura (m)</label><input type="number" step="0.01" class="enf-larg" data-idx="${i}" value="${esc(p.larg||'')}" placeholder="Ex.: 1,80"></div>
        <div class="field"><label>Camadas</label><input type="number" min="0" step="1" class="enf-camadas" data-idx="${i}" value="${esc(camadasValue)}" ${camadasAttrs} placeholder="—" oninput="atualizarCalculosEnfesto()"></div>
      </div>`;
    cont.appendChild(bloco);
  }
}

function lerEnfestoBlocos() {
  const cont = document.getElementById('f-enfestos-blocos');
  if (!cont) return [];
  return Array.from(cont.querySelectorAll('.enfesto-bloco')).map((b, i) => {
    const nomeTecido = b.dataset.nomeTecido || '';
    const ehVies = /vi[eé]s/i.test(nomeTecido);
    const camadasInput = parseInt(b.querySelector('.enf-camadas')?.value) || 0;
    return {
      ordem: i + 1,
      nomeTecido,
      nomeCor: b.dataset.nomeCor || '',
      comp: parseFloat(b.querySelector('.enf-comp').value) || 0,
      larg: parseFloat(b.querySelector('.enf-larg').value) || 0,
      camadas: ehVies ? 1 : camadasInput
    };
  });
}

function addTecidoRow(data = {}) {
  const cont = document.getElementById('tecidos-rows');
  const idx = cont.children.length + 1;
  if (idx > 5) { toast('Máximo 5 tecidos', 'err'); return; }
  const corOpts = '<option value="">—</option>' + STATE.cores.map(c =>
    `<option value="${esc(c.id)}" ${data.corId===c.id?'selected':''}>${esc(c.nome)}</option>`).join('');
  const row = document.createElement('div');
  row.className = 'tecido-row';
  row.innerHTML = `
    <div class="field"><label>Nº</label><input type="text" value="${idx}" readonly style="text-align:center;background:var(--line-2)"></div>
    <div class="field"><label>Tecido</label><select class="tec-sel" onchange="atualizarCalculosEnfesto()">${tecOptions(data.tecidoId)}</select></div>
    <div class="field"><label>Cor</label><select class="tec-cor">${corOpts}</select></div>
    <div class="field">
      <label>Consumo C.1</label>
      <div style="display:flex; gap:4px;">
        <input type="text" class="tec-c1" value="${esc(data.c1||'')}" placeholder="0,000 kg" style="flex:1">
        <button type="button" class="btn small danger" onclick="this.closest('.tecido-row').remove(); reindexTecidos()">✕</button>
      </div>
    </div>`;
  cont.appendChild(row);
}
function modeloCategoriaAtual() {
  const modeloId = document.getElementById('f-modelo')?.value;
  if (!modeloId) return null;
  const m = STATE.modelos.find(x => x.id === modeloId);
  return m?.categoria || null;
}

function tecOptions(selId) {
  const cat = modeloCategoriaAtual();
  const tecs = STATE.tecidos.filter(t => {
    if (!cat) return true;
    if (!t.categoria) return true;
    if (t.categoria === cat) return true;
    if (t.id === selId) return true;
    return false;
  });
  return '<option value="">—</option>' + tecs.map(t =>
    `<option value="${esc(t.id)}" ${selId===t.id?'selected':''}>${esc(t.nome)}</option>`).join('');
}

let aplicandoVinculosDesenho = false;

function aplicarFiltroTecidosPorModelo() {
  document.querySelectorAll('#tecidos-rows .tec-sel').forEach(sel => {
    const currentVal = sel.value;
    sel.innerHTML = tecOptions(currentVal);
  });
}

function aplicarVinculosModelo() {
  const modeloId = document.getElementById('f-modelo')?.value;
  if (!modeloId) return;
  const m = STATE.modelos.find(x => x.id === modeloId);
  if (!m) return;
  const mapa = {
    baseId: 'f-base', marcaId: 'f-griffe',
    designerId: 'f-designer', ftecId: 'f-ftec', coordId: 'f-coordenado'
  };
  let aplicou = false;
  Object.entries(mapa).forEach(([campo, selId]) => {
    const el = document.getElementById(selId);
    if (!el) return;
    if (m[campo]) { el.value = m[campo]; aplicou = true; }
  });
  if (aplicou) {
    atualizarResponsabilidadesOS();
    if (!aplicandoVinculosDesenho) toast('Vínculos do modelo aplicados', 'ok');
  }
}

function onModeloChange() {
  aplicarFiltroTecidosPorModelo();
  if (!aplicandoVinculosDesenho) aplicarVinculosModelo();
  preencherDropdownGradesOS();
  atualizarCalculosEnfesto();
}
function reindexTecidos() {
  document.querySelectorAll('#tecidos-rows .tecido-row').forEach((r, i) => {
    r.querySelector('input[readonly]').value = i + 1;
  });
}

function addVarianteRow(data = {}) {
  const cont = document.getElementById('variantes-rows');
  const idx = cont.children.length + 1;
  if (idx > 4) { toast('Máximo 4 variantes', 'err'); return; }
  const row = document.createElement('div');
  row.className = 'variante-row';
  row.innerHTML = `
    <div class="field"><input type="text" value="Var ${idx}" readonly style="text-align:center;background:var(--line-2)"></div>
    <div class="field"><select class="var-c1">${corOptions(data.cor1)}</select></div>
    <div class="field"><select class="var-c2">${corOptions(data.cor2)}</select></div>
    <div class="field"><select class="var-c3">${corOptions(data.cor3)}</select></div>
    <div class="field" style="display:flex;gap:4px;">
      <input type="text" class="var-obs" value="${esc(data.obs||'')}" placeholder="observação">
      <button type="button" class="btn small danger" onclick="this.closest('.variante-row').remove(); reindexVariantes()">✕</button>
    </div>`;
  cont.appendChild(row);
}
function corOptions(selId) {
  return '<option value="">—</option>' + STATE.cores.map(c =>
    `<option value="${esc(c.id)}" ${selId===c.id?'selected':''}>${esc(c.nome)}</option>`).join('');
}
function reindexVariantes() {
  document.querySelectorAll('#variantes-rows .variante-row').forEach((r, i) => {
    r.querySelector('input[readonly]').value = 'Var ' + (i + 1);
  });
}

function addComponenteRow(data = {}) {
  const cont = document.getElementById('componentes-rows');
  const row = document.createElement('div');
  row.className = 'componente-row';
  const fonteComponentes = STATE.componentes.length
    ? STATE.componentes.map(c => c.nome)
    : STATE.componentesPadrao;
  const compOpts = fonteComponentes.map(c =>
    `<option value="${esc(c)}" ${data.nome===c?'selected':''}>${esc(c)}</option>`).join('');
  const todosTecidos = [...STATE.tecidos.map(t=>({id:'T:'+t.id, nome:t.nome})),
                       ...STATE.materiais.map(m=>({id:'M:'+m.id, nome:m.codigo+' · '+m.desc}))];
  const matOpts = '<option value="">—</option>' + todosTecidos.map(t =>
    `<option value="${esc(t.id)}" ${data.material===t.id?'selected':''}>${esc(t.nome)}</option>`).join('');
  // Retrocompat: se OS antiga tem cor1/cor2/cor3, usa cor1 como única cor exibida
  const corSel = data.cor || data.cor1 || '';
  const corOpts = '<option value="">—</option>' + STATE.cores.map(c =>
    `<option value="${esc(c.id)}" ${corSel===c.id?'selected':''}>${esc(c.nome)}</option>`).join('');
  row.innerHTML = `
    <div class="field">
      <input list="compList" class="comp-nome" value="${esc(data.nome||'')}" placeholder="Componente" onchange="expandirCoresComponente(this)">
      <datalist id="compList">${compOpts}</datalist>
    </div>
    <div class="field"><select class="comp-mat">${matOpts}</select></div>
    <div class="field"><select class="comp-cor">${corOpts}</select></div>
    <div class="field" style="display:flex;gap:4px;">
      <input type="number" class="comp-qtd" min="0" step="0.5" value="${esc(data.qtdPorPeca!=null?data.qtdPorPeca:'')}" placeholder="1" style="flex:1">
      <button type="button" class="btn small danger" onclick="this.closest('.componente-row').remove()">✕</button>
    </div>`;
  cont.appendChild(row);
}

function expandirCoresComponente(inputEl) {
  if (!inputEl) return;
  const nome = (inputEl.value || '').trim();
  if (!nome) return;
  const cad = STATE.componentes.find(x => (x.nome || '').toLowerCase() === nome.toLowerCase());
  if (!cad) return;
  const coresCad = [cad.cor1Id, cad.cor2Id, cad.cor3Id].filter(Boolean);
  if (!coresCad.length) return;
  const row = inputEl.closest('.componente-row');
  if (!row) return;
  const corSel = row.querySelector('.comp-cor');
  if (!corSel) return;
  // Se já há uma cor escolhida, mantém (respeita escolha manual e cor vinda do desenho)
  if (corSel.value) return;
  // Preenche com a primeira cor cadastrada no componente
  corSel.value = coresCad[0];
}

function addAviamentoRow(data = {}) {
  const cont = document.getElementById('aviamentos-rows');
  const row = document.createElement('div');
  row.className = 'componente-row';
  const matOpts = '<option value="">—</option>' + STATE.materiais.map(m =>
    `<option value="${esc(m.id)}" ${data.material===m.id?'selected':''}>${esc(m.codigo)} · ${esc(m.desc)}</option>`).join('');
  row.innerHTML = `
    <div class="field"><select class="av-mat">${matOpts}</select></div>
    <div class="field"><input type="text" class="av-app" value="${esc(data.app||'')}" placeholder="Ex.: V1: Camel / V2: Preto"></div>
    <div class="field"><input type="number" class="av-qtd" min="0" step="0.5" value="${esc(data.qtd!=null?data.qtd:'')}" placeholder="Qtd/peça"></div>
    <div class="field" style="display:flex;gap:4px;">
      <span style="padding:7px 6px;font-size:12px;color:var(--ink-3);flex:1;">un</span>
      <button type="button" class="btn small danger" onclick="this.closest('.componente-row').remove()">✕</button>
    </div>`;
  cont.appendChild(row);
}

function aplicarGradePreset() {
  const id = document.getElementById('f-grade-preset').value;
  if (!id) return;
  const g = STATE.grades.find(x => x.id === id);
  if (!g) return;
  const t = g.tamanhos || {};
  ['p','m','g','gg','g1','g2','g3'].forEach(k => {
    document.getElementById('f-gr-'+k).value = t[k] || 0;
  });
  document.getElementById('f-grade-desc').value = g.nome;

  const fases = Array.isArray(g.fases) ? g.fases : [];

  // Monta dicionário por ordem pra preservar posicionamento (fase 2 → bloco 2)
  const porOrdem = {};
  fases.forEach(f => { if (f.ordem) porOrdem[f.ordem] = f; });
  const ordens = Object.keys(porOrdem).map(Number);
  const maxOrd = ordens.length ? Math.max(...ordens) : 0;

  // Desenho atualmente selecionado na OS (fornece cor quando a fase não tem)
  const desenhoAtual = (() => {
    const id = document.getElementById('f-desenho')?.value;
    return id ? STATE.desenhos.find(x => x.id === id) : null;
  })();

  // Pré-calcula papéis de cada fase pra orientar o fallback de cor.
  // Construído antes do bloco do enfesto pra ser reaproveitado nas linhas de Tecidos.
  const fasesOrd = [];
  for (let n = 1; n <= maxOrd; n++) fasesOrd.push(porOrdem[n] || {});
  const papeisFases = maxOrd > 0 ? calcularPapeisFases(fasesOrd) : [];

  // Cor de um componente específico do desenho (matching por papel + nome).
  // Em camiseta (sem moletom no grade) ribana_1 não é Punhos, é Gola — então
  // o matcher tenta 'punho' primeiro e, se não casar, cai pra 'gola'.
  //   forro_capuz  → "forro"
  //   ribana_1     → "punho" → "gola"
  //   ribana_2     → "barra" → "gola"
  //   ribana_3+    → "cobre" / "gola" / "ribana"
  //   moletom/malha (corpo) → frente / costas / capuz / manga
  //   sem papel    → 1º componente com mesmo tecidoId da fase
  const corDeComponente = (papel, f) => {
    if (!desenhoAtual) return null;
    const componentes = Array.isArray(desenhoAtual.componentes) ? desenhoAtual.componentes : [];
    if (!componentes.length) return null;
    const comps = componentes.map(c => ({
      ...c,
      nome: c.nome || (STATE.componentes.find(x => x.id === c.componenteId) || {}).nome || ''
    }));
    const norm = s => (s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    const pickByName = (kws) => comps.find(c => kws.some(k => norm(c.nome).includes(k)))?.corId || null;
    if (papel === 'forro_capuz')  return pickByName(['forro']);
    if (papel === 'ribana_1')     return pickByName(['punho']) || pickByName(['gola']);
    if (papel === 'ribana_2')     return pickByName(['barra']) || pickByName(['gola']);
    if (papel?.startsWith('ribana_')) return pickByName(['cobre', 'gola', 'ribana']);
    if (papel === 'moletom' || papel === 'malha') return pickByName(['frente', 'costas', 'capuz', 'manga']);
    if (f?.tecidoId) {
      return comps.find(c => c.tecidoId === f.tecidoId)?.corId || null;
    }
    return null;
  };

  // Lista de ordens das fases de CORPO (moletom ou malha sem moletom).
  // Usada para mapear cor primária/secundária/terciária às 1ª/2ª/3ª body fases,
  // independente da posição absoluta (assim acessórios entre as body fases
  // não consomem cor topo-nível).
  const bodyOrdems = [];
  for (let n = 1; n <= maxOrd; n++) {
    const p = papeisFases[n-1]?.papel || '';
    if (p === 'moletom' || p === 'malha') bodyOrdems.push(n);
  }

  // Cor por fase no Enfesto:
  //   - Body fase (moletom/malha-corpo): 1ª body → corPrimária, 2ª → corSecund.,
  //     3ª → corTerciária. Se a posição não tiver cor cadastrada, cai pra
  //     componente correspondente.
  //   - Acessório (forro_capuz, ribana_*): SEMPRE cor do componente
  //     correspondente do desenho (gola, forro, punho, barra, etc.).
  //   - Sem componente correspondente → f.corId da fase da grade → vazio.
  const corPorFase = (n, papel, f) => {
    if (papel === 'moletom' || papel === 'malha') {
      const idx = bodyOrdems.indexOf(n);
      // 1ª fase de corpo → 1ª cor do desenho, 2ª → 2ª, 3ª → 3ª. As cores seguem a
      // sequência canônica do desenho (desc), então a fase acompanha a cor certa
      // mesmo que os campos corPrincipal/Sec/Ter estejam numa ordem divergente.
      const cores = ordenarCoresIdsPorDesc([
        desenhoAtual?.corPrincipalId,
        desenhoAtual?.corSecundariaId,
        desenhoAtual?.corTerciariaId
      ].filter(Boolean), desenhoAtual);
      if (idx >= 0 && idx < cores.length && cores[idx]) return cores[idx];
    }
    return corDeComponente(papel, f) || f.corId || '';
  };

  // Renderiza blocos de Enfesto — um por fase na ordem cadastrada (pode ter blocos vazios no meio)
  if (maxOrd > 0) {
    const prefills = [];
    for (let n = 1; n <= maxOrd; n++) {
      const f = porOrdem[n] || {};
      const papel = papeisFases[n-1] || { label: '', papel: '' };
      const corIdEfetiva = corPorFase(n, papel.papel, f);
      const cor = corIdEfetiva ? STATE.cores.find(c => c.id === corIdEfetiva) : null;
      prefills.push({
        comp: f.comp || '',
        larg: f.larg || '',
        nomeTecido: papel.label || '',
        nomeCor: cor?.nome || ''
      });
    }
    renderEnfestoBlocos(maxOrd, prefills);
  } else if (g.enfestoComprimento || g.enfestoLargura) {
    renderEnfestoBlocos(1, [{ comp: g.enfestoComprimento, larg: g.enfestoLargura }]);
  }

  // Popula linhas de Tecido com tecido + cor de cada fase, na ordem cadastrada.
  // Mesma regra do Enfesto: fase 1-3 → cor primária/secundária/terciária do
  // desenho; fase 4+ → cor do componente.
  if (fases.length && fases.some(f => f.tecidoId || f.corId)) {
    const tecCont = document.getElementById('tecidos-rows');
    if (tecCont) {
      tecCont.innerHTML = '';
      for (let n = 1; n <= maxOrd; n++) {
        const f = porOrdem[n] || {};
        const papel = papeisFases[n-1] || { papel: '' };
        const corIdEfetiva = corPorFase(n, papel.papel, f);
        if (f.tecidoId || corIdEfetiva) {
          addTecidoRow({ tecidoId: f.tecidoId || '', corId: corIdEfetiva });
        }
      }
    }
  }

  // Popula Variante 1 com cores — fase ordem=1 → var-c1, ordem=2 → var-c2, ordem=3 → var-c3
  if (fases.length && fases.some(f => f.corId)) {
    const varCont = document.getElementById('variantes-rows');
    if (varCont) {
      if (!varCont.querySelector('.variante-row')) addVarianteRow();
      const primeira = varCont.querySelector('.variante-row');
      if (primeira) {
        const slots = ['.var-c1', '.var-c2', '.var-c3'];
        fases.forEach(f => {
          const n = f.ordem;
          if (n >= 1 && n <= 3 && f.corId) {
            const sel = primeira.querySelector(slots[n-1]);
            if (sel) sel.value = f.corId;
          }
        });
      }
    }
  }

  atualizarCalculosEnfesto();
  const n = fases.length;
  const msg = n <= 1 ? 'Grade aplicada' : `Grade aplicada — Fase 1 no enfesto, ${n} fases no total`;
  toast(msg, 'ok');
}

/* ========================================================= */
/*              ENFESTO — limites e cálculos                 */
/* ========================================================= */
const LIMITE_CAMADAS = { malha: 80, moletom: 36, ribana: 80, outro: Infinity };
const MULTIPLICADOR_PECAS = { malha: 2, moletom: 1, ribana: 2, outro: 1 };
const LABEL_CATEGORIA = { malha: 'Malha algodão', moletom: 'Moletom', ribana: 'Ribana', outro: 'Outro' };

/**
 * Categoria efetiva de um tecido: respeita a categoria cadastrada, mas
 * se o nome contém "ribana" (case-insensitive), força 'ribana'. Isso cobre
 * tecidos cadastrados antes da categoria "ribana" existir (ex.: "Ribana moletom"
 * salvo como categoria=moletom).
 */
function categoriaEfetivaTecido(t) {
  if (!t) return '';
  if ((t.nome || '').toLowerCase().includes('ribana')) return 'ribana';
  return t.categoria || '';
}

// Qualquer tecido cuja categoria efetiva seja "ribana" (Ribana Moletom, Ribana
// Malha Algodao, Ribana Gola Polo, etc.). Quando a fase do enfesto usa um
// tecido ribana, o campo "Unidades da grade" e habilitado e o calculo de
// camadas usa a regra simples camadasMoletom / unidades em vez de
// MULTIPLICADOR_PECAS.ribana.
function isTecidoRibana(t) {
  return !!t && categoriaEfetivaTecido(t) === 'ribana';
}

// Categoria principal de uma grade — categoria do tecido da fase de menor `ordem`.
// Usada pra filtrar o dropdown de grades pelo tecido do desenho selecionado.
function categoriaPrincipalGrade(g) {
  const fases = Array.isArray(g?.fases) ? g.fases : [];
  if (!fases.length) return '';
  const ordenadas = [...fases].sort((a, b) => (a.ordem || 0) - (b.ordem || 0));
  for (const f of ordenadas) {
    const t = STATE.tecidos.find(x => x.id === f.tecidoId);
    const cat = categoriaEfetivaTecido(t);
    if (cat) return cat;
  }
  return '';
}

// Categoria do tecido principal do desenho selecionado no form da OS.
function categoriaDesenhoOS() {
  const id = document.getElementById('f-desenho')?.value;
  if (!id) return '';
  const d = STATE.desenhos.find(x => x.id === id);
  if (!d?.tecidoPadraoId) return '';
  const t = STATE.tecidos.find(x => x.id === d.tecidoPadraoId);
  return categoriaEfetivaTecido(t);
}

// tipoPeca esperado para a grade, conforme o modelo selecionado na OS.
// Mapeia modelo.categoria → grade.tipoPeca:
//   camiseta → camiseta
//   moletom  → blusa_moletom
//   outro    → outro
function tipoPecaModeloOS() {
  const id = document.getElementById('f-modelo')?.value;
  if (!id) return '';
  const m = STATE.modelos.find(x => x.id === id);
  const cat = m?.categoria || '';
  // modelo.categoria usa 'malha'/'moletom'/'outro'; grade.tipoPeca usa
  // 'camiseta'/'blusa_moletom'/'outro'. Mapeia entre os dois.
  if (cat === 'moletom') return 'blusa_moletom';
  if (cat === 'malha') return 'camiseta';
  return cat; // 'outro' / ''
}

// Variacao implícita do desenho (não há campo dedicado): se tem cor terciária
// → tricolor; cor secundária → bicolor; só principal → basica.
function variacaoDesenhoOS() {
  const id = document.getElementById('f-desenho')?.value;
  if (!id) return '';
  const d = STATE.desenhos.find(x => x.id === id);
  if (!d) return '';
  if (d.corTerciariaId) return 'tricolor';
  if (d.corSecundariaId) return 'bicolor';
  if (d.corPrincipalId) return 'basica';
  return '';
}

// Grades que devem aparecer no dropdown de "Carregar grade pré-cadastrada" da OS,
// filtradas pela categoria do tecido do desenho, tipoPeca casado com o modelo
// e variação (basica/bicolor/tricolor) casada com o número de cores do desenho.
// `extraIds` mantém grades específicas (geralmente a já selecionada) mesmo fora
// do filtro. Grades sem tipoPeca/variacao cadastrados passam pelos respectivos
// filtros (não tem como avaliar) — mas continuam sujeitas aos demais.
function gradesParaDropdownOS(extraIds = []) {
  const cat = categoriaDesenhoOS();
  const tipoModelo = tipoPecaModeloOS();
  const variacao = variacaoDesenhoOS();
  const keep = new Set(extraIds.filter(Boolean));
  // Grades conjugadas ficam ocultas do dropdown — usadas internamente pelo
  // fluxo de auto-geracao (Camiseta Bicolor → Basica conjugada). Continua
  // visivel se for a grade ja salva da OS em edicao (via extraIds).
  const ocultaConjugada = (g) => /conjug/i.test(g.nome || '');
  if (!cat && !tipoModelo && !variacao) {
    return STATE.grades.filter(g => keep.has(g.id) || !ocultaConjugada(g));
  }
  return STATE.grades.filter(g => {
    if (keep.has(g.id)) return true;
    if (ocultaConjugada(g)) return false;
    if (cat && categoriaPrincipalGrade(g) !== cat) return false;
    if (tipoModelo && g.tipoPeca && g.tipoPeca !== tipoModelo) return false;
    if (variacao && g.variacao && g.variacao !== variacao) return false;
    return true;
  });
}

function preencherDropdownGradesOS() {
  const el = document.getElementById('f-grade-preset');
  if (!el) return;
  const cur = el.value || '';
  fillSelect('f-grade-preset', gradesParaDropdownOS([cur]), 'nome', '— nenhuma —');
  if (cur) el.value = cur;
}

/**
 * Determina papel/nome de cada fase em função do tecido e da posição na grade.
 * - Fase com moletom → "Moletom"
 * - Fase com malha, SE a grade também tem moletom → "Forro de capuz"
 * - Fase com ribana → 1ª = "Punhos", 2ª = "Barra", demais = "Ribana N"
 * - Fallback: nome da categoria
 * Retorna array paralelo a `fases` com { papel, label, categoria }.
 */
function calcularPapeisFases(fases) {
  const tecidosMap = new Map(STATE.tecidos.map(t => [t.id, t]));
  const temMoletom = fases.some(f => categoriaEfetivaTecido(tecidosMap.get(f.tecidoId)) === 'moletom');
  let contRib = 0;
  return fases.map(f => {
    const t = tecidosMap.get(f.tecidoId);
    const cat = categoriaEfetivaTecido(t);
    // Papel é sempre calculado pela categoria/posição (usado pra agrupar totais)
    let papel, labelAuto;
    if (cat === 'moletom') { papel = 'moletom'; labelAuto = 'Moletom'; }
    else if (cat === 'malha' && temMoletom) { papel = 'forro_capuz'; labelAuto = 'Forro de capuz'; }
    else if (cat === 'ribana') {
      contRib++;
      papel = 'ribana_'+contRib;
      labelAuto = contRib === 1 ? 'Punhos' : contRib === 2 ? 'Barra' : `Ribana ${contRib}`;
    } else {
      papel = cat || 'outro';
      labelAuto = LABEL_CATEGORIA[cat] || (t?.nome || '');
    }
    // Nome cadastrado pelo user na fase tem prioridade sobre o label automático
    const label = (f?.nome && f.nome.trim()) ? f.nome.trim() : labelAuto;
    return { papel, label, categoria: cat };
  });
}

function multiplicadorDominante() {
  const rows = document.querySelectorAll('#tecidos-rows .tec-sel');
  let mult = 1;
  rows.forEach(sel => {
    if (!sel.value) return;
    const tec = STATE.tecidos.find(t => t.id === sel.value);
    if (!tec || !tec.categoria) return;
    const m = MULTIPLICADOR_PECAS[tec.categoria] || 1;
    if (m > mult) mult = m;
  });
  return mult;
}

// Calcula o limite máximo de camadas baseado nos tecidos selecionados no formulário.
// Pega o menor limite entre todos — se há moletom e malha, vence moletom (36).
function calcularLimiteCamadas() {
  const rows = document.querySelectorAll('#tecidos-rows .tecido-row');
  let limite = Infinity;
  let categoriaRestritiva = null;
  rows.forEach(r => {
    const sel = r.querySelector('.tec-sel');
    if (!sel || !sel.value) return;
    const tec = STATE.tecidos.find(t => t.id === sel.value);
    if (!tec || !tec.categoria) return;
    const lim = LIMITE_CAMADAS[tec.categoria];
    if (lim < limite) { limite = lim; categoriaRestritiva = tec.categoria; }
  });
  return { limite, categoriaRestritiva };
}

function atualizarCalculosEnfesto() {
  const gradeTotal = ['p','m','g','gg','g1','g2','g3']
    .reduce((s, k) => s + (parseInt(document.getElementById('f-gr-'+k)?.value) || 0), 0);
  const camadas = parseInt(document.getElementById('f-enf-camadas')?.value) || 0;
  const { limite, categoriaRestritiva } = calcularLimiteCamadas();

  // Atualiza a dica ao lado do campo de camadas
  const info = document.getElementById('f-enf-limite-info');
  if (info) {
    if (limite === Infinity) {
      info.textContent = '';
    } else {
      const label = categoriaRestritiva === 'moletom' ? 'Moletom' : 'Malha algodão';
      info.textContent = `· máx ${limite} (${label})`;
    }
  }

  // Validação visual
  const alerta = document.getElementById('f-enf-alerta');
  const campoCamadas = document.getElementById('f-enf-camadas');
  if (camadas > limite) {
    alerta.textContent = `⚠ Você informou ${camadas} camadas, mas o limite para ${categoriaRestritiva === 'moletom' ? 'moletom' : 'malha algodão'} é ${limite}. Ajuste o valor ou separe em mais de um enfesto.`;
    alerta.classList.remove('hidden');
    campoCamadas.style.borderColor = 'var(--alert)';
    campoCamadas.style.background = '#fff5f5';
  } else {
    alerta.classList.add('hidden');
    campoCamadas.style.borderColor = '';
    campoCamadas.style.background = '';
  }

  // Área de cálculo — separado por categoria de tecido
  const calcBox = document.getElementById('f-enf-calculo');
  if (camadas > 0 && gradeTotal > 0) {
    // Coleta categorias das fases da grade atual; fallback: linhas de Tecidos do form
    const categoriasUsadas = new Set();
    const gradeId = document.getElementById('f-grade-preset')?.value;
    const grade = gradeId ? STATE.grades.find(g => g.id === gradeId) : null;
    const fases = grade?.fases || [];
    fases.forEach(f => {
      if (!f.tecidoId) return;
      const t = STATE.tecidos.find(x => x.id === f.tecidoId);
      if (t?.categoria) categoriasUsadas.add(t.categoria);
    });
    if (!categoriasUsadas.size) {
      document.querySelectorAll('#tecidos-rows .tec-sel').forEach(sel => {
        if (!sel.value) return;
        const tec = STATE.tecidos.find(t => t.id === sel.value);
        if (tec?.categoria) categoriasUsadas.add(tec.categoria);
      });
    }
    // Coletar fases da grade selecionada (se houver) pra calcular papéis
    const gradeIdSel = document.getElementById('f-grade-preset')?.value;
    const gradeSel = gradeIdSel ? STATE.grades.find(g => g.id === gradeIdSel) : null;
    const fasesGrade = gradeSel?.fases || [];
    const papeis = calcularPapeisFases(fasesGrade);

    // Se nem fases nem tecidos, mostra total genérico
    if (!categoriasUsadas.size) {
      const mult = multiplicadorDominante();
      const totalPecas = gradeTotal * camadas * mult;
      calcBox.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
          <span>Total de peças produzidas:</span>
          <strong style="font-family:'IBM Plex Mono', monospace; font-size: 15px; color: var(--accent-dark);">${totalPecas} peças</strong>
        </div>`;
    } else {
      // Agrupa fases por "grupo de total": moletom / forro_capuz / ribana (combina punhos+barras)
      const temMoletom = papeis.some(p => p.papel === 'moletom');
      const temForro = papeis.some(p => p.papel === 'forro_capuz');
      const temRibana = papeis.some(p => (p.papel || '').startsWith('ribana_'));

      // Fallback pras categorias encontradas nos tecidos-rows (sem grade)
      if (!papeis.length) {
        const grupos = [];
        if (categoriasUsadas.has('moletom')) grupos.push({ papel: 'moletom', label: 'Moletom' });
        if (categoriasUsadas.has('malha')) grupos.push({ papel: 'forro_capuz', label: 'Forro de capuz' });
        if (categoriasUsadas.has('ribana')) grupos.push({ papel: 'ribana', label: 'Punhos e Barras' });
        if (!grupos.length) grupos.push({ papel: 'outro', label: 'Total' });
        // Renderização básica
        const linhas = grupos.map(gr => {
          const cat = gr.papel === 'forro_capuz' ? 'malha' : gr.papel === 'ribana' ? 'ribana' : gr.papel;
          const mult = MULTIPLICADOR_PECAS[cat] || 1;
          const total = gradeTotal * camadas * mult;
          return `<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px dashed var(--line);">
            <span>Total ${gr.label}:</span>
            <strong style="font-family:'IBM Plex Mono', monospace; font-size: 15px; color: var(--accent-dark);">${total} peças</strong>
          </div>`;
        }).join('');
        calcBox.innerHTML = linhas;
      } else {
        // Total moletom: soma fases com papel moletom (mesma grade × camadas)
        const totalMoletom = temMoletom ? (gradeTotal * camadas * (MULTIPLICADOR_PECAS.moletom || 1)) : 0;

        // Busca desenho selecionado pra analisar componentes
        const desenhoIdCalc = document.getElementById('f-desenho')?.value;
        const desenhoCalc = desenhoIdCalc ? STATE.desenhos.find(x => x.id === desenhoIdCalc) : null;
        const compsCalc = Array.isArray(desenhoCalc?.componentes) ? desenhoCalc.componentes : [];

        // Total forro de capuz: só componentes cujo NOME contém a palavra-chave do label da fase
        // (ex.: label "Forro de capuz" → palavra "forro" → só "Forro do capuz" entra, Viés fica de fora)
        let forroInfo = null;
        if (temForro) {
          const labelForroFase = papeis.find(p => p.papel === 'forro_capuz')?.label || 'Forro';
          const keyForro = (labelForroFase || '').toLowerCase().split(/\s+/)[0].replace(/s$/, '');
          const forroComps = compsCalc.filter(c => {
            const tec = STATE.tecidos.find(t => t.id === c.tecidoId);
            if (!tec || categoriaEfetivaTecido(tec) !== 'malha') return false;
            const nome = (c.nome || '').toLowerCase();
            return keyForro && nome.includes(keyForro);
          });
          let qtyForro = 0;
          const detalhesF = [];
          forroComps.forEach(c => {
            const v = parseFloat(c.qtdPorPeca);
            const qty = v > 0 ? v : 1;
            qtyForro += qty;
            detalhesF.push(`${c.nome || '?'} ×${qty}`);
          });
          if (qtyForro === 0) qtyForro = 1; // fallback
          const multF = 2;
          const refF = totalMoletom;
          const camadasF = gradeTotal > 0 ? Math.ceil(refF / (gradeTotal * multF)) : 0;
          forroInfo = { total: refF * qtyForro, camadas: camadasF, detalhes: detalhesF };
        }
        const totalForro = forroInfo ? forroInfo.total : 0;

        // Ribana: um total POR FASE ribana (Punhos, Barra, etc.)
        // Componentes ribana do desenho são agrupados por palavra-chave no nome.
        let ribanaPorFase = [];
        if (temRibana) {
          const referencia = totalMoletom || totalForro;
          const comps = compsCalc;
          const ribanaComps = comps.filter(c => {
            if (!c.tecidoId) return false;
            const tec = STATE.tecidos.find(t => t.id === c.tecidoId);
            return tec && categoriaEfetivaTecido(tec) === 'ribana';
          });
          // Labels das fases ribana na ordem cadastrada na grade
          const labelsFasesRib = papeis.filter(p => (p.papel || '').startsWith('ribana_')).map(p => p.label);
          // Inicializa grupos
          const grupos = labelsFasesRib.map(lbl => ({ label: lbl, qty: 0, detalhes: [] }));
          const sobra = [];
          // Classifica cada componente no grupo cujo label combine com o nome
          ribanaComps.forEach(c => {
            const nomeLow = (c.nome || '').toLowerCase();
            const qty = (nomeLow.includes('manga') || nomeLow.includes('punho')) ? 2 : 1;
            let grupo = grupos.find(g => {
              const key = (g.label || '').toLowerCase().replace(/s$/, '');
              return key && nomeLow.includes(key);
            });
            if (!grupo && grupos.length === 1) grupo = grupos[0]; // única fase ribana pega tudo
            if (grupo) {
              grupo.qty += qty;
              grupo.detalhes.push(`${c.nome || '?'} ×${qty}`);
            } else {
              sobra.push({ nome: c.nome || '?', qty });
            }
          });
          const multRib = MULTIPLICADOR_PECAS.ribana || 2;
          // Multiplicador por label de fase ribana: ribana com unidades cadastradas
          // (qualquer ribana — moletom, malha algodao, gola polo) usa "unidades" da fase;
          // ribanas sem unidades cadastradas usam o multiplicador padrão (2).
          const multPorLabelRib = {};
          papeis.forEach((p, idx) => {
            if (!(p.papel || '').startsWith('ribana_')) return;
            const fase = fasesGrade[idx];
            const tec = STATE.tecidos.find(t => t.id === fase?.tecidoId);
            multPorLabelRib[p.label] = isTecidoRibana(tec)
              ? (parseInt(fase?.unidades) || multRib)
              : multRib;
          });
          // Regra de ribana:
          // - Ribana moletom: escala com unidade media da grade (2 cam moletom = 1 cam ribana).
          // - Outras ribanas (malha algodao, gola polo): so camadasPrincipal × multPrincipal / unidades.
          const nTamanhos = ['p','m','g','gg','g1','g2','g3']
            .filter(k => (parseInt(document.getElementById('f-gr-'+k)?.value) || 0) > 0).length;
          const unidadePorTamMedia = nTamanhos > 0 ? gradeTotal / nTamanhos : 1;
          const multPrincipalEnf = temMoletom
            ? 1
            : (categoriasUsadas.has('malha') ? (MULTIPLICADOR_PECAS.malha || 2) : 1);
          // Mapa label → escalaComGrade (para diferenciar moletom de outras ribanas no calculo)
          const escalaPorLabel = {};
          papeis.forEach((p, idx) => {
            if (!(p.papel || '').startsWith('ribana_')) return;
            const tec = STATE.tecidos.find(t => t.id === fasesGrade[idx]?.tecidoId);
            escalaPorLabel[p.label] = (tec?.nome || '').toLowerCase().includes('moletom');
          });
          const calcCamadasRibana = (mult, label) => {
            const fator = escalaPorLabel[label] ? unidadePorTamMedia : 1;
            return Math.max(1, Math.ceil(camadas * multPrincipalEnf * fator / mult));
          };
          // Para o "Total" em pecas, usa referencia (totalMoletom ou totalForro).
          // Se nao tiver, calcula como total de blusas.
          const referenciaTotal = referencia || (gradeTotal * camadas * multPrincipalEnf);
          ribanaPorFase = grupos
            .filter(g => g.qty > 0)
            .map(g => {
              const mult = multPorLabelRib[g.label] || multRib;
              return {
                label: g.label,
                total: referenciaTotal * g.qty,
                detalhes: g.detalhes,
                camadas: calcCamadasRibana(mult, g.label),
                mult
              };
            });
          // Se tiver componentes sem match, agrupa num fallback "Ribana (outros)"
          if (sobra.length) {
            const qtyTot = sobra.reduce((s, x) => s + x.qty, 0);
            ribanaPorFase.push({
              label: 'Ribana (outros)',
              total: referenciaTotal * qtyTot,
              detalhes: sobra.map(x => `${x.nome} ×${x.qty}`),
              camadas: calcCamadasRibana(multRib, 'Ribana (outros)'),
              mult: multRib
            });
          }
        }

        const blocos = [];
        const labelMoletom = papeis.find(p => p.papel === 'moletom')?.label || 'Moletom';
        const labelForro = papeis.find(p => p.papel === 'forro_capuz')?.label || 'Forro de capuz';

        if (temMoletom) {
          blocos.push(`
            <div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px dashed var(--line);">
              <span>Total ${esc(labelMoletom)}: <span style="font-size:11px;color:var(--ink-3);">(1 camada = 1 peça/tamanho)</span></span>
              <strong style="font-family:'IBM Plex Mono', monospace; font-size: 15px; color: var(--accent-dark);">${totalMoletom} peças</strong>
            </div>`);
        }
        if (temForro && forroInfo) {
          const hintF = forroInfo.detalhes.length
            ? ` <span style="font-size:11px;color:var(--ink-3);">(${esc(forroInfo.detalhes.join(' + '))})</span>`
            : '';
          blocos.push(`
            <div style="padding:4px 0;border-bottom:1px dashed var(--line);">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span>Total ${esc(labelForro)}:${hintF}</span>
                <strong style="font-family:'IBM Plex Mono', monospace; font-size: 15px; color: var(--accent-dark);">${forroInfo.total} peças</strong>
              </div>
              <div style="font-size:11px;color:var(--ink-3);margin-top:2px;">
                Camadas sugeridas: <strong>${forroInfo.camadas}</strong>
                (1 camada = 2 peças/tamanho; ${totalMoletom} peças ÷ ${gradeTotal*2})
              </div>
            </div>`);
        }
        if (ribanaPorFase.length) {
          const nTamHint = ['p','m','g','gg','g1','g2','g3']
            .filter(k => (parseInt(document.getElementById('f-gr-'+k)?.value) || 0) > 0).length;
          ribanaPorFase.forEach(rf => {
            const hint = rf.detalhes.length
              ? ` <span style="font-size:11px;color:var(--ink-3);">(${esc(rf.detalhes.join(' + '))})</span>`
              : '';
            const m = rf.mult || (MULTIPLICADOR_PECAS.ribana || 2);
            blocos.push(`
              <div style="padding:4px 0;border-bottom:1px dashed var(--line);">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                  <span>Total ${esc(rf.label)}:${hint}</span>
                  <strong style="font-family:'IBM Plex Mono', monospace; font-size: 15px; color: var(--accent-dark);">${rf.total} peças</strong>
                </div>
                <div style="font-size:11px;color:var(--ink-3);margin-top:2px;">
                  Camadas sugeridas: <strong>${rf.camadas}</strong>
                  (1 camada = ${m} peça${m===1?'':'s'}/tamanho × ${nTamHint} tamanho${nTamHint===1?'':'s'} = ${m * nTamHint} peças/camada)
                </div>
              </div>`);
          });
        }

        const porTamanho = ['p','m','g','gg','g1','g2','g3']
          .map(k => ({ t: k, qtd: parseInt(document.getElementById('f-gr-'+k)?.value) || 0 }))
          .filter(x => x.qtd > 0)
          .map(x => `${x.t.toUpperCase()}: ${x.qtd}×${camadas}`)
          .join(' · ');
        calcBox.innerHTML = `${blocos.join('')}
          <div style="margin-top:8px;font-size:12px; color: var(--ink-3); font-family:'IBM Plex Mono', monospace;">${porTamanho}</div>`;
      }
    }
  } else {
    calcBox.innerHTML = '<em style="color:var(--ink-3);">Preencha grade e camadas (ou peças-alvo) para ver o cálculo.</em>';
  }
}

function calcularCamadasParaProducao() {
  const target = parseInt(document.getElementById('f-enf-target')?.value) || 0;
  if (target <= 0) { atualizarCalculosEnfesto(); return; }
  const qtdsPorTamanho = ['p','m','g','gg','g1','g2','g3']
    .map(k => parseInt(document.getElementById('f-gr-'+k)?.value) || 0)
    .filter(q => q > 0);
  if (qtdsPorTamanho.length === 0) {
    toast('Preencha a grade antes', 'err');
    return;
  }
  const gradeId = document.getElementById('f-grade-preset')?.value;
  const grade = gradeId ? STATE.grades.find(g => g.id === gradeId) : null;
  const fases = grade?.fases || [];
  const temMoletom = fases.some(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId)) === 'moletom');
  const temMalha = fases.some(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId)) === 'malha');

  // Multiplicador da peça principal
  let multPrincipal = 1;
  if (!temMoletom && temMalha) multPrincipal = MULTIPLICADOR_PECAS.malha || 2;

  const minQtd = Math.min(...qtdsPorTamanho);
  const gradeTotal = qtdsPorTamanho.reduce((s, x) => s + x, 0);
  const camadasPrincipal = Math.ceil(target / (minQtd * multPrincipal));
  const blusas = gradeTotal * camadasPrincipal * multPrincipal;

  // Campo global reflete a peça principal
  const inputGlobal = document.getElementById('f-enf-camadas');
  if (inputGlobal) inputGlobal.value = camadasPrincipal;

  // Componentes do desenho selecionado
  const desenhoId = document.getElementById('f-desenho')?.value;
  const desenho = desenhoId ? STATE.desenhos.find(x => x.id === desenhoId) : null;
  const comps = Array.isArray(desenho?.componentes) ? desenho.componentes : [];

  // Retorna qty/blusa de um componente — usa qtdPorPeca cadastrado ou fallback por nome
  const qtdDoComp = c => {
    const v = parseFloat(c.qtdPorPeca);
    if (v > 0) return v;
    const n = (c.nome || '').toLowerCase();
    return (n.includes('manga') || n.includes('punho')) ? 2 : 1;
  };

  // Papéis das fases
  const papeis = calcularPapeisFases(fases);

  // qty total por blusa de componentes ribana, agrupado pelo label da fase ribana
  const ribanaLabels = papeis.filter(p => (p.papel || '').startsWith('ribana_')).map(p => p.label);
  const qtyPorLabelRibana = {};
  ribanaLabels.forEach(l => { qtyPorLabelRibana[l] = 0; });
  const compsRibana = comps.filter(c => {
    const tec = STATE.tecidos.find(t => t.id === c.tecidoId);
    return tec && categoriaEfetivaTecido(tec) === 'ribana';
  });
  compsRibana.forEach(c => {
    const nome = (c.nome || '').toLowerCase();
    const qtd = qtdDoComp(c);
    let lbl = ribanaLabels.find(l => {
      const key = (l || '').toLowerCase().replace(/s$/, '');
      return key && nome.includes(key);
    });
    if (!lbl && ribanaLabels.length === 1) lbl = ribanaLabels[0];
    if (lbl) qtyPorLabelRibana[lbl] += qtd;
  });

  // qty por blusa de componentes de forro — só os que combinam com a palavra-chave do label da fase forro
  let qtyForro = 0;
  if (temMoletom) {
    const labelForroFase = papeis.find(p => p.papel === 'forro_capuz')?.label || 'Forro';
    const keyForro = (labelForroFase || '').toLowerCase().split(/\s+/)[0].replace(/s$/, '');
    qtyForro = comps
      .filter(c => {
        const tec = STATE.tecidos.find(t => t.id === c.tecidoId);
        if (!tec || categoriaEfetivaTecido(tec) !== 'malha') return false;
        const nome = (c.nome || '').toLowerCase();
        return keyForro && nome.includes(keyForro);
      })
      .reduce((s, c) => s + qtdDoComp(c), 0);
  }

  // Preenche cada bloco conforme papel + qtd dos componentes
  const multRib = MULTIPLICADOR_PECAS.ribana || 2;
  const blocosDom = document.querySelectorAll('#f-enfestos-blocos .enfesto-bloco');
  blocosDom.forEach((bloco, i) => {
    const input = bloco.querySelector('.enf-camadas');
    if (!input) return;
    const papel = papeis[i] || {};
    const fase = fases[i] || {};
    const ehVies = /vi[eé]s/i.test(fase.nome || '') || /vi[eé]s/i.test(papel.label || '') || /vi[eé]s/i.test(bloco.dataset.nomeTecido || '');
    let val;
    if (ehVies) {
      val = 1;
    } else if (papel.papel === 'moletom') {
      // Enfesto moletom: todos componentes moletom na mesma camada → 1 camada = 1 blusa
      val = camadasPrincipal;
    } else if (papel.papel === 'forro_capuz') {
      // Enfesto forro: camadas = metade das camadas de moletom
      val = Math.max(1, Math.ceil(camadasPrincipal / 2));
    } else if ((papel.papel || '').startsWith('ribana_')) {
      // Enfesto ribana com unidades cadastradas:
      // - Ribana moletom: o tecido escala com a unidade da grade (2 barras +
      //   4 punhos por tamanho cobre 2 blusas/tam quando a grade tem 2/tam).
      //   Formula: camadasPrincipal × multPrincipal × (gradeTotal/n_tamanhos) / unidades.
      // - Outras ribanas (malha algodao, gola polo): o multiplicador da fase
      //   ja cobre toda a grade — "10x" significa "10 unidades de cada slot
      //   da grade por camada", entao o gradeTotal nao precisa de ajuste.
      //   Formula: camadasPrincipal × multPrincipal / unidades.
      const fase = fases[i] || {};
      const tecFase = STATE.tecidos.find(t => t.id === fase.tecidoId);
      if (isTecidoRibana(tecFase)) {
        const unidades = parseInt(fase.unidades) || multRib;
        const nomeFase = (tecFase?.nome || '').toLowerCase();
        const escalaComGrade = nomeFase.includes('moletom');
        let fator = 1;
        if (escalaComGrade) {
          const nTamanhos = qtdsPorTamanho.length;
          fator = nTamanhos > 0 ? gradeTotal / nTamanhos : 1;
        }
        val = Math.max(1, Math.ceil(camadasPrincipal * multPrincipal * fator / unidades));
      } else {
        // Ribana padrao (sem unidades cadastradas): usa qtdPorBlusa + mult fixo
        const q = qtyPorLabelRibana[papel.label] || 0;
        val = q > 0
          ? Math.max(1, Math.ceil(blusas * q / (gradeTotal * multRib)))
          : Math.max(1, Math.ceil(camadasPrincipal / multRib));
      }
    } else {
      const cat = papel.categoria || '';
      const mult = MULTIPLICADOR_PECAS[cat] || 1;
      val = Math.ceil(target / (minQtd * mult));
    }
    input.value = val;
  });

  atualizarCalculosEnfesto();
}

/* ========================================================= */
/*           NÚMERO DA OS — sequencial automático            */
/* ========================================================= */
function formatarNumeroOS(n) {
  return String(n).padStart(4, '0');
}

function proximoNumeroOS() {
  // Proximo numero = maior numero presente em STATE.ordens + 1. Assim,
  // se uma OS foi excluida, o numero dela fica livre pra ser reusado.
  // O counter persistido no Supabase nao e mais determinante — ele serve
  // apenas como piso de seguranca pra numeros muito antigos ja usados
  // que podem nao estar mais visiveis (ex.: backups), mas o maior
  // existente sempre ganha quando ha qualquer OS salva.
  const numeros = STATE.ordens
    .map(o => parseInt(o.os))
    .filter(n => !isNaN(n));
  const maxExistente = numeros.length ? Math.max(...numeros) : 0;
  if (maxExistente > 0) return formatarNumeroOS(maxExistente + 1);
  // Sem nenhuma OS existente, cai pro counter (caso tenha sido salvo
  // previamente em uma execucao anterior com OSs ja deletadas).
  const counterAtual = parseInt(STATE.osCounter) || 0;
  return formatarNumeroOS(counterAtual + 1);
}

async function atualizarCounterOS(numeroUsado) {
  // Mantem o counter sincronizado com o maior numero usado, util como
  // fallback quando todas as OSs sao excluidas. Nao influencia o
  // proximoNumeroOS quando ha OSs salvas — ali o max das existentes
  // ganha.
  const n = parseInt(numeroUsado);
  if (isNaN(n)) return;
  const counterAtual = parseInt(STATE.osCounter) || 0;
  if (n > counterAtual) {
    STATE.osCounter = n;
    await DB.set('osCounter', String(n));
  }
}

function previewDesenhoSelecionado() {
  const id = document.getElementById('f-desenho').value;
  const pv = document.getElementById('f-desenho-preview');
  if (!id) { pv.innerHTML = '<span>Nenhum desenho selecionado</span>'; return; }
  const d = STATE.desenhos.find(x => x.id === id);
  pv.innerHTML = d?.img ? `<img src="${d.img}">` : '<span>Sem imagem</span>';
}

// Componentes de um desenho, normalizados: a estrutura nova (d.componentes, com
// tecido+cor+qtd) tem prioridade; a antiga (só d.componentesIds) é convertida.
// Fonte única usada tanto ao aplicar o desenho quanto ao repor manualmente.
function _componentesDoDesenho(d) {
  if (!d) return [];
  const brutos = Array.isArray(d.componentes) && d.componentes.length
    ? d.componentes
    : (d.componentesIds || []).map(id => ({
        componenteId: id,
        nome: (STATE.componentes.find(x => x.id === id) || {}).nome || '',
        tecidoId: d.tecidoPadraoId || '',
        corId: d.corPrincipalId || '',
        qtdPorPeca: 1
      }));
  return brutos.map(c => {
    const cad = STATE.componentes.find(x => x.id === c.componenteId)
             || (c.nome ? STATE.componentes.find(x => x.nome === c.nome) : null);
    return {
      nome: c.nome || cad?.nome || '',
      material: c.tecidoId ? 'T:' + c.tecidoId : '',
      cor: c.corId || cad?.cor1Id || '',
      qtdPorPeca: c.qtdPorPeca != null ? c.qtdPorPeca : 1
    };
  });
}

// Repõe as linhas de Componentes a partir do desenho técnico já selecionado na
// OS. Serve para quando o desenho ganhou componentes DEPOIS de a OS ter sido
// salva: editar o desenho não altera OSs gravadas, então elas ficam com a
// seção vazia (e, por consequência, com 0 peças — somem da expedição e do
// Estoque de corte). Um clique aqui puxa os componentes do desenho vivo.
function reporComponentesDoDesenho() {
  const id = document.getElementById('f-desenho')?.value;
  if (!id) return toast('Selecione um desenho técnico primeiro', 'err');
  const d = STATE.desenhos.find(x => x.id === id);
  if (!d) return toast('Desenho não encontrado', 'err');
  const comps = _componentesDoDesenho(d);
  if (!comps.length) return toast('O desenho selecionado não tem componentes cadastrados', 'err');
  const cont = document.getElementById('componentes-rows');
  if (!cont) return;
  cont.innerHTML = '';
  comps.forEach(c => addComponenteRow(c));
  toast(`${comps.length} componente(s) repostos do desenho ${d.codigo || ''}`.trim() + '. Confira e salve a OS.', 'ok');
}
window.reporComponentesDoDesenho = reporComponentesDoDesenho;


/* ========================================================= */
/*                      SALVAR OS                            */
/* ========================================================= */
function coletaOS() {
  const v = id => document.getElementById(id)?.value || '';
  const getSel = el => ({ id: el.value, text: el.options[el.selectedIndex]?.text || '' });

  const tecidos = Array.from(document.querySelectorAll('#tecidos-rows .tecido-row')).map(r => {
    const tecSel = r.querySelector('.tec-sel');
    const corSel = r.querySelector('.tec-cor');
    return {
      tecidoId: tecSel.value,
      tecidoNome: tecSel.options[tecSel.selectedIndex]?.text || '',
      corId: corSel?.value || '',
      corNome: corSel?.options[corSel.selectedIndex]?.text || '',
      c1: r.querySelector('.tec-c1').value
    };
  }).filter(t => t.tecidoId);

  const variantes = Array.from(document.querySelectorAll('#variantes-rows .variante-row')).map((r, i) => {
    const c1 = r.querySelector('.var-c1');
    const c2 = r.querySelector('.var-c2');
    const c3 = r.querySelector('.var-c3');
    return {
      num: i + 1,
      cor1: c1.value, cor1Nome: c1.options[c1.selectedIndex]?.text || '',
      cor2: c2.value, cor2Nome: c2.options[c2.selectedIndex]?.text || '',
      cor3: c3 ? c3.value : '', cor3Nome: c3 ? (c3.options[c3.selectedIndex]?.text || '') : '',
      obs: r.querySelector('.var-obs').value
    };
  }).filter(v => v.cor1 || v.cor2 || v.cor3);

  // Lê grade/camadas primeiro para calcular quantidades de componentes por tamanho
  const gP = parseInt(v('f-gr-p'))||0, gM = parseInt(v('f-gr-m'))||0;
  const gG = parseInt(v('f-gr-g'))||0, gGG = parseInt(v('f-gr-gg'))||0, gG1 = parseInt(v('f-gr-g1'))||0;
  const gG2 = parseInt(v('f-gr-g2'))||0, gG3 = parseInt(v('f-gr-g3'))||0;
  const camadasN = parseInt(v('f-enf-camadas'))||0;
  // multPrincipal: 1 camada produz quantas peças por slot da grade.
  // Moletom = 1, Malha algodão (camiseta) = 2 (tubo/dobrado corta em camada dupla).
  // Sem isso, qtdPorTamanho dos componentes sai pela metade em camisetas.
  // Mesma lógica usada em calcularCamadasParaProducao e na linha "Total por tamanho" da impressão.
  const _gradeIdSel = v('f-grade-preset');
  const _gradeSel = _gradeIdSel ? STATE.grades.find(g => g.id === _gradeIdSel) : null;
  const _fasesSel = _gradeSel?.fases || [];
  const _temMoletom = _fasesSel.some(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId)) === 'moletom')
    || tecidos.some(t => categoriaEfetivaTecido(STATE.tecidos.find(x => x.id === t.tecidoId)) === 'moletom');
  const _temMalha = !_temMoletom && (
    _fasesSel.some(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId)) === 'malha')
    || tecidos.some(t => categoriaEfetivaTecido(STATE.tecidos.find(x => x.id === t.tecidoId)) === 'malha')
  );
  const multPrincipal = _temMoletom ? 1 : (_temMalha ? (MULTIPLICADOR_PECAS.malha || 2) : 1);
  const pecasPorTamanho = {
    p:  gP  * camadasN * multPrincipal,
    m:  gM  * camadasN * multPrincipal,
    g:  gG  * camadasN * multPrincipal,
    gg: gGG * camadasN * multPrincipal,
    g1: gG1 * camadasN * multPrincipal,
    g2: gG2 * camadasN * multPrincipal,
    g3: gG3 * camadasN * multPrincipal
  };

  const componentes = Array.from(document.querySelectorAll('#componentes-rows .componente-row')).map(r => {
    const nomeEl = r.querySelector('.comp-nome');
    if (!nomeEl) return null;
    const mat = r.querySelector('.comp-mat');
    const cor = r.querySelector('.comp-cor');
    const qtdEl = r.querySelector('.comp-qtd');
    const qtdPorPeca = parseFloat(qtdEl?.value) || 0;
    const qtdPorTamanho = {};
    let qtdTotal = 0;
    for (const t of ['p','m','g','gg','g1','g2','g3']) {
      const v = (pecasPorTamanho[t] || 0) * qtdPorPeca;
      qtdPorTamanho[t] = v;
      qtdTotal += v;
    }
    return {
      nome: nomeEl.value,
      material: mat.value, materialNome: mat.options[mat.selectedIndex]?.text || '',
      cor: cor?.value || '', corNome: cor?.value ? (cor.options[cor.selectedIndex]?.text || '') : '',
      qtdPorPeca, qtdPorTamanho, qtdTotal
    };
  }).filter(c => c && c.nome);

  const aviamentos = Array.from(document.querySelectorAll('#aviamentos-rows .componente-row')).map(r => {
    const mat = r.querySelector('.av-mat');
    if (!mat) return null;
    const qtdPorPeca = parseFloat(r.querySelector('.av-qtd')?.value) || 0;
    const qtdPorTamanho = {};
    let qtdTotal = 0;
    for (const t of ['p','m','g','gg','g1','g2','g3']) {
      const v = (pecasPorTamanho[t] || 0) * qtdPorPeca;
      qtdPorTamanho[t] = v;
      qtdTotal += v;
    }
    return {
      material: mat.value,
      materialNome: mat.options[mat.selectedIndex]?.text || '',
      app: r.querySelector('.av-app')?.value || '',
      qtd: qtdPorPeca,        // retrocompat: texto antigo de qtd virou número
      qtdPorPeca, qtdPorTamanho, qtdTotal
    };
  }).filter(a => a && a.material);

  const etapas = Array.from(document.querySelectorAll('#etapas-container input:checked')).map(c => c.value);

  const grade = {
    descricao: v('f-grade-desc'),
    p: parseInt(v('f-gr-p'))||0, m: parseInt(v('f-gr-m'))||0,
    g: parseInt(v('f-gr-g'))||0, gg: parseInt(v('f-gr-gg'))||0, g1: parseInt(v('f-gr-g1'))||0,
    g2: parseInt(v('f-gr-g2'))||0, g3: parseInt(v('f-gr-g3'))||0
  };
  grade.total = grade.p+grade.m+grade.g+grade.gg+grade.g1+grade.g2+grade.g3;

  const blocosEnfesto = lerEnfestoBlocos();
  const primeiroBloco = blocosEnfesto[0] || { comp: 0, larg: 0 };
  const enfesto = {
    comprimento: primeiroBloco.comp || 0,
    largura: primeiroBloco.larg || 0,
    camadas: parseInt(v('f-enf-camadas')) || 0,
    target: parseInt(v('f-enf-target')) || 0,
    blocos: blocosEnfesto
  };
  enfesto.totalPecas = grade.total * enfesto.camadas;

  // helper: retorna {id, nome} a partir de um select
  const getSelect = id => {
    const el = document.getElementById(id);
    if (!el) return { id: '', nome: '' };
    const selIdx = el.selectedIndex;
    const txt = selIdx >= 0 ? el.options[selIdx]?.text || '' : '';
    return { id: el.value, nome: txt.startsWith('—') ? '' : txt };
  };

  const griffe = getSelect('f-griffe');
  const linha = getSelect('f-linha');
  const base = getSelect('f-base');
  const bloco = getSelect('f-bloco');
  const designer = getSelect('f-designer');
  const ftec = getSelect('f-ftec');
  const coord = getSelect('f-coordenado');

  return {
    id: v('f-id') || uid(),
    os: v('f-os'),
    codigo: v('f-codigo'),
    data: v('f-data'),
    coordenadoId: coord.id,
    coordenadoNome: coord.nome,
    colecaoId: v('f-colecao'),
    colecaoNome: document.getElementById('f-colecao').options[document.getElementById('f-colecao').selectedIndex]?.text || '',
    modeloId: v('f-modelo'),
    modeloNome: document.getElementById('f-modelo').options[document.getElementById('f-modelo').selectedIndex]?.text || '',
    blocoId: bloco.id,      blocoNome: bloco.nome,
    linhaId: linha.id,      linhaNome: linha.nome,
    griffeId: griffe.id,    griffeNome: griffe.nome,
    baseId: base.id,        baseNome: base.nome,
    designerId: designer.id, designerNome: designer.nome,
    ftecId: ftec.id,        ftecNome: ftec.nome,
    desenhoId: v('f-desenho'),
    gradeId: v('f-grade-preset'),
    fases: (() => {
      const gId = v('f-grade-preset');
      if (!gId) return [];
      const gFull = STATE.grades.find(x => x.id === gId);
      if (!gFull || !Array.isArray(gFull.fases)) return [];
      return gFull.fases.map(f => ({
        ordem: f.ordem,
        nome: f.nome || '',
        tecidoId: f.tecidoId || '',
        tecidoNome: (STATE.tecidos.find(t => t.id === f.tecidoId) || {}).nome || '',
        corId: f.corId || '',
        corNome: (STATE.cores.find(c => c.id === f.corId) || {}).nome || '',
        comp: f.comp || '',
        larg: f.larg || ''
      }));
    })(),
    tecidos, grade, enfesto, etapas, variantes, componentes, aviamentos,
    obs: v('f-obs'),
    atencao: v('f-atencao'),
    criadoEm: new Date().toISOString()
  };
}

// Validação antes de salvar: limite de camadas.
// Retorna true se pode prosseguir, false se o usuário cancelou.
function validarAntesDeSalvar(data) {
  const { limite, categoriaRestritiva } = calcularLimiteCamadas();
  const camadas = data.enfesto?.camadas || 0;
  if (camadas > 0 && camadas > limite) {
    const catLabel = categoriaRestritiva === 'moletom' ? 'moletom (máx 36)' : 'malha algodão (máx 80)';
    return confirm(`⚠ Atenção: você informou ${camadas} camadas, mas o limite para ${catLabel} é ${limite}.\n\nDeseja salvar mesmo assim?`);
  }
  return true;
}

/* ========================================================= */
/*   REGRA: CAMISETA BICOLOR -> auto-gera CAMISETA BÁSICA    */
/* ========================================================= */
// Quando uma OS e salva com desenho "Camiseta Bicolor" e a grade
// "P-M-G-G1-G2-G3 (CONJUGADO COM BÁSICA) | CM.BICOLOR", gera
// automaticamente uma OS conjugada com desenho "Camiseta Básica | Branco"
// e grade "M-G (CONJUGADO COM BICOLOR) | CM.BÁSICA", reaproveitando o
// peças-alvo (target) da bicolor pra calcular as camadas da básica.

const REGRA_BICOLOR_BASICA = {
  gradeBicolorNome: 'P-M-G-G1-G2-G3 | CM.BICOLOR',
  desenhoBasicaNome: 'Camiseta Básica | Branco',
  gradeBasicaNome: 'M-G (CONJUGADO) | CM.BÁSICA'
};

function _normNome(s) {
  return (s || '').toString().normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// As cores passaram a ser cadastradas COM O TECIDO NO NOME ("Preto Malha Algodão",
// "Preto Moletom") para que a gramatura (g/m²) seja única por tecido+cor — o mesmo
// "Preto" pesa diferente em malha e em moletom. Os três helpers abaixo conciliam
// essa nomenclatura com as partes do app que continuam raciocinando na cor PURA
// (sigla do SKU, ordem de cores do desc) ou que já mostram o tecido ao lado.

// Maior nome de tecido cadastrado que casa como sufixo do nome normalizado `n`
// ('' se nenhum casar). "Maior" importa porque há tecidos aninhados: "Preto
// Ribana Malha Algodão" casa tanto "Malha Algodão" quanto "Ribana Malha
// Algodão", e o certo é o segundo.
function _sufixoTecidoNorm(n) {
  let sufixo = '';
  (STATE.tecidos || []).forEach(t => {
    const tn = _normNome(t.nome);
    if (!tn || tn.length <= sufixo.length) return;
    if (n.endsWith(' ' + tn)) sufixo = tn;
  });
  return sufixo;
}

// Nome "base" da cor, normalizado e sem o tecido: "Preto Malha Algodão" → "preto".
// Para COMPARAR (sigla do SKU, ordem do desc). Se nenhum tecido casar, devolve o
// nome normalizado inteiro ("Preto" → "preto", comportamento antigo).
function corBaseNome(nome) {
  const n = _normNome(nome);
  if (!n) return '';
  const s = _sufixoTecidoNorm(n);
  return s ? n.slice(0, n.length - s.length - 1).trim() : n;
}

// Nome da cor sem o tecido, PRESERVANDO acento e caixa do cadastro — para EXIBIR
// a cor da peça (banner impresso, etiqueta), onde o tecido não interessa e o nome
// composto estouraria a caixa. "Café Ribana Moletom" → "Café". Diferente do
// corBaseNome, que normaliza e serve para comparação, não para imprimir.
function corNomeCurto(nome) {
  const c = (nome == null ? '' : String(nome)).trim();
  if (!c) return '';
  const s = _sufixoTecidoNorm(_normNome(c));
  if (!s) return c;
  const palavras = c.split(/\s+/);
  return palavras.slice(0, Math.max(1, palavras.length - s.split(' ').length)).join(' ');
}

// Rótulo curto da cor para linhas que JÁ mostram o tecido numa coluna ao lado —
// evita "Malha Algodão · Preto Malha Algodão" e o estouro de largura na folha.
// Só corta quando o sufixo é exatamente o tecido DAQUELA linha. Corta por
// palavras (não por índice) para não depender do tamanho após tirar acentos.
function corSemTecido(corNome, tecidoNome) {
  const c = (corNome == null ? '' : String(corNome)).trim();
  const tn = _normNome(tecidoNome);
  if (!c || !tn) return c;
  const n = _normNome(c);
  // Exige que o tecido seja EXATAMENTE o sufixo resolvido pelo corBaseNome, e não
  // um endsWith solto: existem tecidos aninhados ("Ribana Malha Algodão" termina
  // em "Malha Algodão"), e um endsWith cortaria "Preto Ribana Malha Algodão" para
  // "Preto Ribana" numa linha de Malha Algodão.
  if (n === tn || n !== corBaseNome(c) + ' ' + tn) return c;
  const palavras = c.split(/\s+/);
  const corta = tn.split(' ').length;
  return palavras.slice(0, Math.max(1, palavras.length - corta)).join(' ');
}

// Canonicaliza (cor, tecido) para o nome COMPOSTO cadastrado. A Contabilidade
// ainda manda a cor pura ("Preto") nas compras por NF, enquanto as OSs baixam
// pelo nome cadastrado ("Preto Malha Algodão"). Como a chave do estoque é
// tecido||cor, sem isto a entrada e a saída caem em linhas diferentes e o saldo
// do tecido racha em duas. Se não achar cor cadastrada que case tecido+cor,
// devolve o nome recebido — nunca inventa nem descarta movimento.
function corCanonicaPorTecido(corNome, tecidoNome) {
  const cn = _normNome(corNome);
  const tn = _normNome(tecidoNome);
  if (!cn) return corNome || '';
  // Já veio no formato composto de uma cor cadastrada → nada a fazer.
  const jaComposta = (STATE.cores || []).some(
    c => _normNome(c.nome) === cn && corBaseNome(c.nome) !== cn);
  if (jaComposta || !tn) return corNome || '';
  // Casamento EXATO de "cor base + tecido": com tecidos aninhados ("Ribana Malha
  // Algodão" termina em "Malha Algodão"), um endsWith faria a compra de Malha
  // Algodão cair na cor da Ribana e baixar do saldo errado.
  const alvo = (STATE.cores || []).find(
    c => corBaseNome(c.nome) === cn && _normNome(c.nome) === cn + ' ' + tn);
  return alvo ? alvo.nome : (corNome || '');
}

function _desenhoEhCamisetaBicolor(d) {
  if (!d) return false;
  const desc = _normNome(d.desc);
  const cod = _normNome(d.codigo);
  return (desc.includes('camiseta') && desc.includes('bicolor'))
      || (cod.includes('camiseta') && cod.includes('bicolor'));
}

function deveGerarConjugadaBasica(osBicolor) {
  // Evita loop: se a propria OS ja e uma conjugada, nao gera outra
  if (osBicolor.conjugadaPaiId) return false;
  // Se ja existe a conjugada e ela ainda esta na lista, nao duplica
  if (osBicolor.conjugadaId && STATE.ordens.find(o => o.id === osBicolor.conjugadaId)) return false;
  const desenho = STATE.desenhos.find(d => d.id === osBicolor.desenhoId);
  if (!_desenhoEhCamisetaBicolor(desenho)) return false;
  const grade = STATE.grades.find(g => g.id === osBicolor.gradeId);
  if (!grade) return false;
  if (_normNome(grade.nome) !== _normNome(REGRA_BICOLOR_BASICA.gradeBicolorNome)) return false;
  return true;
}

async function gerarConjugadaBasica(osBicolor) {
  const desBasica = STATE.desenhos.find(d => _normNome(d.desc) === _normNome(REGRA_BICOLOR_BASICA.desenhoBasicaNome));
  if (!desBasica) {
    toast(`Desenho "${REGRA_BICOLOR_BASICA.desenhoBasicaNome}" não cadastrado — OS conjugada não foi gerada`, 'err');
    return null;
  }
  const grBasica = STATE.grades.find(g => _normNome(g.nome) === _normNome(REGRA_BICOLOR_BASICA.gradeBasicaNome));
  if (!grBasica) {
    toast(`Grade "${REGRA_BICOLOR_BASICA.gradeBasicaNome}" não cadastrada — OS conjugada não foi gerada`, 'err');
    return null;
  }

  const target = parseInt(osBicolor.enfesto?.target) || 0;
  const tamanhos = grBasica.tamanhos || {};
  const qtdsValidos = ['p','m','g','gg','g1','g2','g3']
    .map(k => parseInt(tamanhos[k]) || 0)
    .filter(q => q > 0);
  const minQtd = qtdsValidos.length ? Math.min(...qtdsValidos) : 0;
  const camadas = (target > 0 && minQtd > 0)
    ? Math.ceil(target / minQtd)
    : (parseInt(osBicolor.enfesto?.camadas) || 1);

  // Clona o contexto do bicolor (data, equipe, colecao, marca, etc.) e ajusta
  const novaOs = JSON.parse(JSON.stringify(osBicolor));
  novaOs.id = uid();
  novaOs.os = proximoNumeroOS();
  novaOs.codigo = desBasica.codigo || '';
  novaOs.desenhoId = desBasica.id;
  novaOs.gradeId = grBasica.id;
  novaOs.conjugadaPaiId = osBicolor.id;
  delete novaOs.conjugadaId;

  // Grade nova (a partir do cadastro da basica)
  novaOs.grade = {
    descricao: grBasica.nome,
    p: parseInt(tamanhos.p) || 0,
    m: parseInt(tamanhos.m) || 0,
    g: parseInt(tamanhos.g) || 0,
    gg: parseInt(tamanhos.gg) || 0,
    g1: parseInt(tamanhos.g1) || 0,
    g2: parseInt(tamanhos.g2) || 0,
    g3: parseInt(tamanhos.g3) || 0
  };
  novaOs.grade.total = novaOs.grade.p + novaOs.grade.m + novaOs.grade.g
                     + novaOs.grade.gg + novaOs.grade.g1 + novaOs.grade.g2 + novaOs.grade.g3;

  // Fases do enfesto a partir da grade nova
  novaOs.fases = Array.isArray(grBasica.fases) ? grBasica.fases.map(f => ({
    ordem: f.ordem,
    nome: f.nome || '',
    tecidoId: f.tecidoId || '',
    tecidoNome: (STATE.tecidos.find(t => t.id === f.tecidoId) || {}).nome || '',
    corId: f.corId || '',
    corNome: (STATE.cores.find(c => c.id === f.corId) || {}).nome || '',
    comp: f.comp || '',
    larg: f.larg || ''
  })) : [];

  novaOs.enfesto = {
    comprimento: parseFloat(novaOs.fases[0]?.comp) || 0,
    largura: parseFloat(novaOs.fases[0]?.larg) || 0,
    camadas,
    target,
    blocos: novaOs.fases.length
      ? novaOs.fases.map(f => ({ comp: parseFloat(f.comp) || 0, larg: parseFloat(f.larg) || 0 }))
      : [{ comp: 0, larg: 0 }],
    totalPecas: novaOs.grade.total * camadas
  };

  // Componentes do desenho da basica (se houver)
  const compsDes = Array.isArray(desBasica.componentes) ? desBasica.componentes : [];
  if (compsDes.length) {
    const pecasPorTamanho = {
      p: novaOs.grade.p * camadas,
      m: novaOs.grade.m * camadas,
      g: novaOs.grade.g * camadas,
      gg: novaOs.grade.gg * camadas,
      g1: novaOs.grade.g1 * camadas,
      g2: novaOs.grade.g2 * camadas,
      g3: novaOs.grade.g3 * camadas
    };
    novaOs.componentes = compsDes.map(c => {
      const cad = STATE.componentes.find(x => x.id === c.componenteId);
      const qtdPorPeca = c.qtdPorPeca != null ? c.qtdPorPeca : 1;
      const qtdPorTamanho = {};
      let qtdTotal = 0;
      for (const t of ['p','m','g','gg','g1','g2','g3']) {
        const v = (pecasPorTamanho[t] || 0) * qtdPorPeca;
        qtdPorTamanho[t] = v;
        qtdTotal += v;
      }
      return {
        nome: c.nome || cad?.nome || '',
        material: c.tecidoId ? 'T:' + c.tecidoId : '',
        materialNome: (STATE.tecidos.find(t => t.id === c.tecidoId) || {}).nome || '',
        cor: c.corId || '',
        corNome: (STATE.cores.find(co => co.id === c.corId) || {}).nome || '',
        qtdPorPeca, qtdPorTamanho, qtdTotal
      };
    });
  }

  // Aviamentos do desenho da basica (se houver)
  const avsDes = Array.isArray(desBasica.aviamentos) ? desBasica.aviamentos : [];
  if (avsDes.length) {
    const pecasTot = novaOs.grade.total * camadas;
    novaOs.aviamentos = avsDes.map(av => {
      const qtdPorPeca = parseFloat(av.qtdPorPeca) || 1;
      return {
        material: av.materialId,
        materialNome: (STATE.materiais.find(m => m.id === av.materialId) || {}).desc || '',
        app: av.aplicacao || '',
        qtd: qtdPorPeca,
        qtdPorPeca,
        qtdPorTamanho: {},
        qtdTotal: pecasTot * qtdPorPeca
      };
    });
  }

  // Marca o vinculo na bicolor (sera persistido no proximo saveState)
  const idxBicolor = STATE.ordens.findIndex(o => o.id === osBicolor.id);
  if (idxBicolor >= 0) {
    STATE.ordens[idxBicolor].conjugadaId = novaOs.id;
    osBicolor.conjugadaId = novaOs.id;
  }

  STATE.ordens.push(novaOs);
  await saveState('ordens');
  await atualizarCounterOS(novaOs.os);
  return novaOs;
}

async function aplicarRegraConjugadaSeAplicavel(osBicolor) {
  if (!deveGerarConjugadaBasica(osBicolor)) return null;
  const conjugada = await gerarConjugadaBasica(osBicolor);
  if (conjugada) {
    toast(`OS conjugada gerada: OS ${conjugada.os} (Camiseta Básica)`, 'ok');
  }
  return conjugada;
}

async function salvarOS() {
  const data = coletaOS();
  if (!data.os && !data.codigo) {
    return toast('Preencha ao menos número da OS ou código do desenho', 'err');
  }
  if (!validarAntesDeSalvar(data)) return;
  const idx = STATE.ordens.findIndex(o => o.id === data.id);
  if (idx >= 0) STATE.ordens[idx] = data; else STATE.ordens.push(data);
  await saveState('ordens');
  await atualizarCounterOS(data.os);
  osEditId = null;
  await aplicarBaixaEstoqueOS(data);
  await aplicarRegraConjugadaSeAplicavel(data);
  toast('OS ' + data.os + ' salva', 'ok');
  // Mantem etiquetas/etiqueta-<numero>.pdf em sincronia com a grade/qtde atual.
  // Sem silent: o toast 'PDF etiquetas salvo: ...' confirma a regravacao em disco.
  salvarPdfEtiquetasAuto(data, dadosEtiquetaParaOS(data));
  goto('lista-os');
}

/* ========================================================= */
/*           PASTA DE PDFs (File System Access API)          */
/* ========================================================= */
// Salva o DirectoryHandle no IndexedDB para persistir entre sessoes
// (handles nao sao serializaveis pra localStorage). O Chrome/Edge
// preserva a permissao concedida; se o usuario revogar, queryPermission
// volta a 'prompt' e pedimos de novo via requestPermission.
const PDF_DB_NAME = 'gerador-os-pdf';
const PDF_DB_STORE = 'handles';
const PDF_DB_KEY = 'output-folder';
let pdfFolderHandle = null;

function _openPdfDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(PDF_DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(PDF_DB_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function savePdfFolderHandle(handle) {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).put(handle, PDF_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function loadPdfFolderHandle() {
  try {
    const db = await _openPdfDb();
    const handle = await new Promise((res, rej) => {
      const tx = db.transaction(PDF_DB_STORE, 'readonly');
      const req = tx.objectStore(PDF_DB_STORE).get(PDF_DB_KEY);
      req.onsuccess = () => res(req.result || null);
      req.onerror = () => rej(req.error);
    });
    db.close();
    return handle;
  } catch (e) {
    console.warn('loadPdfFolderHandle', e);
    return null;
  }
}

async function clearPdfFolderHandle() {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).delete(PDF_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function ensureFolderPermission(handle, mode = 'readwrite') {
  if (!handle || typeof handle.queryPermission !== 'function') return false;
  const opts = { mode };
  if ((await handle.queryPermission(opts)) === 'granted') return true;
  if ((await handle.requestPermission(opts)) === 'granted') return true;
  return false;
}

async function pickPdfFolder() {
  if (!('showDirectoryPicker' in window)) {
    toast('Navegador não suporta seleção de pasta. Use Chrome ou Edge no desktop.', 'err');
    return null;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await savePdfFolderHandle(handle);
    pdfFolderHandle = handle;
    return handle;
  } catch (e) {
    if (e.name === 'AbortError') return null;
    console.error('pickPdfFolder', e);
    toast('Falha ao selecionar pasta: ' + (e.message || e), 'err');
    return null;
  }
}

async function conectarPastaPdf() {
  const handle = await pickPdfFolder();
  if (handle) {
    toast(`Pasta conectada: ${handle.name}`, 'ok');
    atualizarPdfFolderStatus();
  }
}

async function desconectarPastaPdf() {
  await clearPdfFolderHandle();
  pdfFolderHandle = null;
  toast('Pasta desconectada', '');
  atualizarPdfFolderStatus();
}

/* ----- Pasta de backup automatico (JSON) ----- */
// Mesma abordagem da pasta de PDF (File System Access + IndexedDB).
// Reusa o mesmo DB/store, com chave diferente.
const BACKUP_DB_KEY = 'backup-folder';
let backupFolderHandle = null;

async function saveBackupFolderHandle(handle) {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).put(handle, BACKUP_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function loadBackupFolderHandle() {
  try {
    const db = await _openPdfDb();
    const handle = await new Promise((res, rej) => {
      const tx = db.transaction(PDF_DB_STORE, 'readonly');
      const req = tx.objectStore(PDF_DB_STORE).get(BACKUP_DB_KEY);
      req.onsuccess = () => res(req.result || null);
      req.onerror = () => rej(req.error);
    });
    db.close();
    return handle;
  } catch (e) {
    console.warn('loadBackupFolderHandle', e);
    return null;
  }
}

async function clearBackupFolderHandle() {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).delete(BACKUP_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

/* ========================================================= */
/*        SNAPSHOTS DE CONTINGÊNCIA (LOCAL + PASTA)          */
/* ========================================================= */
// A cada alteração persistida, guardamos uma cópia do blob inteiro:
//  - LOCAL: IndexedDB próprio (ring dos últimos N), sobrevive a apagamento
//    do servidor e não depende de rede/pasta;
//  - PASTA: arquivo versionado snapshots/snap-<ts>.json na pasta conectada
//    (sincroniza pro Drive), ring de M arquivos.
// Objetivo: qualquer perda vira rollback de 1 clique em Configurações.
const SNAP_DB_NAME = 'gerador-os-snapshots';
const SNAP_DB_STORE = 'snaps';
const SNAP_MAX_LOCAL = 30;          // quantos snapshots locais manter
const SNAP_MAX_PASTA = 15;          // quantos arquivos na pasta manter
const SNAP_MIN_INTERVALO_MS = 20000; // no máximo 1 snapshot a cada 20s
let _ultimoSnapTs = 0;
// Marca se o app já viu dados de verdade nesta sessão. Serve à trava
// anti-apagamento: se já tivemos dados, um flush "vazio" é bloqueado.
let _appJaTeveDados = false;
let _permitirFlushVazio = false; // liberado só em ações intencionais (limpar/restaurar)

function _openSnapDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(SNAP_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const s = req.result.createObjectStore(SNAP_DB_STORE, { keyPath: 'id', autoIncrement: true });
      s.createIndex('ts', 'ts');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// Conta itens de uma chave do blob (que no cloudCache é string JSON).
function _contarItens(cache, k) {
  try {
    const v = cache && cache[k];
    const a = typeof v === 'string' ? JSON.parse(v) : v;
    return Array.isArray(a) ? a.length : 0;
  } catch (e) { return 0; }
}

// Um blob "vazio" = sem nenhuma OS E sem nenhum desenho. É o formato
// de um apagamento acidental (cloudCache zerado sendo gravado por cima).
function _blobEstaVazio(cache) {
  return _contarItens(cache, 'ordens') === 0 && _contarItens(cache, 'desenhos') === 0;
}

async function salvarSnapshotContingencia({ forcar = false } = {}) {
  try {
    if (!cloudCache || _blobEstaVazio(cloudCache)) return; // nunca snapshota lixo
    const agora = Date.now();
    if (!forcar && (agora - _ultimoSnapTs) < SNAP_MIN_INTERVALO_MS) return;
    _ultimoSnapTs = agora;
    const registro = {
      ts: agora,
      iso: new Date(agora).toISOString(),
      by: (currentUser && currentUser.email) || null,
      resumo: { ordens: _contarItens(cloudCache, 'ordens'), desenhos: _contarItens(cloudCache, 'desenhos') },
      data: JSON.parse(JSON.stringify(cloudCache))
    };
    // 1) LOCAL (IndexedDB) + poda
    try {
      const db = await _openSnapDb();
      await new Promise((res, rej) => {
        const tx = db.transaction(SNAP_DB_STORE, 'readwrite');
        tx.objectStore(SNAP_DB_STORE).add(registro);
        tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
      });
      // poda: mantém os SNAP_MAX_LOCAL mais recentes
      const ids = await new Promise((res, rej) => {
        const tx = db.transaction(SNAP_DB_STORE, 'readonly');
        const req = tx.objectStore(SNAP_DB_STORE).getAllKeys();
        req.onsuccess = () => res(req.result || []); req.onerror = () => rej(req.error);
      });
      if (ids.length > SNAP_MAX_LOCAL) {
        const excluir = ids.slice(0, ids.length - SNAP_MAX_LOCAL);
        await new Promise((res, rej) => {
          const tx = db.transaction(SNAP_DB_STORE, 'readwrite');
          const st = tx.objectStore(SNAP_DB_STORE);
          excluir.forEach(id => st.delete(id));
          tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
        });
      }
      db.close();
    } catch (e) { console.warn('snapshot local', e); }
    // 2) PASTA (arquivo versionado) + poda — best-effort
    escreverSnapshotNaPasta(registro).catch(e => console.warn('snapshot pasta', e));
  } catch (e) {
    console.warn('salvarSnapshotContingencia', e);
  }
}

async function escreverSnapshotNaPasta(registro) {
  const raiz = backupFolderHandle || (await loadBackupFolderHandle()) || pdfFolderHandle || (await loadPdfFolderHandle());
  if (!raiz) return;
  if (!(await ensureFolderPermission(raiz, 'readwrite'))) return;
  const dir = await raiz.getDirectoryHandle('snapshots', { create: true });
  const nome = 'snap-' + registro.iso.replace(/[:.]/g, '-') + '.json';
  const fh = await dir.getFileHandle(nome, { create: true });
  const w = await fh.createWritable();
  await w.write(JSON.stringify({ __meta: { iso: registro.iso, by: registro.by, resumo: registro.resumo }, ...registro.data }));
  await w.close();
  // poda: mantém os SNAP_MAX_PASTA arquivos mais recentes
  try {
    const nomes = [];
    for await (const [n, h] of dir.entries()) {
      if (h.kind === 'file' && /^snap-.*\.json$/.test(n)) nomes.push(n);
    }
    nomes.sort();
    for (const n of nomes.slice(0, Math.max(0, nomes.length - SNAP_MAX_PASTA))) {
      try { await dir.removeEntry(n); } catch (e) { /* ok */ }
    }
  } catch (e) { /* diretório sem iterador — ignora poda */ }
}

async function listarSnapshotsLocais() {
  const cont = document.getElementById('snapshotsLocaisList');
  if (!cont) return;
  cont.innerHTML = '<div class="empty" style="padding:20px;">Carregando...</div>';
  let regs = [];
  try {
    const db = await _openSnapDb();
    regs = await new Promise((res, rej) => {
      const tx = db.transaction(SNAP_DB_STORE, 'readonly');
      const req = tx.objectStore(SNAP_DB_STORE).getAll();
      req.onsuccess = () => res(req.result || []); req.onerror = () => rej(req.error);
    });
    db.close();
  } catch (e) {
    cont.innerHTML = `<div class="empty" style="padding:20px;">Erro ao ler snapshots: ${esc(e.message || e)}</div>`;
    return;
  }
  if (!regs.length) { cont.innerHTML = '<div class="empty" style="padding:20px;">Nenhum snapshot ainda — é criado automaticamente a cada alteração.</div>'; return; }
  regs.sort((a, b) => b.ts - a.ts);
  cont.innerHTML = `<table class="table">
    <thead><tr><th>Quando</th><th>Conteúdo</th><th class="col-actions">Ação</th></tr></thead>
    <tbody>${regs.map(r => `
      <tr>
        <td>${esc(new Date(r.ts).toLocaleString('pt-BR'))}</td>
        <td>${r.resumo ? `${r.resumo.ordens} OS · ${r.resumo.desenhos} desenhos` : '—'}</td>
        <td class="col-actions"><button class="btn small danger" onclick="restaurarSnapshotLocal(${r.id})">Restaurar</button></td>
      </tr>`).join('')}
    </tbody></table>`;
}

async function restaurarSnapshotLocal(id) {
  if (!exigirAdmin('restaurar snapshots')) return;
  let reg = null;
  try {
    const db = await _openSnapDb();
    reg = await new Promise((res, rej) => {
      const tx = db.transaction(SNAP_DB_STORE, 'readonly');
      const req = tx.objectStore(SNAP_DB_STORE).get(id);
      req.onsuccess = () => res(req.result || null); req.onerror = () => rej(req.error);
    });
    db.close();
  } catch (e) { toast('Erro ao ler snapshot', 'err'); return; }
  if (!reg || !reg.data) { toast('Snapshot não encontrado', 'err'); return; }
  const quando = new Date(reg.ts).toLocaleString('pt-BR');
  const conf = prompt(
    `Restaurar o snapshot de ${quando} (${reg.resumo ? reg.resumo.ordens + ' OS' : ''})?\n\n` +
    `Isso vai SOBRESCREVER os dados atuais (de todos) com essa versão.\n\n` +
    `Para confirmar, digite RESTAURAR:`
  );
  if (conf === null) return;
  if ((conf || '').trim().toUpperCase() !== 'RESTAURAR') { toast('Palavra não conferiu — nada foi restaurado.', 'err'); return; }
  cloudCache = JSON.parse(JSON.stringify(reg.data));
  if (supa && currentUser) {
    setSyncStatus('saving');
    try {
      const { error } = await supa.from('shared_data').upsert({
        id: 'main', data: cloudCache, updated_at: new Date().toISOString(), updated_by: currentUser.id
      }, { onConflict: 'id' });
      if (error) throw error;
      setSyncStatus('ok');
    } catch (e) { setSyncStatus('error'); toast('Erro ao gravar no servidor: ' + (e.message || e), 'err'); return; }
  }
  await loadState();
  goto('home');
  toast(`Snapshot de ${quando} restaurado`, 'ok');
}

async function pickBackupFolder() {
  if (!('showDirectoryPicker' in window)) {
    toast('Navegador não suporta seleção de pasta. Use Chrome ou Edge no desktop.', 'err');
    return null;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await saveBackupFolderHandle(handle);
    backupFolderHandle = handle;
    return handle;
  } catch (e) {
    if (e.name === 'AbortError') return null;
    console.error('pickBackupFolder', e);
    toast('Falha ao selecionar pasta: ' + (e.message || e), 'err');
    return null;
  }
}

async function conectarPastaBackup() {
  const handle = await pickBackupFolder();
  if (handle) {
    toast(`Pasta de backup conectada: ${handle.name}`, 'ok');
    atualizarBackupFolderStatus();
    // Faz um backup imediato com o estado atual
    const ok = await escreverBackupJson();
    if (ok) toast('Backup inicial gravado', 'ok');
  }
}

async function desconectarPastaBackup() {
  await clearBackupFolderHandle();
  backupFolderHandle = null;
  toast('Pasta de backup desconectada', '');
  atualizarBackupFolderStatus();
}

async function escreverBackupJsonAgora() {
  const ok = await escreverBackupJson();
  if (ok) toast('Backup JSON salvo na pasta', 'ok');
  else toast('Falha ao salvar backup. Conecte a pasta primeiro.', 'err');
}

async function escreverBackupJson() {
  const handle = backupFolderHandle || (await loadBackupFolderHandle());
  if (!handle) return false;
  const ok = await ensureFolderPermission(handle, 'readwrite');
  if (!ok) return false;
  backupFolderHandle = handle;
  try {
    const dados = cloudCache || {};
    const payload = {
      __meta: {
        gerado_em: new Date().toISOString(),
        gerado_por: (currentUser && currentUser.email) || null,
        formato: 'gerador-os-snapshot-v1'
      },
      ...dados
    };
    const json = JSON.stringify(payload, null, 2);
    const fileHandle = await handle.getFileHandle('gerador-os-dados.json', { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(json);
    await writable.close();
    return true;
  } catch (e) {
    console.warn('escreverBackupJson', e);
    return false;
  }
}

async function atualizarBackupFolderStatus() {
  const el = document.getElementById('backupFolderStatus');
  if (!el) return;
  if (!('showDirectoryPicker' in window)) {
    el.innerHTML = '<span style="color: var(--alert);">Este navegador não suporta a API de pasta. Use Chrome ou Edge no desktop.</span>';
    return;
  }
  const handle = backupFolderHandle || (await loadBackupFolderHandle());
  if (!handle) {
    el.innerHTML = '<span style="color: var(--ink-3);">Nenhuma pasta conectada. O backup automático não está ativo.</span>';
    return;
  }
  backupFolderHandle = handle;
  let permLabel = 'pronta — backup gravado a cada mudança';
  try {
    const perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') permLabel = 'precisa renovar permissão (clique em "Conectar pasta")';
  } catch (_) {}
  el.innerHTML = `<strong>Conectada:</strong> <code>${esc(handle.name)}</code> — ${permLabel}`;
}

async function atualizarPdfFolderStatus() {
  const el = document.getElementById('pdfFolderStatus');
  if (!el) return;
  if (!('showDirectoryPicker' in window)) {
    el.innerHTML = '<span style="color: var(--alert);">Este navegador não suporta a API de pasta. Use Chrome ou Edge no desktop.</span>';
    return;
  }
  const handle = pdfFolderHandle || (await loadPdfFolderHandle());
  if (!handle) {
    el.innerHTML = '<span style="color: var(--ink-3);">Nenhuma pasta conectada. Os PDFs não serão salvos automaticamente até você conectar uma pasta.</span>';
    return;
  }
  pdfFolderHandle = handle;
  let permLabel = 'pronta';
  try {
    const perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') permLabel = 'precisa renovar permissão (clique em "Conectar pasta")';
  } catch (_) {}
  el.innerHTML = `<strong>Conectada:</strong> <code>${esc(handle.name)}</code> — ${permLabel}`;
}

function sanitizeForFilename(s) {
  return String(s || '').replace(/[\\/:*?"<>|\x00-\x1F]/g, '').replace(/\s+/g, ' ').trim();
}

function pdfFilenameForOS(o) {
  const numero = sanitizeForFilename(o.os) || 'sem-numero';
  return `OS-${numero}.pdf`;
}

function etiquetaFilenameForOS(o) {
  const numero = sanitizeForFilename(o.os) || 'sem-numero';
  return `etiqueta-${numero}.pdf`;
}

// Gera PDF das etiquetas direto com jsPDF (sem html2canvas, pois o conteudo
// e so texto). Cada etiqueta vira uma pagina de 100mm x 50mm. `dados`
// precisa ter { marca, os, qtde, tam, cor, modelo, numEtiquetas }.
//
// Todos os textos (marca + 6 linhas) sao desenhados no mesmo tamanho, e
// esse tamanho e maximizado automaticamente pra ocupar a area da etiqueta
// sem estourar a borda — medindo a largura real via pdf.getTextWidth.
function gerarPdfEtiquetas(dados) {
  const _jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
  if (typeof _jsPDF !== 'function') throw new Error('jsPDF não carregada');

  const pdf = new _jsPDF({
    unit: 'mm',
    format: [100, 50],
    orientation: 'landscape',
    compress: true
  });

  // Geometria:
  //   Borda: rect(2, 2, 96, 46) -> de (2,2) a (98,48)
  //   Area util de texto: x 3.5..96.5 (93mm), y 3..47 (44mm)
  const xLeft = 3.5;
  const xCenter = 50;
  const innerWidth = 93;
  const boxTop = 2;
  const boxHeight = 46;
  const verticalPad = 1.5;
  const innerHeight = boxHeight - 2 * verticalPad; // 43mm
  const lineFactor = 1.18;   // espacamento entre linhas (relativo ao fontSize)
  const ptToMm = 0.3527778;

  // Trunca com '…' ate caber em maxWidth (mm), na fonte/tamanho correntes.
  const fitText = (s, maxWidth) => {
    const str = String(s == null ? '' : s);
    if (pdf.getTextWidth(str) <= maxWidth) return str;
    let cut = str;
    while (cut.length > 0 && pdf.getTextWidth(cut + '…') > maxWidth) {
      cut = cut.slice(0, -1);
    }
    return cut + '…';
  };

  pdf.setFont('helvetica', 'bold');

  const total = Math.max(1, dados.numEtiquetas);
  const tams = dados.tamanhosPacotes || [];
  for (let i = 0; i < total; i++) {
    if (i > 0) pdf.addPage([100, 50], 'landscape');

    // Borda
    pdf.setLineWidth(0.3);
    pdf.rect(2, 2, 96, 46);

    // Cada linha tem uma escala (s): 1 = normal, 2 = dobro (o tamanho / o
    // conteúdo do pacote saem em destaque). c = centralizada.
    const ehReposicao = dados.temReposicao && i === total - 1;
    const destaque = ehReposicao
      ? { t: ETIQUETA_CONTEUDO_REPOSICAO, s: 1.6, c: true }   // conteúdo (texto longo)
      : { t: tams[i] || dados.tam, s: 2, c: true };           // SÓ o tamanho do pacote, dobro
    const linhas = [
      { t: String(dados.marca || ''), s: 1, c: true },
      { t: `OS: ${dados.os}`, s: 1 },
      { t: `MODELO: ${dados.modelo}`, s: 1 },
      { t: `QTDE: ${dados.qtde}`, s: 1 },
      { t: `TAM: ${dados.tam}`, s: 1 },                       // TODOS os tamanhos da grade, normal
      { t: `COR: ${dados.cor}`, s: 1 },
      { t: `LOTE: ${i + 1}/${total}`, s: 1 },
      destaque
    ];
    // Moletom: composição do pacote (só nas etiquetas de tamanho, não na reposição).
    if (!ehReposicao && dados.composicao) {
      dados.composicao.forEach(c => linhas.push({ t: c, s: 0.7, c: true }));
    }

    // Mede a 10pt e escala linearmente pra achar o maior fontSize base que cabe
    // em largura (cada linha ocupa largura × sua escala) e altura (soma das
    // escalas × altura de linha).
    pdf.setFontSize(10);
    const maxWAt10 = Math.max(...linhas.map(L => pdf.getTextWidth(L.t) * L.s), 0.1);
    const sumEscala = linhas.reduce((a, L) => a + L.s, 0);
    const sizeByWidth  = (10 * innerWidth) / maxWAt10;
    const sizeByHeight = innerHeight / (sumEscala * ptToMm * lineFactor);
    const fontSize = Math.min(sizeByWidth, sizeByHeight, 22);
    const lh = fontSize * ptToMm * lineFactor; // mm (altura de 1 linha na escala 1)

    let y = boxTop + (boxHeight - sumEscala * lh) / 2; // centraliza vertical
    linhas.forEach((L, idx) => {
      pdf.setFontSize(fontSize * L.s);
      const x = L.c ? xCenter : xLeft;
      pdf.text(fitText(L.t, innerWidth), x, y, { align: L.c ? 'center' : 'left', baseline: 'top' });
      y += L.s * lh;
      // Separador fino logo abaixo da MARCA (1a linha).
      if (idx === 0) {
        pdf.setLineWidth(0.18);
        pdf.line(4, y - lh * 0.12, 96, y - lh * 0.12);
      }
    });
  }

  return pdf.output('blob');
}

// Salva o PDF de etiquetas na subpasta "etiquetas" dentro da pasta de PDFs
// configurada. Silencioso quando nao ha pasta conectada — nao bloqueia o
// fluxo de impressao. Retorna true/false.
async function salvarPdfEtiquetasAuto(o, dados, { silent = false } = {}) {
  let handle = pdfFolderHandle || (await loadPdfFolderHandle());
  if (!handle) return false;
  const ok = await ensureFolderPermission(handle, 'readwrite');
  if (!ok) return false;
  pdfFolderHandle = handle;
  try {
    const blob = gerarPdfEtiquetas(dados);
    const filename = etiquetaFilenameForOS(o);
    const subfolder = await handle.getDirectoryHandle('etiquetas', { create: true });
    const fileHandle = await subfolder.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    if (!silent) toast(`PDF etiquetas salvo: etiquetas/${filename}`, 'ok');
    return true;
  } catch (e) {
    console.warn('salvarPdfEtiquetasAuto', e);
    return false;
  }
}

async function gerarPdfDaSheet() {
  // Usa html2canvas + jsPDF direto (sem o wrapper html2pdf, que em algumas
  // versoes dispara um download alem de retornar o blob, causando o
  // dialogo "Salvar como" do Windows).
  const _html2canvas = window.html2canvas;
  const _jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
  if (typeof _html2canvas !== 'function') throw new Error('html2canvas não carregada');
  if (typeof _jsPDF !== 'function') throw new Error('jsPDF não carregada');
  const sheet = document.getElementById('print-sheet');
  if (!sheet) throw new Error('Print sheet não encontrada');
  const prevZoom = sheet.style.zoom;
  const prevTransform = sheet.style.transform;
  const prevOrigin = sheet.style.transformOrigin;
  sheet.style.zoom = '';
  sheet.style.transform = '';
  sheet.style.transformOrigin = '';
  // Anula a ampliacao de leitura da tela (.sheet-scaler) durante a foto.
  // Precisa ser via classe: o zoom vem do styles.css, entao limpar o
  // style inline acima nao alcanca ele. Sem isso o html2canvas 1.4.1 —
  // que nao implementa CSS zoom, mas mede o elemento ja ampliado —
  // fotografa a folha em escala errada, e o PDF sai diferente da tela.
  document.body.classList.add('pdf-capture');
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  try {
    const canvas = await _html2canvas(sheet, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });
    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    const pdf = new _jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
    const pageW = 210, pageH = 297;
    const ratio = canvas.height / canvas.width;
    let imgW = pageW;
    let imgH = pageW * ratio;
    // Se altura excede A4, reduz pra caber centralizada na pagina
    if (imgH > pageH) {
      imgH = pageH;
      imgW = pageH / ratio;
    }
    const x = (pageW - imgW) / 2;
    pdf.addImage(imgData, 'JPEG', x, 0, imgW, imgH, undefined, 'FAST');
    return pdf.output('blob');
  } finally {
    document.body.classList.remove('pdf-capture');
    sheet.style.zoom = prevZoom;
    sheet.style.transform = prevTransform;
    sheet.style.transformOrigin = prevOrigin;
  }
}

async function savePdfToFolder(blob, filename) {
  let handle = pdfFolderHandle || (await loadPdfFolderHandle());
  if (!handle) {
    toast('Conectando pasta pra salvar PDFs...', '');
    handle = await pickPdfFolder();
    if (!handle) return false;
  }
  const ok = await ensureFolderPermission(handle, 'readwrite');
  if (!ok) {
    toast('Permissão da pasta negada', 'err');
    return false;
  }
  pdfFolderHandle = handle;
  try {
    const fileHandle = await handle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return true;
  } catch (e) {
    console.error('savePdfToFolder', e);
    toast('Falha ao salvar PDF: ' + (e.message || e), 'err');
    return false;
  }
}

/* ========================================================= */
/*     PASTA E PDF DAS ORDENS DE EXPEDIÇÃO (OE)              */
/* ========================================================= */
// Mesma abordagem da pasta de PDF das OS (File System Access + IndexedDB),
// porém com pasta de destino PRÓPRIA — as OE (folha do plano de expedição)
// são salvas separadas das OS. Reusa o mesmo DB/store, chave diferente.
const OE_DB_KEY = 'oe-folder';
let oeFolderHandle = null;
let _oeSalvando = false;

async function saveOeFolderHandle(handle) {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).put(handle, OE_DB_KEY);
    tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
  });
  db.close();
}
async function loadOeFolderHandle() {
  try {
    const db = await _openPdfDb();
    const handle = await new Promise((res, rej) => {
      const tx = db.transaction(PDF_DB_STORE, 'readonly');
      const req = tx.objectStore(PDF_DB_STORE).get(OE_DB_KEY);
      req.onsuccess = () => res(req.result || null); req.onerror = () => rej(req.error);
    });
    db.close();
    return handle;
  } catch (e) { console.warn('loadOeFolderHandle', e); return null; }
}
async function clearOeFolderHandle() {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).delete(OE_DB_KEY);
    tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function pickOeFolder() {
  if (!('showDirectoryPicker' in window)) {
    toast('Navegador não suporta seleção de pasta. Use Chrome ou Edge no desktop.', 'err');
    return null;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await saveOeFolderHandle(handle);
    oeFolderHandle = handle;
    return handle;
  } catch (e) {
    if (e.name === 'AbortError') return null;
    console.error('pickOeFolder', e);
    toast('Falha ao selecionar pasta: ' + (e.message || e), 'err');
    return null;
  }
}
async function conectarPastaOe() {
  const handle = await pickOeFolder();
  if (handle) {
    toast(`Pasta das OE conectada: ${handle.name}`, 'ok');
    atualizarOeFolderStatus();
  }
}
async function desconectarPastaOe() {
  await clearOeFolderHandle();
  oeFolderHandle = null;
  toast('Pasta das OE desconectada', '');
  atualizarOeFolderStatus();
}
async function atualizarOeFolderStatus() {
  const el = document.getElementById('oeFolderStatus');
  if (!el) return;
  if (!('showDirectoryPicker' in window)) {
    el.innerHTML = '<span style="color: var(--alert);">Este navegador não suporta a API de pasta. Use Chrome ou Edge no desktop.</span>';
    return;
  }
  const handle = oeFolderHandle || (await loadOeFolderHandle());
  if (!handle) {
    el.innerHTML = '<span style="color: var(--ink-3);">Nenhuma pasta conectada. As OE não serão salvas automaticamente até você conectar uma pasta.</span>';
    return;
  }
  oeFolderHandle = handle;
  let permLabel = 'pronta — a folha do plano é salva ao abrir/gerar';
  try {
    const perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') permLabel = 'precisa renovar permissão (clique em "Conectar pasta")';
  } catch (_) {}
  el.innerHTML = `<strong>Conectada:</strong> <code>${esc(handle.name)}</code> — ${permLabel}`;
}

// Nome do arquivo da OE = período coberto pelo plano (estável: regerar o
// mesmo período reescreve o mesmo arquivo, igual às OS pelo número).
function oeFilenameForPlano() {
  const { ini, fim } = _expRange(expPlanoModo, expPlanoAncora);
  const base = (ini === fim) ? `OE-${ini}` : `OE-${ini}_a_${fim}`;
  return sanitizeForFilename(base) + '.pdf';
}

// Gera PDF multi-página da folha do plano de expedição (#print-sheet-exp),
// que pode ocupar várias A4. Fatiamos o canvas alto em páginas A4.
async function gerarPdfDaSheetExp() {
  const _html2canvas = window.html2canvas;
  const _jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
  if (typeof _html2canvas !== 'function') throw new Error('html2canvas não carregada');
  if (typeof _jsPDF !== 'function') throw new Error('jsPDF não carregada');
  const sheet = document.getElementById('print-sheet-exp');
  if (!sheet) throw new Error('Folha do plano não encontrada');
  document.body.classList.add('pdf-capture');
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  try {
    const canvas = await _html2canvas(sheet, { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false });
    // Pontos de corte "seguros" = fim de cada bloco de expedição (em px do
    // canvas). Assim a virada de página nunca parte um bloco no meio.
    const sheetRect = sheet.getBoundingClientRect();
    const ratio = canvas.width / sheetRect.width; // CSS px -> canvas px
    const cortesSeguros = Array.from(sheet.querySelectorAll('.exp-print-bloco'))
      .map(b => (b.getBoundingClientRect().bottom - sheetRect.top) * ratio)
      .filter(v => v > 0 && v <= canvas.height)
      .sort((a, b) => a - b);
    const pdf = new _jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
    const pageWmm = 210, pageHmm = 297;
    const pxPorMm = canvas.width / pageWmm;      // px do canvas por mm de largura A4
    const pageHpx = Math.floor(pageHmm * pxPorMm); // px que cabem numa A4 de altura
    let y = 0, pagina = 0;
    while (y < canvas.height - 1) {
      const maxY = y + pageHpx;
      let cut;
      if (maxY >= canvas.height) {
        cut = canvas.height;
      } else {
        // maior fim-de-bloco que cabe inteiro nesta página
        const cand = cortesSeguros.filter(v => v > y + 1 && v <= maxY);
        cut = cand.length ? Math.max(...cand) : maxY; // bloco > 1 página: corte duro
      }
      const sliceH = Math.max(1, Math.round(cut - y));
      const tmp = document.createElement('canvas');
      tmp.width = canvas.width; tmp.height = sliceH;
      const ctx = tmp.getContext('2d');
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, tmp.width, tmp.height);
      ctx.drawImage(canvas, 0, y, canvas.width, sliceH, 0, 0, canvas.width, sliceH);
      const imgData = tmp.toDataURL('image/jpeg', 0.95);
      if (pagina > 0) pdf.addPage();
      pdf.addImage(imgData, 'JPEG', 0, 0, pageWmm, sliceH / pxPorMm, undefined, 'FAST');
      y += sliceH; pagina++;
    }
    return pdf.output('blob');
  } finally {
    document.body.classList.remove('pdf-capture');
  }
}

// Salva a folha do plano atual como PDF na pasta das OE. Idempotente por
// período (reescreve o arquivo do mesmo período). silent = sem toasts de
// progresso (usado no auto-save ao abrir a folha).
async function salvarPdfOeNaPasta({ silent = false } = {}) {
  if (_oeSalvando) return false;
  let handle = oeFolderHandle || (await loadOeFolderHandle());
  if (!handle) {
    if (silent) return false;
    toast('Conecte a pasta das OE em Configurações primeiro.', 'err');
    return false;
  }
  if (!(await ensureFolderPermission(handle, 'readwrite'))) {
    if (!silent) toast('Permissão da pasta das OE negada', 'err');
    return false;
  }
  oeFolderHandle = handle;
  _oeSalvando = true;
  try {
    renderPrintPlanoExpedicao();
    await new Promise(r => setTimeout(r, 150));
    if (!silent) toast('Gerando PDF da OE...', '');
    const blob = await gerarPdfDaSheetExp();
    const filename = oeFilenameForPlano();
    const fh = await handle.getFileHandle(filename, { create: true });
    const w = await fh.createWritable();
    await w.write(blob);
    await w.close();
    if (!silent) toast(`OE salva: ${filename}`, 'ok');
    return true;
  } catch (e) {
    console.error('salvarPdfOeNaPasta', e);
    if (!silent) toast('Falha ao salvar OE: ' + (e.message || e), 'err');
    return false;
  } finally {
    _oeSalvando = false;
  }
}

async function salvarEImprimir() {
  const data = coletaOS();
  if (!data.os && !data.codigo) {
    return toast('Preencha ao menos número da OS ou código do desenho', 'err');
  }
  if (!validarAntesDeSalvar(data)) return;
  const idx = STATE.ordens.findIndex(o => o.id === data.id);
  if (idx >= 0) STATE.ordens[idx] = data; else STATE.ordens.push(data);
  await saveState('ordens');
  await atualizarCounterOS(data.os);
  osEditId = null;
  await aplicarBaixaEstoqueOS(data);
  // Aplica regra de conjugada (camiseta bicolor -> camiseta basica)
  const conjugada = await aplicarRegraConjugadaSeAplicavel(data);
  // Renderiza e navega pra print page pra que o .sheet tenha layout
  // computado (html2canvas precisa do elemento visivel com dimensoes).
  // Apos salvar o PDF, vai pra lista — sem dialogo de impressao.
  renderPrintSheet(data);
  goto('print');
  await new Promise(r => setTimeout(r, 250));
  toast('Gerando PDF...', '');
  try {
    const blob = await gerarPdfDaSheet();
    const filename = pdfFilenameForOS(data);
    const saved = await savePdfToFolder(blob, filename);
    if (saved) {
      toast(`PDF salvo: ${filename}`, 'ok');
      // Regrava a etiqueta junto: quem clica em "Salvar e Gerar PDF" espera
      // que TUDO que sai dessa OS pra disco fique atualizado.
      salvarPdfEtiquetasAuto(data, dadosEtiquetaParaOS(data));
      // Se gerou conjugada, gera o PDF dela tambem
      if (conjugada) {
        await new Promise(r => setTimeout(r, 400));
        renderPrintSheet(conjugada);
        await new Promise(r => setTimeout(r, 250));
        try {
          const blobC = await gerarPdfDaSheet();
          const fnC = pdfFilenameForOS(conjugada);
          const okC = await savePdfToFolder(blobC, fnC);
          if (okC) toast(`PDF conjugada salvo: ${fnC}`, 'ok');
          salvarPdfEtiquetasAuto(conjugada, dadosEtiquetaParaOS(conjugada));
        } catch (e) {
          console.warn('PDF conjugada', e);
        }
      }
      setTimeout(() => goto('lista-os'), 700);
    } else {
      // Sem pasta ou erro: continua na print page pra o usuario poder
      // ao menos imprimir manualmente ou tentar conectar a pasta.
    }
  } catch (e) {
    console.error('salvarEImprimir/PDF', e);
    toast('Falha ao gerar PDF: ' + (e.message || e), 'err');
  }
}

/* ========================================================= */
/*    ETIQUETAS ADESIVAS (1 por pagina, 10x5cm, LOTE 1..N)   */
/* ========================================================= */
// Conteúdo do pacote de reposição (a última etiqueta). Texto do usuário.
const ETIQUETA_CONTEUDO_REPOSICAO = 'Viés/Reposição/Ribana';

// Composição de um pacote de blusa de MOLETOM (360 peças = 36 blusas). Sai em
// cada etiqueta de pacote de moletom (não na de reposição). Duas linhas,
// agrupadas por quantidade (as de 36 e as de 72). Camiseta não recebe lista.
const ETIQUETA_COMPOSICAO_MOLETOM = [
  'Frente 36 · Costa 36 · Bolso 36 · Barra 36',
  'Mangas 72 · Capuz 72 · Punhos 72'
];

// Tamanhos da grade expandidos em PACOTES: um item por vaga de tamanho, na
// ordem P..G3. Segue a mesma regra por tipo de _expTotalTamanhosGrade:
//   • Camiseta: repete o tamanho conforme a quantidade (2M → 'M','M').
//   • Moletom : 1 item por tamanho distinto (multiplicador não repete pacote).
// Ex. camiseta: 2M-4G-2GG → ['M','M','G','G','G','G','GG','GG'].
// Ex. moletom : 2X P ao G3 → ['P','M','G','GG','G1','G2','G3'].
// Prefere a grade viva (como a folha e o volume), caindo no snapshot da OS.
function _tamanhosDaGradeExpandido(o) {
  const ordem = ['p','m','g','gg','g1','g2','g3'];
  const rotulo = { p:'P', m:'M', g:'G', gg:'GG', g1:'G1', g2:'G2', g3:'G3' };
  let tam = null;
  if (o && o.gradeId) {
    const g = (STATE.grades || []).find(x => x.id === o.gradeId);
    if (g && g.tamanhos) tam = g.tamanhos;
  }
  if (!tam && o && o.grade) tam = o.grade;
  const umPorTamanho = _osEhMoletom(o);
  const out = [];
  if (tam) ordem.forEach(k => {
    const q = parseInt(tam[k]) || 0;
    if (q <= 0) return;
    if (umPorTamanho) out.push(rotulo[k]);
    else for (let i = 0; i < q; i++) out.push(rotulo[k]);
  });
  return out;
}

// Uma etiqueta por pagina (100mm x 50mm), uma por PACOTE — mesma regra do
// volume de expedição: 1 por vaga de tamanho da grade + 1 de reposição. As
// etiquetas de tamanho sao iguais (só o LOTE muda); a ÚLTIMA é o pacote de
// reposição e mostra o conteúdo (${ETIQUETA_CONTEUDO_REPOSICAO}) no lugar de
// tamanho/qtde.
// Calcula os dados que vao para cada etiqueta a partir de uma OS. Centralizado
// num helper porque tambem e usado pelos auto-saves silenciosos de
// salvarOS/salvarEImprimir, fora do fluxo de impressao.
function dadosEtiquetaParaOS(o) {
  const os = o.os || o.codigo || '—';
  const marca = (o.griffeNome || o.griffe || 'MARCA').toUpperCase();
  const camadas = o.enfesto?.camadas || 0;
  const fasesP = o.fases || [];
  const tecsP = o.tecidos || [];
  const temMoletom = fasesP.some(f => {
    const t = STATE.tecidos.find(x => x.id === f.tecidoId);
    return t && categoriaEfetivaTecido(t) === 'moletom';
  }) || tecsP.some(t => {
    const tec = STATE.tecidos.find(x => x.id === t.tecidoId);
    return tec && categoriaEfetivaTecido(tec) === 'moletom';
  });
  const temMalha = !temMoletom && (
    fasesP.some(f => {
      const t = STATE.tecidos.find(x => x.id === f.tecidoId);
      return t && categoriaEfetivaTecido(t) === 'malha';
    }) || tecsP.some(t => {
      const tec = STATE.tecidos.find(x => x.id === t.tecidoId);
      return tec && categoriaEfetivaTecido(tec) === 'malha';
    })
  );
  const multPrincipal = temMoletom ? 1 : (temMalha ? 2 : 1);
  const totalGrade = o.grade?.total || 0;
  const qtde = (totalGrade > 0 && camadas > 0) ? (totalGrade * camadas * multPrincipal) : totalGrade;
  const sizesAtivos = ['p','m','g','gg','g1','g2','g3']
    .filter(k => (o.grade?.[k] || 0) > 0)
    .map(s => s.toUpperCase());
  const tam = sizesAtivos.join('-') || (o.grade?.descricao || '—');

  const desenho = o.desenhoId ? STATE.desenhos.find(x => x.id === o.desenhoId) : null;
  // Cor da PEÇA: corNomeCurto tira o tecido do nome da cor e o Set colapsa as
  // repetições — preto na malha + preto na ribana é "PRETO", não "PRETO/PRETO".
  const _corNome = id => id ? corNomeCurto(STATE.cores.find(c => c.id === id)?.nome || '') : '';
  const coresDesenho = [...new Set([
    _corNome(desenho?.corPrincipalId),
    _corNome(desenho?.corSecundariaId),
    _corNome(desenho?.corTerciariaId)
  ].filter(Boolean))];
  const cor = (coresDesenho.length > 1
    ? coresDesenho.join('/')
    : (corNomeCurto(o.fases?.[0]?.corNome || o.tecidos?.[0]?.corNome || '')
       || coresDesenho[0]
       || '—')).toString().toUpperCase();

  const desenhoNome = String(desenho?.desc || desenho?.codigo || o.codigo || '—')
    .split('|')[0]
    .trim()
    .toUpperCase();

  // Uma etiqueta por PACOTE — mesma regra do volume de expedição. Cada pacote
  // de tamanho leva o SEU tamanho (P, M, G…); a última etiqueta é o pacote de
  // reposição (conteúdo = ETIQUETA_CONTEUDO_REPOSICAO). Total = nº de vagas de
  // tamanho + 1. Mínimo 1 pra não bloquear OS sem grade.
  // Ex.: P-G1-G2 = 4 etiquetas (P, G1, G2, Reposição); 2M-4G-2GG = 9.
  const tamanhosPacotes = _tamanhosDaGradeExpandido(o);
  const temReposicao = tamanhosPacotes.length > 0;
  const numEtiquetas = temReposicao ? tamanhosPacotes.length + 1 : 1;

  // Moletom: cada etiqueta de pacote (de tamanho) recebe a lista de composição.
  const composicao = temMoletom ? ETIQUETA_COMPOSICAO_MOLETOM : null;

  return { marca, os, qtde, tam, cor, modelo: desenhoNome, numEtiquetas, tamanhosPacotes, temReposicao, composicao };
}

function imprimirEtiquetas(osId) {
  const o = STATE.ordens.find(x => x.id === osId);
  if (!o) { toast('OS não encontrada', 'err'); return; }

  const { marca, os, qtde, tam, cor, modelo: desenhoNome, numEtiquetas, tamanhosPacotes, temReposicao, composicao } = dadosEtiquetaParaOS(o);

  const escEt = s => String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  // Cada etiqueta é um pacote: as de tamanho mostram o SEU tamanho em destaque
  // (fonte dobrada); a última é o pacote de reposição, com o conteúdo. Moletom:
  // as etiquetas de tamanho recebem a lista de composição do pacote.
  const compHtml = composicao ? composicao.map(c => `<div class="comp">${escEt(c)}</div>`).join('') : '';
  const corpo = Array.from({ length: numEtiquetas }, (_, i) => {
    const ehRep = temReposicao && i === numEtiquetas - 1;
    const destaque = ehRep
      ? `<div class="big rep">${escEt(ETIQUETA_CONTEUDO_REPOSICAO)}</div>`
      : `<div class="big">${escEt((tamanhosPacotes && tamanhosPacotes[i]) || tam)}</div>${compHtml}`;
    return `
    <div class="page">
      <div class="label">
        <div class="head">${escEt(marca)}</div>
        <div class="row">OS: ${escEt(os)}</div>
        <div class="row">MODELO: ${escEt(desenhoNome)}</div>
        <div class="row">QTDE: ${escEt(qtde)}</div>
        <div class="row">TAM: ${escEt(tam)}</div>
        <div class="row">COR: ${escEt(cor)}</div>
        <div class="row">LOTE: ${i + 1}/${numEtiquetas}</div>
        ${destaque}
      </div>
    </div>`;
  }).join('');

  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Etiquetas — OS ${escEt(os)}</title>
<style>
  /* Pagina 10x5cm (100mm x 50mm landscape), 1 etiqueta por pagina. */
  /* Total de paginas = quantidade de tamanhos ativos na grade da OS. */
  @page { size: 100mm 50mm; margin: 0; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; background: #fff; color: #000; }
  body { font-family: 'IBM Plex Sans', system-ui, -apple-system, Segoe UI, Arial, sans-serif; }
  .toolbar {
    padding: 12px;
    background: #f4f4f4;
    border-bottom: 1px solid #ccc;
    text-align: center;
    position: sticky;
    top: 0;
    z-index: 10;
  }
  .toolbar button {
    padding: 8px 18px;
    font-size: 14px;
    cursor: pointer;
    margin: 0 4px;
    border: 1px solid #888;
    background: #fff;
    border-radius: 3px;
  }
  .toolbar button.primary {
    background: #16a34a;
    color: #fff;
    border-color: #15803d;
    font-weight: 600;
  }
  .page {
    width: 100mm;
    height: 50mm;
    padding: 2mm;
    page-break-after: always;
    margin: 0 auto 6px auto;
    background: #fff;
  }
  .page:last-child { page-break-after: auto; }
  .label {
    width: 100%;
    height: 100%;
    border: 1px solid #000;
    padding: 1.5mm 3.5mm;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }
  /* Mesma fonte/peso pra marca e linhas — visual uniforme; o que muda */
  /* e so alinhamento (marca centralizada) e o separador fino abaixo.  */
  .label .head,
  .label .row {
    font-size: 11pt;
    font-weight: 800;
    letter-spacing: .03em;
    line-height: 1.05;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .label .head {
    text-align: center;
    border-bottom: 1px solid #000;
    padding-bottom: 0.6mm;
    margin-bottom: 0.4mm;
  }
  .label .row { text-align: left; }
  /* Tamanho do pacote (ou conteúdo da reposição) em destaque: fonte dobrada. */
  .label .big {
    font-size: 22pt;         /* dobro das linhas (11pt) */
    font-weight: 800;
    text-align: center;
    line-height: 1.0;
    letter-spacing: .02em;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    margin-top: 0.4mm;
  }
  .label .big.rep { font-size: 15pt; }  /* conteúdo é texto mais longo */
  /* Composição do pacote de moletom — linhas pequenas abaixo do tamanho. */
  .label .comp {
    font-size: 7.5pt;
    font-weight: 600;
    text-align: center;
    line-height: 1.15;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  @media print {
    .toolbar { display: none !important; }
    .page { margin: 0; }
    body { background: #fff; }
  }
</style>
</head>
<body>
  <div class="toolbar">
    <button class="primary" onclick="window.print()">🖨 Imprimir</button>
    <button onclick="window.close()">Fechar</button>
    <span style="margin-left:12px;color:#555;font-size:13px;">${numEtiquetas} etiqueta${numEtiquetas>1?'s':''} · LOTE 1${numEtiquetas>1?'..'+numEtiquetas:''} · 10×5cm</span>
  </div>
  ${corpo}
  <script>
    window.addEventListener('load', () => { setTimeout(() => window.print(), 350); });
  </script>
</body>
</html>`;

  const w = window.open('', '_blank', 'width=900,height=1100');
  if (!w) {
    toast('Popup bloqueado pelo navegador. Permita popups deste site.', 'err');
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();

  // Auto-save em segundo plano: gera o PDF e salva em <pasta-pdf>/etiquetas/
  // (subpasta criada se nao existir). Silencioso se a pasta nao estiver
  // conectada — nao bloqueia o popup de impressao.
  salvarPdfEtiquetasAuto(o, { marca, os, qtde, tam, cor, modelo: desenhoNome, numEtiquetas });
}

function imprimirEtiquetasAtual() {
  if (!printOsAtual) { toast('Abra uma OS antes', 'err'); return; }
  imprimirEtiquetas(printOsAtual.id);
}

async function salvarEImprimirEtiquetas() {
  const data = coletaOS();
  if (!data.os && !data.codigo) {
    return toast('Preencha ao menos número da OS ou código do desenho', 'err');
  }
  if (!validarAntesDeSalvar(data)) return;
  const idx = STATE.ordens.findIndex(o => o.id === data.id);
  if (idx >= 0) STATE.ordens[idx] = data; else STATE.ordens.push(data);
  await saveState('ordens');
  await atualizarCounterOS(data.os);
  osEditId = null;
  imprimirEtiquetas(data.id);
}

// Usado so pelo botao "Imprimir / Salvar PDF" (window.print()). O caminho
// principal — "Salvar e Gerar PDF" — nao passa por aqui: la o jsPDF ja
// encaixa a foto na A4 sozinho.
//
// Esta funcao NAO mexe mais na geometria da .sheet (largura/padding/altura).
// Ela ja E a A4; reescrever isso era o que fazia o papel sair diferente da
// tela. O unico ajuste possivel e encolher a folha inteira pelo wrapper.
function ajustarImpressaoParaA4() {
  const sheet = document.querySelector('.sheet');
  const scaler = document.querySelector('.sheet-scaler');
  if (!sheet || !scaler) return;

  // Mede a folha na geometria de saida: .pdf-capture zera a ampliacao de
  // leitura, entao scrollHeight vem em mm reais de A4.
  scaler.style.removeProperty('zoom');
  document.body.classList.add('pdf-capture');
  void sheet.offsetHeight; // forca reflow pra leitura correta

  const pxPerMm = 3.7795275591;
  const maxHpx = 297 * pxPerMm; // A4 cheia — a margem ja esta no padding
  const natH = sheet.scrollHeight;

  document.body.classList.remove('pdf-capture');

  // A .sheet tem min-height 297mm, entao so passa disso quando o conteudo
  // realmente estourou a pagina. Ai encolhe tudo por igual pra caber em 1
  // folha (1% de folga evita o caso borderline por arredondamento).
  // zoom afeta LAYOUT, entao tabelas, fontes e quebras encolhem juntas.
  const scale = natH > maxHpx ? (maxHpx / natH) * 0.99 : 1;
  scaler.style.setProperty('zoom', scale.toFixed(4), 'important');
}

window.addEventListener('beforeprint', ajustarImpressaoParaA4);
window.addEventListener('afterprint', function() {
  const scaler = document.querySelector('.sheet-scaler');
  if (!scaler) return;
  // Devolve a ampliacao de leitura da tela (volta pra regra do styles.css).
  scaler.style.removeProperty('zoom');
  document.body.classList.remove('pdf-capture');
});

/* ========================================================= */
/*                    LISTA DE OS                            */
/* ========================================================= */
// Número da OS como inteiro p/ ordenação (ex.: "0282" -> 282). OS sem número
// (salva só com código) vai pro fim da lista.
function numeroOSordenacao(o) {
  const n = parseInt(String(o?.os || '').replace(/\D/g, ''), 10);
  return Number.isNaN(n) ? Infinity : n;
}

function renderListaOS() {
  const tb = document.getElementById('tbl-os');
  if (!STATE.ordens.length) { tb.innerHTML = `<tr><td colspan="8" class="empty">Nenhuma OS cadastrada ainda.</td></tr>`; return; }
  // Ordem decrescente pelo número da OS (maior primeiro); OS sem número no fim.
  const ordenadas = STATE.ordens.slice().sort((a, b) => {
    const na = numeroOSordenacao(a), nb = numeroOSordenacao(b);
    if (na === Infinity && nb === Infinity) return String(a.os || '').localeCompare(String(b.os || ''));
    if (na === Infinity) return 1;
    if (nb === Infinity) return -1;
    return nb - na || String(b.os || '').localeCompare(String(a.os || ''));
  });
  // Filtro por número da OS (busca livre; ignora espaços).
  const buscaEl = document.getElementById('busca-os');
  const termo = (buscaEl ? buscaEl.value : '').trim().toLowerCase();
  const filtradas = termo
    ? ordenadas.filter(o => String(o.os || '').toLowerCase().includes(termo))
    : ordenadas;
  if (!filtradas.length) { tb.innerHTML = `<tr><td colspan="8" class="empty">Nenhuma OS encontrada para "${esc(termo)}".</td></tr>`; return; }
  tb.innerHTML = filtradas.map(o => {
    // Mesma miniatura da lista de desenhos: acha o desenho técnico da OS por
    // desenhoId (padrão) ou, para OS antigas sem esse vínculo, pelo código.
    const des = (o.desenhoId && STATE.desenhos.find(d => d.id === o.desenhoId))
      || (o.codigo && STATE.desenhos.find(d => (d.codigo || '').trim() === (o.codigo || '').trim()))
      || null;
    const thumb = `<div style="width:60px;height:45px;background:#f5f2ea;display:flex;align-items:center;justify-content:center;border:1px solid var(--line);overflow:hidden">${des && des.img ? `<img src="${des.img}" style="max-width:100%;max-height:100%;object-fit:contain;">` : '—'}</div>`;
    return `
    <tr>
      <td>${thumb}</td>
      <td><strong>${esc(o.os)||'—'}</strong></td>
      <td><span class="badge">${esc(o.codigo)||'—'}</span></td>
      <td>${esc(o.modeloNome)||'—'}</td>
      <td>${esc(o.colecaoNome)||'—'}</td>
      <td>${esc(formatDate(o.data))}</td>
      <td>${o.grade?.total||0} pç</td>
      <td class="col-actions row-actions">
        <button class="edit" onclick="verOS('${o.id}')">visualizar</button>
        <button class="edit" onclick="imprimirEtiquetas('${o.id}')">etiquetas</button>
        <button class="edit" onclick="editarOS('${o.id}')">editar</button>
        <button class="edit" onclick="duplicarOS('${o.id}')">duplicar</button>
        <button class="del admin-only" onclick="excluirOS('${o.id}')">excluir</button>
      </td>
    </tr>`;
  }).join('');
}

function formatDate(iso) {
  if (!iso) return '—';
  const [y,m,d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

let printOsAtual = null;

// Marca/desmarca etapa do checklist da OS pronta. Persiste em o.progresso e
// salva STATE.ordens — outros usuarios veem a evolucao ao reabrir a OS.
async function togglarChecklistEtapa(osId, etapaNome, checked) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.etapasCheck = os.progresso.etapasCheck || {};
  os.progresso.etapasSeq = os.progresso.etapasSeq || {};
  if (checked) {
    os.progresso.etapasCheck[etapaNome] = true;
    // Carimbo de ordem de marcação: o volume da OS fica no campo da etapa marcada
    // por ÚLTIMO (faseAtualOS usa o maior seq). Date.now() = "mais recente".
    os.progresso.etapasSeq[etapaNome] = Date.now();
  } else {
    delete os.progresso.etapasCheck[etapaNome];
    delete os.progresso.etapasSeq[etapaNome];
  }
  try { await saveState('ordens'); } catch (e) { console.warn('togglarChecklistEtapa', e); }
  // Marcar "Expedição" aqui É o ato de selecionar a OS pra ser expedida: ela
  // entra sozinha no plano, na próxima janela. Trocar a janela é depois, no
  // planejamento.
  try { await sincronizarPlanoExpedicaoDaOS(os, etapaNome, checked); }
  catch (e) { console.warn('sincronizarPlanoExpedicaoDaOS', e); }
}

async function togglarChecklistTarefa(osId, etapaNome, tarefaNome, checked) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.tarefasCheck = os.progresso.tarefasCheck || {};
  os.progresso.tarefasCheck[etapaNome] = os.progresso.tarefasCheck[etapaNome] || {};
  if (checked) os.progresso.tarefasCheck[etapaNome][tarefaNome] = true;
  else delete os.progresso.tarefasCheck[etapaNome][tarefaNome];
  try { await saveState('ordens'); } catch (e) { console.warn('togglarChecklistTarefa', e); }
}

async function togglarChecklistEnfesto(osId, ordem, checked) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.enfestosCheck = os.progresso.enfestosCheck || {};
  if (checked) os.progresso.enfestosCheck[ordem] = true;
  else delete os.progresso.enfestosCheck[ordem];
  try { await saveState('ordens'); } catch (e) { console.warn('togglarChecklistEnfesto', e); }
}

// Normaliza para HH:MM o que foi digitado nos campos de horário da folha (os
// Início/Fim de enfesto e de corte). Eram texto livre: quem digitava "730" no
// ritmo do chão de fábrica via "730" na folha, e cada pessoa gravava de um jeito
// ("7h30", "7:3", "0730"), o que impedia comparar tempos entre fases e OSs.
// Aceita o jeito rápido de digitar e devolve sempre o mesmo formato:
//   "7" → 07:00 · "19" → 19:00 · "730" → 07:30 · "0730" → 07:30 · "7:5" → 07:05
// Texto que não vira hora válida (ex.: "2575") volta como veio — reformatar
// destruiria o que a pessoa escreveu sem ela perceber.
function _horaFmt(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  let h, m;
  if (s.includes(':')) {
    const [a, b] = s.split(':');
    h = parseInt(String(a).replace(/\D/g, ''), 10);
    m = parseInt(String(b).replace(/\D/g, ''), 10) || 0;
  } else {
    const d = s.replace(/\D/g, '');
    if (!d) return s;
    if (d.length <= 2) { h = parseInt(d, 10); m = 0; }
    else { m = parseInt(d.slice(-2), 10); h = parseInt(d.slice(0, -2), 10); }
  }
  if (!Number.isFinite(h) || !Number.isFinite(m) || h > 23 || m > 59) return s;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

// Salva o tempo de Início/Fim digitado em cada fase de enfesto na folha
// impressa. campo ∈ {enfIni, enfFim, corIni, corFim} (enfesto e corte).
// Valor vazio remove a chave. Persiste em progresso.enfestosTempos[ordem].
async function salvarTempoEnfesto(osId, ordem, campo, valor) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.enfestosTempos = os.progresso.enfestosTempos || {};
  os.progresso.enfestosTempos[ordem] = os.progresso.enfestosTempos[ordem] || {};
  const v = (valor || '').trim();
  if (v) os.progresso.enfestosTempos[ordem][campo] = v;
  else delete os.progresso.enfestosTempos[ordem][campo];
  try { await saveState('ordens'); } catch (e) { console.warn('salvarTempoEnfesto', e); }
}

// Salva o valor digitado (à mão) de cada TOM em cada fase de enfesto na folha.
// As tonalidades podem variar em qualquer fase, então cada fase tem seus campos
// de Tom 1/2/3 (os mesmos tons ativos no "Total por tamanho"). Texto livre —
// persiste em progresso.enfestosTons[ordem][tom]; vazio remove a chave.
async function salvarTomEnfesto(osId, ordem, tom, valor) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.enfestosTons = os.progresso.enfestosTons || {};
  os.progresso.enfestosTons[ordem] = os.progresso.enfestosTons[ordem] || {};
  const v = (valor || '').trim();
  if (v) os.progresso.enfestosTons[ordem][tom] = v;
  else delete os.progresso.enfestosTons[ordem][tom];
  try { await saveState('ordens'); } catch (e) { console.warn('salvarTomEnfesto', e); }
}

// Salva o tempo de Início/Fim do corte, mostrado junto da etapa "Corte" em
// Etapas de Produção. Um par único por OS. campo ∈ {ini, fim}.
async function salvarTempoCorte(osId, campo, valor) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.corteTempo = os.progresso.corteTempo || {};
  const v = (valor || '').trim();
  if (v) os.progresso.corteTempo[campo] = v;
  else delete os.progresso.corteTempo[campo];
  try { await saveState('ordens'); } catch (e) { console.warn('salvarTempoCorte', e); }
}

// Salva as observações digitadas direto na folha de OS (caixa "Observações").
// Grava no mesmo campo o.obs usado pelo formulário de cadastro (f-obs).
async function salvarObsOS(osId, valor) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.obs = (valor || '').trim();
  try { await saveState('ordens'); } catch (e) { console.warn('salvarObsOS', e); }
}

// Calcula os tons efetivamente marcados como prefixo consecutivo: Tom 2 so
// vale se Tom 1 estiver marcado; Tom 3 so vale se Tom 1 e Tom 2 estiverem.
// Sanitiza dados antigos ou estado inconsistente sem precisar limpar.
function tonsEfetivos(ttTons) {
  const out = [];
  if (ttTons && ttTons[1]) out.push(1);
  if (ttTons && ttTons[1] && ttTons[2]) out.push(2);
  if (ttTons && ttTons[1] && ttTons[2] && ttTons[3]) out.push(3);
  return out;
}

// Multiplicador de peças por camada: quantas unidades cada camada rende em cada
// vaga da grade. Moletom = 1 (1 camada = 1 blusa); malha sem moletom (camiseta)
// = 2. Sem isso o total por tamanho sai pela metade na camiseta.
function multiplicadorPecaOS(o) {
  const cat = tecId => {
    const t = (STATE.tecidos || []).find(x => x.id === tecId);
    return t ? categoriaEfetivaTecido(t) : null;
  };
  const fases = (o && o.fases) || [];
  const tecs = (o && o.tecidos) || [];
  const tem = c => fases.some(f => cat(f.tecidoId) === c) || tecs.some(t => cat(t.tecidoId) === c);
  if (tem('moletom')) return 1;
  return tem('malha') ? 2 : 1;
}

// Fonte ÚNICA dos números do "Total por tamanho": quantidade por tamanho, por
// tonalidade, total de cada tom e total geral. A folha de OS e a folha de OE
// (plano de expedição) leem daqui, então não têm como mostrar números diferentes.
//
// Regra do balanceador (confirmada com o Junior): o V é uniforme por linha de tom
// (o mesmo número em todas as células visíveis daquele tom); o ÚLTIMO tom marcado
// é o balanceador e recebe, em cada tamanho, o total da coluna menos a soma dos V
// dos tons editáveis — assim a soma das colunas bate com a linha "Total geral".
function totaisPorTamanhoTomOS(o) {
  const keys = ['p','m','g','gg','g1','g2','g3'];
  const g = (o && o.grade) || {};
  const cam = (o && o.enfesto && o.enfesto.camadas) || 0;
  const mult = multiplicadorPecaOS(o);
  const prog = (o && o.progresso) || {};
  const colTotal = k => (g[k] || 0) * cam * mult;
  const tamanhos = keys.filter(k => (g[k] || 0) > 0);
  const tons = tonsEfetivos(prog.totalTamanhoTons || {});
  const valores = prog.totalTamanhoTomValor || {};
  const balancer = tons.length ? tons[tons.length - 1] : null;
  const vTom = tom => Math.max(0, Number(valores[tom]) || 0);
  let somaEditaveis = 0;
  tons.forEach(t => { if (t !== balancer) somaEditaveis += vTom(t); });
  // Enquanto NADA foi digitado, as linhas de tom saem VAZIAS: a quantidade por
  // tamanho fica só na linha "Total por tamanho", logo acima. Repetir o mesmo
  // número numa linha de tom (fosse o Tom 1 ou o balanceador) duplicava a linha
  // de cima e confundia quem lê. A divisão só aparece quando alguém digita.
  const semDigitacao = somaEditaveis === 0;
  const linhas = tons.map(tom => {
    const cels = {};
    let total = 0;
    tamanhos.forEach(k => {
      let v;
      if (semDigitacao) v = 0;
      else if (tom === balancer) v = Math.max(0, colTotal(k) - somaEditaveis);
      else v = vTom(tom);
      cels[k] = v;
      total += v;
    });
    return { tom, cels, total, balanceador: tom === balancer, editavel: tom !== balancer };
  });
  return {
    keys, tamanhos, tons, linhas, colTotal, vTom,
    balancer, somaEditaveis, semDigitacao,
    totalGeral: (g.total || 0) * cam * mult,
  };
}

async function togglarTotalTamanhoTom(osId, tom, checked) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.totalTamanhoTons = os.progresso.totalTamanhoTons || {};
  const tNum = Number(tom);
  const t = os.progresso.totalTamanhoTons;
  if (checked) {
    // Bloqueia se prereq nao atendido (Tom 2 exige Tom 1; Tom 3 exige 1+2)
    if (tNum === 2 && !t[1]) {
      if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
      return;
    }
    if (tNum === 3 && (!t[1] || !t[2])) {
      if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
      return;
    }
    t[tNum] = true;
  } else {
    delete t[tNum];
    // Cascade: desmarcar Tom 1 derruba 2 e 3; desmarcar Tom 2 derruba 3
    if (tNum === 1) { delete t[2]; delete t[3]; }
    else if (tNum === 2) { delete t[3]; }
  }
  try { await saveState('ordens'); } catch (e) { console.warn('togglarTotalTamanhoTom', e); }
  // Mudou o nº de tonalidades → mudou o volume: cada tom é ensacado separado.
  // Propaga para as expedições futuras desta OS, senão a OE seguiria com o
  // número congelado de quando a OS entrou no plano.
  const nCargas = await propagarVolumesExpedicaoOS(os);
  if (nCargas) {
    toast(`Volume da expedição atualizado para ${_expSugestaoVolumes(os)} — ${nCargas} carga(s)`, 'ok');
    const secExp = document.querySelector('section.page[data-page="expedicao"]');
    if (secExp && !secExp.classList.contains('hidden')) renderExpedicaoPlano();
  }
  if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
}

// Salva o valor uniforme V do tom (mesmo numero em todas as celulas visiveis
// da linha) e atualiza o DOM dos demais inputs sincronizados, das celulas
// do balanceador e das colunas "Total". O re-render completo nao acontece
// aqui pra preservar o foco do input quando o usuario tabula entre celulas.
async function salvarValorTotalTamanhoTom(osId, tom, valor) {
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.totalTamanhoTomValor = os.progresso.totalTamanhoTomValor || {};
  const tNum = Number(tom);
  // Clampa V: a soma dos V dos tons editaveis em cada coluna nao pode passar
  // de colTotal daquela coluna — assim o balanceador (ultimo tom marcado)
  // nunca fica negativo em nenhuma coluna. Como V e uniforme por linha, o
  // gargalo e a menor colTotal entre as colunas visiveis.
  const tomsSel = tonsEfetivos(os.progresso.totalTamanhoTons || {});
  const balancerTom = tomsSel.length ? tomsSel[tomsSel.length - 1] : null;
  let somaOutros = 0;
  tomsSel.forEach(tt => {
    if (tt === balancerTom || tt === tNum) return;
    somaOutros += Math.max(0, Number(os.progresso.totalTamanhoTomValor[tt]) || 0);
  });
  const g = os.grade || {};
  const cam = os.enfesto?.camadas || 0;
  const mult = calcularMultPrincipalImpressao(os);
  let minCol = Infinity;
  ['p','m','g','gg','g1','g2','g3'].forEach(k => {
    if ((g[k] || 0) > 0) {
      const colTotal = g[k] * cam * mult;
      if (colTotal < minCol) minCol = colTotal;
    }
  });
  const max = minCol === Infinity ? 0 : Math.max(0, minCol - somaOutros);
  const n = Math.max(0, Math.min(max, Math.floor(Number(valor) || 0)));
  os.progresso.totalTamanhoTomValor[tNum] = n;
  // Ajusta o DOM se o valor digitado foi clampado pra menos
  const txt = n > 0 ? String(n) : '';
  document.querySelectorAll(`input[data-tt-tom-input="${tNum}"]`).forEach(i => {
    if (i.value !== txt) i.value = txt;
  });
  atualizarLinhasTomNoDOM();
  try { await saveState('ordens'); } catch (e) { console.warn('salvarValorTotalTamanhoTom', e); }
}

// Sincroniza visualmente o valor V entre todos os inputs da mesma linha de
// tom (digitar em uma celula preenche as outras com o mesmo numero) e
// recalcula no DOM as celulas do balanceador e as colunas "Total" das
// linhas — sem re-render pra preservar o foco do input.
function propagarValorTomTamanho(input, tom) {
  const v = input.value;
  document.querySelectorAll(`input[data-tt-tom-input="${tom}"]`).forEach(i => {
    if (i !== input) i.value = v;
  });
  atualizarLinhasTomNoDOM();
}

function atualizarLinhasTomNoDOM() {
  const os = printOsAtual;
  if (!os) return;
  const balancerCells = document.querySelectorAll('[data-tt-balancer-cell]');
  if (!balancerCells.length && !document.querySelector('[data-tt-row-total]')) return;
  // Toms editaveis = aqueles que tem input renderizado (um por linha basta)
  const vPorTom = {};
  [1,2,3].forEach(tt => {
    const first = document.querySelector(`input[data-tt-tom-input="${tt}"]`);
    if (first) vPorTom[tt] = Math.max(0, Number(first.value) || 0);
  });
  // Atualiza cada celula do balanceador: colTotal(k) - soma dos V editaveis
  balancerCells.forEach(c => {
    const size = c.dataset.ttBalancerSize;
    const colTotal = calcularColTotalAlvoImpressao(os, size);
    let somaEditaveis = 0;
    Object.values(vPorTom).forEach(v => { somaEditaveis += v; });
    const v = Math.max(0, colTotal - somaEditaveis);
    c.textContent = v > 0 ? String(v) : '';
  });
  // Atualiza as colunas "Total" de cada linha de tom (soma da linha)
  document.querySelectorAll('[data-tt-row-total]').forEach(c => {
    const tt = Number(c.dataset.ttRowTotal);
    let sum = 0;
    if (vPorTom[tt] != null) {
      // Tom editavel: V × numero de celulas visiveis (inputs)
      const inputs = document.querySelectorAll(`input[data-tt-tom-input="${tt}"]`);
      sum = vPorTom[tt] * inputs.length;
    } else {
      // Tom balanceador: soma das celulas balanceadoras
      document.querySelectorAll(`[data-tt-balancer-cell][data-tt-balancer-tom="${tt}"]`).forEach(bc => {
        sum += Math.max(0, Number(bc.textContent) || 0);
      });
    }
    c.textContent = sum > 0 ? String(sum) : '';
  });
}

// Recalcula multiplicador principal (moletom=1, malha=2, outro=1) usado na
// folha de impressao — mesma logica usada pra montar a linha "Total geral".
function calcularMultPrincipalImpressao(o) {
  const fasesP = o.fases || [];
  const tecsP = o.tecidos || [];
  const temMoletom = fasesP.some(f => {
    const t = STATE.tecidos.find(x => x.id === f.tecidoId);
    return t && categoriaEfetivaTecido(t) === 'moletom';
  }) || tecsP.some(t => {
    const tec = STATE.tecidos.find(x => x.id === t.tecidoId);
    return tec && categoriaEfetivaTecido(tec) === 'moletom';
  });
  const temMalha = !temMoletom && (
    fasesP.some(f => {
      const t = STATE.tecidos.find(x => x.id === f.tecidoId);
      return t && categoriaEfetivaTecido(t) === 'malha';
    }) || tecsP.some(t => {
      const tec = STATE.tecidos.find(x => x.id === t.tecidoId);
      return tec && categoriaEfetivaTecido(tec) === 'malha';
    })
  );
  return temMoletom ? 1 : (temMalha ? 2 : 1);
}

function calcularTotalGeralAlvoImpressao(o) {
  const g = o.grade || {};
  const cam = o.enfesto?.camadas || 0;
  return (g.total || 0) * cam * calcularMultPrincipalImpressao(o);
}

function calcularColTotalAlvoImpressao(o, size) {
  const g = o.grade || {};
  const cam = o.enfesto?.camadas || 0;
  return (g[size] || 0) * cam * calcularMultPrincipalImpressao(o);
}

// Sincroniza o estado dos <input.os-check> da folha com o.progresso, sem
// re-renderizar a sheet inteira. Usado pelo realtime/polling para refletir
// mudancas de outros usuarios sem piscar a tela nem perder o scroll.
function aplicarProgressoCheckboxes(os) {
  if (!os) return;
  const prog = os.progresso || {};
  document.querySelectorAll('.os-check[data-etapa]').forEach(inp => {
    const etapaNome = inp.dataset.etapa;
    const tarefaNome = inp.dataset.tarefa;
    const desejado = tarefaNome
      ? !!prog.tarefasCheck?.[etapaNome]?.[tarefaNome]
      : !!prog.etapasCheck?.[etapaNome];
    if (inp.checked !== desejado) inp.checked = desejado;
  });
  document.querySelectorAll('.os-check[data-enfesto]').forEach(inp => {
    const desejado = !!prog.enfestosCheck?.[inp.dataset.enfesto];
    if (inp.checked !== desejado) inp.checked = desejado;
  });
  document.querySelectorAll('.os-check[data-tt-tom]').forEach(inp => {
    const tom = inp.dataset.ttTom;
    const desejado = !!prog.totalTamanhoTons?.[tom];
    if (inp.checked !== desejado) inp.checked = desejado;
  });
  // Tempos de enfesto (texto, por fase) e de corte (par único da etapa Corte).
  // Não mexe no campo em foco pra não atropelar quem está digitando.
  document.querySelectorAll('input[data-enf-tempo]').forEach(inp => {
    if (inp === document.activeElement) return;
    const ord = inp.dataset.enfTempo;
    const campo = inp.dataset.enfCampo;
    const desejado = prog.enfestosTempos?.[ord]?.[campo] || '';
    if (inp.value !== desejado) inp.value = desejado;
  });
  document.querySelectorAll('input[data-corte-tempo]').forEach(inp => {
    if (inp === document.activeElement) return;
    const desejado = prog.corteTempo?.[inp.dataset.corteTempo] || '';
    if (inp.value !== desejado) inp.value = desejado;
  });
}

function verOS(id) {
  const o = STATE.ordens.find(x => x.id === id);
  if (!o) return;
  printOsAtual = o;
  renderPrintSheet(o);
  goto('print');
  // Auto-save em segundo plano: gera o PDF e salva na pasta conectada
  // (se houver). Nao bloqueia o usuario — ele ja pode imprimir fisico
  // imediatamente. Sem pasta conectada, nao faz nada (silencioso).
  autoSalvarPdfPrintAtual(o);
}

async function autoSalvarPdfPrintAtual(o) {
  const handle = pdfFolderHandle || (await loadPdfFolderHandle());
  if (!handle) return;
  const ok = await ensureFolderPermission(handle, 'readwrite');
  if (!ok) return;
  pdfFolderHandle = handle;
  // Da tempo do .sheet ficar com layout calculado apos o goto('print')
  await new Promise(r => setTimeout(r, 250));
  try {
    const blob = await gerarPdfDaSheet();
    const filename = pdfFilenameForOS(o);
    const fileHandle = await handle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    toast(`PDF salvo: ${filename}`, 'ok');
  } catch (e) {
    console.warn('autoSalvarPdfPrintAtual', e);
  }
}

function editarOsAtual() {
  if (printOsAtual) editarOS(printOsAtual.id);
}

function editarOS(id) {
  const o = STATE.ordens.find(x => x.id === id);
  if (!o) return;
  osEditId = id;
  goto('nova-os');
  // precisa de timeout curto pra select options já estarem renderizadas
  setTimeout(() => {
    document.getElementById('os-form-title').textContent = 'Editar OS ' + (o.os || o.codigo || '');
    document.getElementById('f-id').value = o.id;
    document.getElementById('f-os').value = o.os || '';
    document.getElementById('f-codigo').value = o.codigo || '';
    document.getElementById('f-data').value = o.data || '';
    // coordenado agora é select de desenho
    document.getElementById('f-coordenado').value = o.coordenadoId || '';
    document.getElementById('f-colecao').value = o.colecaoId || '';
    document.getElementById('f-modelo').value = o.modeloId || '';
    // novos selects do cabeçalho — tenta por ID; se vier de uma OS antiga (texto livre), tenta casar por nome
    setSelectByIdOrName('f-bloco', o.blocoId, o.bloco || o.blocoNome, STATE.blocos);
    setSelectByIdOrName('f-linha', o.linhaId, o.linha || o.linhaNome, STATE.linhas);
    setSelectByIdOrName('f-griffe', o.griffeId, o.griffe || o.griffeNome, STATE.marcas);
    setSelectByIdOrName('f-base', o.baseId, o.base || o.baseNome, STATE.bases);
    setSelectByIdOrName('f-designer', o.designerId, o.designer || o.designerNome, STATE.equipe);
    setSelectByIdOrName('f-ftec', o.ftecId, o.ftec || o.ftecNome, STATE.equipe);
    document.getElementById('f-desenho').value = o.desenhoId || '';
    previewDesenhoSelecionado();
    // Em edicao, aplica o mesmo filtro de "Nova OS" (categoria do desenho +
    // tipoPeca do modelo + variacao). A grade ja selecionada e preservada
    // via extraIds, mesmo que nao case com o filtro atual — isso garante que
    // o usuario continua vendo a opcao salva.
    const gradeEl = document.getElementById('f-grade-preset');
    if (gradeEl) {
      fillSelect('f-grade-preset', gradesParaDropdownOS([o.gradeId]), 'nome', '— nenhuma —');
      gradeEl.value = o.gradeId || '';
    }
    document.getElementById('f-grade-desc').value = o.grade?.descricao || '';
    ['p','m','g','gg','g1','g2','g3'].forEach(k => {
      document.getElementById('f-gr-'+k).value = o.grade?.[k] || 0;
    });
    // enfesto — blocos (novo) ou legado (comprimento/largura único)
    const blocosSalvos = Array.isArray(o.enfesto?.blocos) && o.enfesto.blocos.length
      ? o.enfesto.blocos
      : (o.enfesto?.comprimento || o.enfesto?.largura)
        ? [{ comp: o.enfesto.comprimento, larg: o.enfesto.largura }]
        : [{}];
    // Recupera nomeTecido correspondente a cada bloco — via nomeTecido salvo, ou lookup nas fases da OS
    const blocosComNomes = blocosSalvos.map((b, i) => {
      let nomeTecido = b.nomeTecido || '';
      if (!nomeTecido && Array.isArray(o.fases)) {
        const fase = o.fases.find(f => (f.ordem || 0) === (i+1));
        if (fase) nomeTecido = fase.nome || fase.tecidoNome || '';
      }
      return { ...b, nomeTecido };
    });
    renderEnfestoBlocos(blocosComNomes.length, blocosComNomes);
    document.getElementById('f-enf-camadas').value = o.enfesto?.camadas || '';
    document.getElementById('f-enf-target').value = o.enfesto?.target || '';
    document.getElementById('f-obs').value = o.obs || '';
    document.getElementById('f-atencao').value = o.atencao || '';
    // tecidos
    document.getElementById('tecidos-rows').innerHTML = '';
    (o.tecidos||[]).forEach(t => addTecidoRow(t));
    if (!o.tecidos?.length) { addTecidoRow(); addTecidoRow(); }
    // variantes
    document.getElementById('variantes-rows').innerHTML = '';
    (o.variantes||[]).forEach(vv => addVarianteRow(vv));
    if (!o.variantes?.length) addVarianteRow();
    // componentes
    document.getElementById('componentes-rows').innerHTML = '';
    (o.componentes||[]).forEach(c => addComponenteRow(c));
    if (!o.componentes?.length) { addComponenteRow(); addComponenteRow(); }
    // aviamentos
    document.getElementById('aviamentos-rows').innerHTML = '';
    (o.aviamentos||[]).forEach(a => addAviamentoRow(a));
    // etapas — marca as que estão em o.etapas e aplica a ordem salva
    document.querySelectorAll('#etapas-container .etapa-check').forEach(lbl => {
      const input = lbl.querySelector('input');
      const on = (o.etapas||[]).includes(input.value);
      input.checked = on;
      lbl.classList.toggle('checked', on);
    });
    aplicarOrdemEtapas(o.etapas || []);
    atualizarCalculosEnfesto();
    osEditId = null; // reset para permitir nova edição após salvar
  }, 60);
}

async function excluirOS(id) {
  if (!exigirAdmin('excluir OS')) return;
  if (!confirm('Excluir esta OS?')) return;
  STATE.ordens = STATE.ordens.filter(x => x.id !== id);
  await saveState('ordens');
  await estornarBaixaEstoqueOS(id);
  toast('OS excluída', 'ok');
  renderListaOS();
}

async function duplicarOS(id) {
  const o = STATE.ordens.find(x => x.id === id);
  if (!o) return toast('OS não encontrada', 'err');
  // Deep clone — id, numero da OS e data sao regerados; resto e copia exata
  const copia = JSON.parse(JSON.stringify(o));
  copia.id = uid();
  copia.os = proximoNumeroOS();
  copia.data = new Date().toISOString().slice(0, 10);
  STATE.ordens.push(copia);
  await saveState('ordens');
  await atualizarCounterOS(copia.os);
  toast(`OS ${copia.os} duplicada a partir de ${o.os}`, 'ok');
  renderListaOS();
}

/* ========================================================= */
/*               RENDER DA FOLHA PARA IMPRESSÃO              */
/* ========================================================= */
function ordenarComponentesPorFase(comps, o) {
  const fases = (o?.fases || []).slice().sort((a,b) => (a.ordem||0) - (b.ordem||0));

  // Sem fases (OS sem grade): usa ordem canônica
  if (!fases.length) {
    const canon = (c) => {
      const material = c.material || '';
      if (!material.startsWith('T:')) return 90;
      const tec = STATE.tecidos.find(t => t.id === material.slice(2));
      if (!tec) return 91;
      const cat = categoriaEfetivaTecido(tec);
      if (cat === 'moletom') return 0;
      if (cat === 'malha') return 1;
      if (cat === 'ribana') {
        const n = (c.nome || '').toLowerCase();
        if (n.includes('punho')) return 2;
        if (n.includes('barra')) return 3;
        return 4;
      }
      return 50;
    };
    return [...comps].map((c,i)=>({c,i,p:canon(c)})).sort((a,b)=>a.p-b.p||a.i-b.i).map(x=>x.c);
  }

  // Determina a posição de cada fase pelo índice no array ordenado
  const temMoletomGrade = fases.some(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId)) === 'moletom');
  const posPorTecidoId = new Map();
  const posPorCategoria = new Map();
  const fasesRibana = []; // {pos, label}
  let contRib = 0;
  fases.forEach((f, pos) => {
    const tec = STATE.tecidos.find(t => t.id === f.tecidoId);
    if (!tec) return;
    const cat = categoriaEfetivaTecido(tec);
    if (!posPorTecidoId.has(f.tecidoId)) posPorTecidoId.set(f.tecidoId, pos);
    if (cat && !posPorCategoria.has(cat)) posPorCategoria.set(cat, pos);
    if (cat === 'ribana') {
      contRib++;
      const autoLbl = contRib === 1 ? 'Punhos' : contRib === 2 ? 'Barra' : `Ribana ${contRib}`;
      fasesRibana.push({ pos, label: (f.nome && f.nome.trim()) || autoLbl });
    }
  });

  const prioridade = (c) => {
    const material = c.material || '';
    if (!material.startsWith('T:')) return 100;
    const tecId = material.slice(2);
    const tec = STATE.tecidos.find(t => t.id === tecId);
    if (!tec) return 101;
    const cat = categoriaEfetivaTecido(tec);

    // Ribana: desempata entre fases ribana pelo nome do componente
    if (cat === 'ribana' && fasesRibana.length) {
      const nome = (c.nome || '').toLowerCase();
      for (const r of fasesRibana) {
        const key = (r.label || '').toLowerCase().split(/\s+/)[0].replace(/s$/, '');
        if (key && nome.includes(key)) return r.pos;
      }
      return fasesRibana[0].pos;
    }
    // Moletom: primeira fase moletom
    if (cat === 'moletom' && posPorCategoria.has('moletom')) return posPorCategoria.get('moletom');
    // Forro de capuz (malha com moletom na grade)
    if (cat === 'malha' && temMoletomGrade && posPorCategoria.has('malha')) return posPorCategoria.get('malha');
    // Match por tecidoId exato
    if (posPorTecidoId.has(tecId)) return posPorTecidoId.get(tecId);
    // Match por categoria
    if (posPorCategoria.has(cat)) return posPorCategoria.get(cat);
    return 100;
  };

  return [...comps].map((c, i) => ({ c, i, p: prioridade(c) }))
    .sort((a, b) => a.p - b.p || a.i - b.i)
    .map(x => x.c);
}

function renderComponentesDetalheBox(o) {
  const comps = ordenarComponentesPorFase(o.componentes || [], o);
  if (!comps.length) return '';
  // Quais tamanhos mostrar? Só os que têm peças > 0 em alguma linha (ou que estão na grade)
  const tamanhos = ['p','m','g','gg','g1','g2','g3'];
  const grade = o.grade || {};
  const tamanhosUsados = tamanhos.filter(t => (grade[t] || 0) > 0);
  const colsTam = tamanhosUsados.length ? tamanhosUsados : ['p','m','g','gg']; // default
  const fmt = n => Number(n || 0).toLocaleString('pt-BR');

  let totalGeral = 0;
  const linhas = comps.map(c => {
    const totalLinha = c.qtdTotal || 0;
    totalGeral += totalLinha;
    return `<tr>
      <td><strong>${esc(c.nome || '—')}</strong></td>
      <td>${esc((c.materialNome || '').replace(/^—\s*/,'')) || '—'}</td>
      <td>${esc(c.corNome || '') || '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fmt(c.qtdPorPeca)}</td>
      ${colsTam.map(t => `<td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fmt(c.qtdPorTamanho?.[t])}</td>`).join('')}
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;background:#fff59d;">${fmt(totalLinha)}</td>
    </tr>`;
  }).join('');

  // no-print: o bloco fica na tela (os numeros alimentam contabilidade e
  // faturamento do custo da OS) e nao sai no papel nem no PDF. O dado em si
  // vive em o.componentes, salvo na OS — esconder a tabela nao apaga nada.
  return `
    <table class="side-table no-print" style="border-top:none;width:100%;">
      <thead>
        <tr><th colspan="${4 + colsTam.length + 1}" class="subhead" style="background:#c9e8d0;">Componentes — totais por tamanho <span style="font-weight:400;font-size:6.5pt;color:#555;text-transform:none;letter-spacing:0;">(só na tela — não sai na impressão)</span></th></tr>
        <tr>
          <th>Componente</th>
          <th>Tecido / Material</th>
          <th>Cor</th>
          <th style="width:36px;">/pç</th>
          ${colsTam.map(t => `<th style="width:36px;">${t.toUpperCase()}</th>`).join('')}
          <th style="width:48px;background:#fff59d;">Total</th>
        </tr>
      </thead>
      <tbody>
        ${linhas}
        <tr style="background:#c9e8d0;font-weight:700;">
          <td colspan="${3 + colsTam.length + 1}" style="padding:3px 5px;">TOTAL GERAL COMPONENTES</td>
          <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fmt(totalGeral)}</td>
        </tr>
      </tbody>
    </table>
  `;
}

function renderAviamentosDetalheBox(o) {
  const avs = o.aviamentos || [];
  if (!avs.length) return '';
  const tamanhos = ['p','m','g','gg','g1','g2','g3'];
  const grade = o.grade || {};
  const tamanhosUsados = tamanhos.filter(t => (grade[t] || 0) > 0);
  const colsTam = tamanhosUsados.length ? tamanhosUsados : ['p','m','g','gg'];
  const fmt = n => Number(n || 0).toLocaleString('pt-BR');

  let totalGeral = 0;
  const linhas = avs.map(a => {
    const mat = STATE.materiais.find(m => m.id === a.material);
    const nome = mat ? `${mat.codigo} · ${mat.desc}` : (a.materialNome || '—');
    const totalLinha = a.qtdTotal || 0;
    totalGeral += totalLinha;
    return `<tr>
      <td><strong>${esc(nome)}</strong></td>
      <td>${esc(a.app || '—')}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fmt(a.qtdPorPeca)}</td>
      ${colsTam.map(t => `<td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fmt(a.qtdPorTamanho?.[t])}</td>`).join('')}
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;background:#fff59d;">${fmt(totalLinha)} un</td>
    </tr>`;
  }).join('');

  // no-print pelo mesmo motivo do bloco de Componentes acima: fica na tela pra
  // contabilidade/custo da OS e sai do papel e do PDF. Os dados continuam em
  // o.aviamentos, salvos na OS.
  return `
    <table class="side-table no-print" style="border-top:none;width:100%;">
      <thead>
        <tr><th colspan="${3 + colsTam.length + 1}" class="subhead" style="background:#ffe0b2;">Aviamentos — totais por tamanho <span style="font-weight:400;font-size:6.5pt;color:#555;text-transform:none;letter-spacing:0;">(só na tela — não sai na impressão)</span></th></tr>
        <tr>
          <th>Aviamento</th>
          <th>Aplicação</th>
          <th style="width:36px;">/pç</th>
          ${colsTam.map(t => `<th style="width:36px;">${t.toUpperCase()}</th>`).join('')}
          <th style="width:60px;background:#fff59d;">Total</th>
        </tr>
      </thead>
      <tbody>
        ${linhas}
      </tbody>
    </table>
  `;
}

// Peso/gramatura (g/m²) de um tecido cadastrado, buscado pelo NOME (as fases
// guardam tecido por nome). Retorna 0 se não cadastrado ou sem peso.
function gramaturaTecidoPorNome(nome) {
  if (!nome) return 0;
  const alvo = _normNome(nome);
  const t = (STATE.tecidos || []).find(x => _normNome(x.nome) === alvo);
  return t ? (parseFloat(t.peso) || 0) : 0;
}

// Peso/gramatura (g/m²) de uma COR cadastrada, buscada pelo NOME. A gramatura
// passou a ser cadastrada por cor (varia conforme a cor); tem prioridade sobre
// a do tecido. Retorna 0 se não cadastrada ou sem peso (aí cai no tecido).
function gramaturaCorPorNome(nome) {
  if (!nome) return 0;
  const alvo = _normNome(nome);
  const c = (STATE.cores || []).find(x => _normNome(x.nome) === alvo);
  return c ? (parseFloat(c.peso) || 0) : 0;
}

// Resolve cada fase do enfesto de uma OS e calcula o consumo em kg.
// Fórmula (confirmada): kg = comprimento(m) × largura(m) × camadas × peso(g/m²) / 1000.
// É a fonte única usada tanto na folha de impressão (coluna Consumo) quanto
// na baixa automática de estoque. Espelha exatamente a resolução de comp/larg/
// camadas/tecido que a impressão usa, para que os números batam.
function consumoEnfestoOS(o) {
  const e = o.enfesto || {};
  const tecs = o.tecidos || [];
  const blocos = Array.isArray(e.blocos) && e.blocos.length
    ? e.blocos
    : (e.comprimento || e.largura ? [{ ordem: 1, comp: e.comprimento, larg: e.largura }] : []);
  const camadasGlobal = e.camadas || 0;
  const fasesPorOrdem = {};
  (o.fases || []).forEach(f => { if (f?.ordem) fasesPorOrdem[f.ordem] = f; });
  const linhas = blocos.length
    ? blocos.map((b, i) => ({ b, i }))
    : tecs.map((t, i) => ({ b: { ordem: i + 1, nomeTecido: t.tecidoNome, nomeCor: t.corNome }, i }));
  return linhas.map(({ b, i }) => {
    const ord = b.ordem || (i + 1);
    const fase = fasesPorOrdem[ord] || {};
    let nomeEnf = b.nomeTecido || fase.tecidoNome || '';
    let cor = b.nomeCor || fase.corNome || '';
    if (!cor && nomeEnf.includes(' · ')) {
      const parts = nomeEnf.split(' · ');
      nomeEnf = parts[0];
      cor = parts.slice(1).join(' · ');
    }
    const tecidoReal = fase.tecidoNome || tecs[i]?.tecidoNome || '';
    // OSs salvas ANTES do rename das cores gravaram a cor pura ("Preto"), que
    // não existe mais no cadastro — sem canonicalizar, gramaturaCorPorNome falha,
    // o kg dessas OSs zera ao reimprimir/re-salvar e a chave tecido||cor do
    // estoque diverge das OSs novas. Resolve "Preto"+"Ribana Moletom" para
    // "Preto Ribana Moletom"; se não achar cadastro que case, devolve como veio.
    const corReal = corCanonicaPorTecido(cor || tecs[i]?.corNome || '', tecidoReal);
    const ehVies = /vi[eé]s/i.test(fase.nome || '') || /vi[eé]s/i.test(b.nomeTecido || '') || /vi[eé]s/i.test(nomeEnf);
    const camadas = ehVies ? 1 : (b.camadas || camadasGlobal || 0);
    const comp = (parseFloat(fase.comp) > 0 ? parseFloat(fase.comp) : parseFloat(b.comp)) || 0;
    const larg = (parseFloat(fase.larg) > 0 ? parseFloat(fase.larg) : parseFloat(b.larg)) || 0;
    // Gramatura: prioridade para a COR (varia conforme a cor); se a cor não
    // tem peso cadastrado, cai no peso do TECIDO (compatibilidade). Por fim
    // tenta pelo nome do enfesto.
    const peso = gramaturaCorPorNome(corReal)
      || gramaturaTecidoPorNome(tecidoReal)
      || gramaturaTecidoPorNome(nomeEnf);
    const kg = (comp * larg * camadas * peso) / 1000;
    return { ordem: ord, nomeEnf, tecidoReal, corReal, comp, larg, camadas, peso, kg, ehVies };
  });
}

// Consumo agregado por (tecido, cor) de uma OS — usado na baixa de estoque.
// Soma os kg de todas as fases que usam o mesmo tecido+cor; ignora fases sem kg.
function consumoAgregadoPorTecidoCor(o) {
  const mapa = new Map();
  consumoEnfestoOS(o).forEach(L => {
    if (!(L.kg > 0)) return;
    const tecidoNome = L.tecidoReal || L.nomeEnf || '';
    const corNome = L.corReal || '';
    const k = _normNome(tecidoNome) + '||' + _normNome(corNome);
    const cur = mapa.get(k) || { tecidoNome, corNome, kg: 0 };
    cur.kg += L.kg;
    mapa.set(k, cur);
  });
  return Array.from(mapa.values());
}

// Reserva de estoque ao salvar a OS. Ao gerar a OS o material fica como
// RESERVADO (comprometido, mas ainda em estoque). A baixa definitiva (saída)
// só acontece quando o usuário aponta a OS como produzida (darBaixaMaterialOS).
// Idempotente por osId: remove os movimentos anteriores desta OS e recria do
// consumo atual — preservando o status 'consumido' se a OS já tinha sido baixada.
async function aplicarBaixaEstoqueOS(data) {
  if (!data || !data.id) return;
  if (!Array.isArray(STATE.estoqueMov)) STATE.estoqueMov = [];
  // Se a OS já estava baixada (produzida), mantém o status ao recalcular.
  const jaConsumida = STATE.estoqueMov.some(
    m => m.origem === 'os' && m.osId === data.id && m.status === 'consumido');
  const status = jaConsumida ? 'consumido' : 'reservado';
  const antes = STATE.estoqueMov.length;
  STATE.estoqueMov = STATE.estoqueMov.filter(m => !(m.origem === 'os' && m.osId === data.id));
  const itens = consumoAgregadoPorTecidoCor(data);
  const hoje = new Date().toISOString().slice(0, 10);
  itens.forEach(it => {
    STATE.estoqueMov.push({
      id: uid(),
      tipo: 'saida',
      tecidoNome: it.tecidoNome,
      corNome: it.corNome,
      kg: Math.round(it.kg * 1000) / 1000,
      data: hoje,
      origem: 'os',
      osId: data.id,
      osNumero: data.os || '',
      status,
      consumidoEm: jaConsumida ? hoje : '',
      obs: ''
    });
  });
  if (STATE.estoqueMov.length !== antes || itens.length) {
    try { await saveState('estoqueMov'); } catch (e) { console.warn('reserva estoque', e); }
  }
}

// Aponta a OS como produzida → converte a RESERVA em SAÍDA definitiva (baixa real).
async function darBaixaMaterialOS(osId) {
  if (!exigirAdmin('dar baixa de material')) return;
  const hoje = new Date().toISOString().slice(0, 10);
  let mudou = false;
  (STATE.estoqueMov || []).forEach(m => {
    if (m.origem === 'os' && m.osId === osId && m.status !== 'consumido') {
      m.status = 'consumido'; m.consumidoEm = hoje; mudou = true;
    }
  });
  if (!mudou) return;
  try { await saveState('estoqueMov'); } catch (e) { console.warn('baixa material', e); }
  toast('Baixa de material registrada', 'ok');
  renderEstoque();
}

// Desfaz a baixa: volta a OS para RESERVADO.
async function estornarBaixaMaterialOS(osId) {
  if (!exigirAdmin('estornar baixa de material')) return;
  let mudou = false;
  (STATE.estoqueMov || []).forEach(m => {
    if (m.origem === 'os' && m.osId === osId && m.status === 'consumido') {
      m.status = 'reservado'; m.consumidoEm = ''; mudou = true;
    }
  });
  if (!mudou) return;
  try { await saveState('estoqueMov'); } catch (e) { console.warn('estorno baixa', e); }
  toast('Baixa estornada — voltou para reservado', 'ok');
  renderEstoque();
}

// Estorna (remove) as saídas automáticas de uma OS — usado ao excluir a OS.
async function estornarBaixaEstoqueOS(osId) {
  if (!Array.isArray(STATE.estoqueMov)) return;
  const antes = STATE.estoqueMov.length;
  STATE.estoqueMov = STATE.estoqueMov.filter(m => !(m.origem === 'os' && m.osId === osId));
  if (STATE.estoqueMov.length !== antes) {
    try { await saveState('estoqueMov'); } catch (e) { console.warn('estorno estoque', e); }
  }
}

function renderEnfestoBox(o) {
  const e = o.enfesto || {};
  const tecs = o.tecidos || [];
  // Blocos: usa e.blocos (novo) ou reconstrói um bloco único a partir dos campos legados
  const blocos = Array.isArray(e.blocos) && e.blocos.length
    ? e.blocos
    : (e.comprimento || e.largura ? [{ ordem: 1, comp: e.comprimento, larg: e.largura }] : []);
  // Renderiza se houver enfesto OU tecidos — campo unico mescla as duas infos
  const temAlgo = blocos.length || e.camadas || tecs.length;
  if (!temAlgo) return '';

  const fmt = n => n ? Number(n).toFixed(2).replace('.',',') : '—';
  const fmtKg = n => Number(n).toFixed(3).replace('.',',');
  const fmtBob = n => {
    if (n === 0) return '0';
    if (Number.isInteger(n)) return String(n);
    if (Math.abs(n - 0.5) < 1e-9) return '½';
    return Number(n).toFixed(2).replace(/0+$/, '').replace(/[.]$/, '').replace('.', ',');
  };
  // Previsão de consumo (bobinas) por fase — vem do cadastro da grade viva
  // (previsão de demanda). Se a grade não tem previsão, a coluna Consumo segue
  // mostrando o kg calculado do enfesto (comportamento antigo).
  const gradeVivaPrev = o.gradeId ? STATE.grades.find(g => g.id === o.gradeId) : null;
  const bobPorOrdem = {};
  let gradeTemPrevisao = false;
  if (gradeVivaPrev && Array.isArray(gradeVivaPrev.fases)) {
    gradeVivaPrev.fases.forEach(f => {
      const b = parseBobinas(f.bobinas);
      if (b != null) { bobPorOrdem[f.ordem] = b; if (b > 0) gradeTemPrevisao = true; }
    });
  }

  // Consumo por fase (fonte única — mesma usada na baixa de estoque)
  const consumo = consumoEnfestoOS(o);
  // Sem linha de total: cada fase é um enfesto de um tecido e uma cor próprios,
  // então somar as fases produz um número que não serve pra comprar nem separar
  // material. O consumo fica só onde tem significado — na linha da própria fase,
  // com as bobinas previstas em cima e a estimativa em kg embaixo.

  const enfestosCheck = (o.progresso && o.progresso.enfestosCheck) || {};
  const enfestosTempos = (o.progresso && o.progresso.enfestosTempos) || {};
  // As tonalidades podem aparecer em qualquer fase, então cada fase SEMPRE
  // ganha campos em branco pros três tons (Tom 1/2/3), independente do que está
  // marcado no "Total por tamanho" — preenchíveis à mão e persistidos por fase.
  const enfestosTons = (o.progresso && o.progresso.enfestosTons) || {};
  const tomsSelEnf = [1, 2, 3];
  const campoTom = (ord, tom, val) =>
    `<input type="text" value="${esc(val || '')}" `
    + `data-enf-tom="${esc(String(ord))}" data-enf-tomnum="${tom}" `
    + `onchange="salvarTomEnfesto('${esc(o.id)}', '${esc(String(ord))}', '${tom}', this.value)" `
    + `style="width:48px;border:none;border-bottom:1px solid #888;background:transparent;text-align:center;`
    + `font-family:'IBM Plex Mono',monospace;font-size:6.5pt;padding:0 1px;">`;
  const linhaTons = (ord, tv) => tomsSelEnf.length
    ? `<div style="display:flex;align-items:center;gap:6px;padding:1px 0;font-family:'IBM Plex Mono',monospace;font-size:6pt;line-height:1.3;">
        <span style="font-weight:700;min-width:44px;text-transform:uppercase;letter-spacing:.04em;">Tons</span>
        ${tomsSelEnf.map(tom => `<span style="color:#555;">Tom ${tom}</span>${campoTom(ord, tom, tv[tom])}`).join('')}
      </div>`
    : '';
  // Campo de tempo preenchível (Início/Fim) — persiste em progresso.enfestosTempos[ord].
  // Texto livre (não type="time") pra imprimir como linha limpa de preencher à mão
  // e também aceitar digitação na tela. Sincroniza entre usuários via realtime.
  const campoTempo = (ord, campo, val) =>
    `<input type="text" inputmode="numeric" placeholder="--:--" value="${esc(val || '')}" `
    + `data-enf-tempo="${esc(String(ord))}" data-enf-campo="${campo}" `
    + `onchange="this.value=_horaFmt(this.value); salvarTempoEnfesto('${esc(o.id)}', '${esc(String(ord))}', '${campo}', this.value)" `
    + `style="width:44px;border:none;border-bottom:1px solid #888;background:transparent;text-align:center;`
    + `font-family:'IBM Plex Mono',monospace;font-size:6.5pt;padding:0 1px;">`;
  const linhaTempo = (lbl, ord, campoIni, campoFim, t) =>
    `<div style="display:flex;align-items:center;gap:5px;padding:1px 0;font-family:'IBM Plex Mono',monospace;font-size:6pt;line-height:1.3;">
      <span style="font-weight:700;min-width:44px;text-transform:uppercase;letter-spacing:.04em;">${lbl}</span>
      <span style="color:#555;">Início</span>${campoTempo(ord, campoIni, t[campoIni])}
      <span style="color:#555;">Fim</span>${campoTempo(ord, campoFim, t[campoFim])}
    </div>`;
  const linhasEnfestos = consumo.map(L => {
    const ord = L.ordem;
    const camBloco = L.ehVies ? 1 : (L.camadas || 0);
    const compEf = L.comp || '';
    const largEf = L.larg || '';
    const ckEnf = !!enfestosCheck[ord];
    const t = enfestosTempos[ord] || {};
    return `<tr>
      <td style="text-align:center;"><input type="checkbox" class="os-check" ${ckEnf?'checked':''} data-enfesto="${esc(String(ord))}" onchange="togglarChecklistEnfesto('${esc(o.id)}', this.dataset.enfesto, this.checked)" style="margin:0;"></td>
      <td style="text-align:center;font-weight:700;">${ord}</td>
      <td>${esc(L.nomeEnf) || '—'}</td>
      <td>${esc(L.tecidoReal) || '—'}</td>
      <td>${esc(corSemTecido(L.corReal, L.tecidoReal)) || '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${compEf ? fmt(compEf)+' m' : '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${largEf ? fmt(largEf)+' m' : '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;">${camBloco || '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;white-space:nowrap;padding:0;">
        <div style="border-bottom:1px solid #cfcfcf;padding:2px 3px;font-weight:700;" title="Bobinas previstas (cadastro da grade)">${gradeTemPrevisao && bobPorOrdem[ord] != null ? fmtBob(bobPorOrdem[ord]) : '—'}</div>
        <div style="padding:2px 3px;font-weight:400;font-size:6pt;color:#444;" title="Estimativa por gramatura × comprimento">${L.kg > 0 ? fmtKg(L.kg)+' kg' : '—'}</div>
      </td>
    </tr>` + (L.ehVies ? '' : `
    <tr class="enfesto-tempos">
      <td style="background:#f7faf8;"></td>
      <td colspan="8" style="padding:2px 5px;background:#f7faf8;">
        ${linhaTempo('Enfesto', ord, 'enfIni', 'enfFim', t)}
        ${linhaTons(ord, enfestosTons[ord] || {})}
      </td>
    </tr>`);
  }).join('');

  return `
    <table class="side-table tab-tecidos" style="table-layout:fixed;width:100%;">
      <colgroup>
        <col style="width:22px;">
        <col style="width:32px;">
        <col style="width:52px;">
        <col style="width:54px;">
        <col style="width:58px;">
        <col style="width:42px;">
        <col style="width:56px;">
        <col style="width:26px;">
        <col style="width:58px;">
      </colgroup>
      <thead>
        <tr><th colspan="9" class="subhead" style="background:#c9e8d0;">Enfesto${consumo.length>1?'s':''}</th></tr>
        <tr>
          <th style="font-size:6.5pt;white-space:nowrap;">✓</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Fase</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Enfesto</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Tecido</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Cor</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Compr.</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Largura</th>
          <th style="font-size:6.5pt;white-space:nowrap;">CAM</th>
          <th style="font-size:6.5pt;white-space:nowrap;line-height:1.1;">Consumo<div style="font-size:4.8pt;font-weight:400;color:#666;">bobinas<br>estim. (g×c)</div></th>
        </tr>
      </thead>
      <tbody>
        ${linhasEnfestos}
      </tbody>
    </table>
  `;
}

function renderPrintSheet(o) {
  printOsAtual = o;
  // Se a OS aponta para uma grade cadastrada, usa os dados ATUAIS da grade
  // (fases, distribuição de tamanhos e descrição). Só o enfesto local (comp/larg/camadas
  // digitados na OS) permanece salvo. Isso faz alterações posteriores na grade refletirem
  // automaticamente na impressão.
  if (o.gradeId) {
    const gViva = STATE.grades.find(g => g.id === o.gradeId);
    if (gViva) {
      const fasesAtualizadas = Array.isArray(gViva.fases) ? gViva.fases.map(f => ({
        ordem: f.ordem,
        nome: f.nome || '',
        tecidoId: f.tecidoId || '',
        tecidoNome: (STATE.tecidos.find(t => t.id === f.tecidoId) || {}).nome || '',
        corId: f.corId || '',
        corNome: (STATE.cores.find(c => c.id === f.corId) || {}).nome || '',
        comp: f.comp || '',
        larg: f.larg || ''
      })) : [];
      const tamanhos = (gViva.tamanhos || {});
      const gradeAtualizada = {
        ...(o.grade || {}),
        descricao: gViva.nome || o.grade?.descricao || '',
        p: tamanhos.p != null ? tamanhos.p : (o.grade?.p || 0),
        m: tamanhos.m != null ? tamanhos.m : (o.grade?.m || 0),
        g: tamanhos.g != null ? tamanhos.g : (o.grade?.g || 0),
        gg: tamanhos.gg != null ? tamanhos.gg : (o.grade?.gg || 0),
        g1: tamanhos.g1 != null ? tamanhos.g1 : (o.grade?.g1 || 0),
        g2: tamanhos.g2 != null ? tamanhos.g2 : (o.grade?.g2 || 0),
        g3: tamanhos.g3 != null ? tamanhos.g3 : (o.grade?.g3 || 0)
      };
      gradeAtualizada.total = gradeAtualizada.p + gradeAtualizada.m + gradeAtualizada.g
        + gradeAtualizada.gg + gradeAtualizada.g1 + gradeAtualizada.g2 + gradeAtualizada.g3;
      o = { ...o, fases: fasesAtualizadas, grade: gradeAtualizada };
    }
  }
  const desenho = STATE.desenhos.find(d => d.id === o.desenhoId);
  const imgHtml = desenho?.img
    ? `<img src="${desenho.img}" alt="Desenho técnico">`
    : `<div class="no-img">Nenhum desenho técnico selecionado</div>`;

  // Texto informativo da COR do desenho técnico — barra em CAIXA ALTA logo acima
  // do desenho. Junta TODAS as cores usadas nas variantes (Cor 1, Cor 2 e Cor 3),
  // sem repetir, e ORDENADAS pela sequência canônica do desc do desenho — assim um
  // tricolor mostra as três cores na ordem certa (ex.: "VERDE / PRETO / BEGE"),
  // mesmo quando a variante da OS herdou uma ordem trocada dos campos de cor.
  // corNomeCurto tira o tecido ANTES do Set: o banner é a cor da PEÇA, não do
  // rolo. Um tricolor que usa preto na malha e preto na ribana tem duas cores
  // cadastradas distintas, mas o banner deve dizer "PRETO" uma vez só — sem o
  // corte, sairia "PRETO MALHA ALGODÃO / PRETO RIBANA MALHA ALGODÃO" e estouraria
  // a caixa de 324px que o auto-ajuste de fonte abaixo assume.
  const coresDesenho = ordenarCoresNomesPorDesc([...new Set(
    (o.variantes || [])
      .flatMap(v => [v.cor1Nome, v.cor2Nome, v.cor3Nome])
      .filter(c => c && c !== '—')
      .map(corNomeCurto)
  )], desenho);
  const corTexto = coresDesenho.join(' / ').toUpperCase();
  // Fonte auto-ajustada pra caber SEMPRE em uma linha so, inclusive com tres
  // cores (o CSS poe white-space: nowrap, entao encolher aqui e o que impede
  // o texto de vazar da caixa).
  //
  // Duas coisas foram medidas no Chrome com a fonte real (IBM Plex Sans 800,
  // letter-spacing .04em) contra os 324px uteis do banner, varrendo os 298
  // combos possiveis das 12 cores cadastradas (12 de 1 cor + 66 de 2 + 220
  // de 3):
  //
  // - O piso e 8pt, NAO 20pt. O piso de 20 era o bug que quebrava o texto em
  //   duas linhas: com tres cores a conta pede 11-16pt e o Math.max devolvia
  //   20 assim mesmo. Hoje o pior caso ("VERMELHO / OFF-WHITE / MOSTARDA")
  //   sai em 11pt e cabe.
  //
  // - A constante e 214, nao 230. Com 230, quatro combos vazavam por 1-4px,
  //   todos com "MARROM": a conta assume 0.62em por caractere na media, e
  //   MARROM e quase so glifo largo (M, R, O). 220 foi o maior valor sem
  //   nenhum estouro; 214 deixa margem pra cores novas com letras largas.
  //   Se cadastrarem nomes bem mais longos, vale remedir.
  const corFont = corTexto
    ? Math.max(8, Math.min(30, Math.floor(214 / (corTexto.length * 0.62))))
    : 0;
  // Banner é IRMÃO acima da .desenho-area (não filho) — assim a área do desenho e
  // a imagem ficam idênticas ao original e nada pode escondê-las. Estilos INLINE
  // de propósito, pra funcionar mesmo com um styles.css antigo em cache.
  const corBannerHtml = corTexto
    ? `<div class="desenho-cor" style="width:100%;box-sizing:border-box;padding:6px 8px;text-align:center;font-weight:800;text-transform:uppercase;letter-spacing:.04em;line-height:1.05;white-space:nowrap;color:#000;font-size:${corFont}pt;border-bottom:1.5px solid #000;background:#fff;">${esc(corTexto)}</div>`
    : '';

  const g = o.grade || {};
  const vars_ = o.variantes || [];
  const comps = ordenarComponentesPorFase(o.componentes || [], o);

  // Variantes
  let variantesHtml = '';
  for (let i = 0; i < 4; i++) {
    const v = vars_[i];
    variantesHtml += `<tr>
      <td class="var-head" style="text-align:center;width:36px;">VAR ${i+1}</td>
      <td class="cor-cell">${v?.cor1Nome && v.cor1Nome !== '—' ? esc(v.cor1Nome) : '—'}</td>
      <td class="cor-cell">${v?.cor2Nome && v.cor2Nome !== '—' ? esc(v.cor2Nome) : '—'}</td>
      <td class="cor-cell">${v?.cor3Nome && v.cor3Nome !== '—' ? esc(v.cor3Nome) : '—'}</td>
    </tr>`;
  }


  // Lista curta de aviamentos/componentes removida — info já aparece em
  // "Componentes — totais por tamanho" e "Aviamentos — totais por tamanho"

  // Resolve nome atual de equipe pelo ID (pega nome + função vigentes, não o snapshot salvo)
  const nomeEquipeAtual = (id, fallback) => {
    if (!id) return fallback || '—';
    const p = STATE.equipe.find(x => x.id === id);
    if (!p) return fallback || '—';
    // Prioriza função via ID (se existir) — pega o nome atual mesmo se a função foi renomeada
    let funcaoNome = '';
    if (p.funcaoId) {
      const f = STATE.funcoes.find(x => x.id === p.funcaoId);
      if (f) funcaoNome = f.nome || '';
    }
    if (!funcaoNome && p.funcao) {
      // Fallback: tenta achar função por nome (match case-insensitive); se achar, usa nome atual
      const f = STATE.funcoes.find(x => (x.nome || '').trim().toLowerCase() === (p.funcao || '').trim().toLowerCase());
      funcaoNome = f?.nome || p.funcao || '';
    }
    return p.nome + (funcaoNome ? ' ('+funcaoNome+')' : '');
  };
  const nomeDesigner = nomeEquipeAtual(o.designerId, o.designerNome || o.designer);
  const nomeFtec = nomeEquipeAtual(o.ftecId, o.ftecNome || o.ftec);
  const nomeCoord = nomeEquipeAtual(o.coordenadoId, o.coordenadoNome || o.coordenado);

  // Label do bloco "Coordenador" — usa o nome ATUAL da função (cadastro)
  // do equipe vinculado ao coordenado da OS. Sem coordenado, busca uma função
  // cujo nome contenha "coordenador". Fallback: rótulo padrão "Coordenador".
  const pCoord = o.coordenadoId ? STATE.equipe.find(x => x.id === o.coordenadoId) : null;
  let labelCoord = '';
  if (pCoord?.funcaoId) {
    const f = STATE.funcoes.find(x => x.id === pCoord.funcaoId);
    if (f?.nome) labelCoord = f.nome;
  }
  if (!labelCoord && pCoord?.funcao) labelCoord = pCoord.funcao;
  if (!labelCoord) {
    const f = STATE.funcoes.find(x => /coordenador/i.test(x.nome || ''));
    if (f?.nome) labelCoord = f.nome;
  }
  if (!labelCoord) labelCoord = 'Coordenador';
  // Valor: só o nome da pessoa (a função já está no label)
  const nomeCoordPessoa = pCoord?.nome || o.coordenadoNome || o.coordenado || '—';

  // SKU(s) do produto acabado para o cabeçalho.
  const skuStr = skusDaOS(o).join(' / ') || '—';

  // Blusa moletom TRICOLOR é a única grade que não cabe em 297mm: são 6 fases de
  // enfesto contra 2 de uma camiseta, e o excedente fazia a folha ser reduzida
  // por inteiro na hora de virar PDF — encolhendo a LARGURA junto e deixando
  // ~23mm de branco de cada lado. A classe liga uma versão mais densa (só
  // espaçamento, nenhum campo a menos) que traz a folha de volta a 297mm.
  // Medido com a OS 0435: 334,7mm -> 297mm. Escopo restrito de propósito, para
  // as camisetas — que já cabem — continuarem exatamente como estão.
  const gradeDaOS = o.gradeId ? (STATE.grades || []).find(g => g.id === o.gradeId) : null;
  const ehMoletomTricolor = !!gradeDaOS
    && gradeDaOS.tipoPeca === 'blusa_moletom'
    && gradeDaOS.variacao === 'tricolor';
  const folhaEl = document.getElementById('print-sheet');
  folhaEl.classList.toggle('sheet-densa', ehMoletomTricolor);

  folhaEl.innerHTML = `
    <!-- CABEÇALHO -->
    <div class="sheet-header">
      <div class="cell brand-cell">${esc(o.griffeNome || o.griffe || 'MARCA')}</div>
      <div class="cell"><span class="mini">Coleção</span>${esc(o.colecaoNome || '—')}</div>
      <div class="cell"><span class="mini">${esc(o.blocoNome || o.bloco || 'R1 BLOCO 1')}</span></div>
      <div class="cell"><span class="mini">Data</span>${esc(formatDate(o.data))}</div>
      <div class="cell des-cell" style="flex-direction:column;align-items:center;justify-content:center;">
        <span class="mini">OS Nº:</span>
        <span style="font-size:13pt;letter-spacing:.05em;">${esc(o.os || '—')}</span>
        <span class="mini" style="margin-top:2px;">SKU</span>
        <span style="font-size:8pt;font-weight:700;font-family:'IBM Plex Mono',monospace;white-space:nowrap;line-height:1.1;text-align:center;">${esc(skuStr)}</span>
      </div>
      <div class="cell adult-cell">${esc((o.linhaNome || o.linha || 'ADULTO').toUpperCase())}</div>
    </div>

    <!-- LINHA SECUNDÁRIA: descrição -->
    <div style="display:grid;grid-template-columns:1fr 1.5fr 1fr 1fr;border:1.5px solid #000;border-top:none;font-size:7.5pt;">
      <div style="padding:3px 6px;border-right:1px solid #000;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;text-transform:uppercase;color:#555;letter-spacing:.05em;">Desenho</strong><br>${esc(o.codigo||'—')}</div>
      <div style="padding:3px 6px;border-right:1px solid #000;background:#fff59d;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;text-transform:uppercase;letter-spacing:.05em;">Descrição</strong><br><span style="font-weight:700;">${esc(o.modeloNome||'—')}</span></div>
      <div style="padding:3px 6px;border-right:1px solid #000;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;text-transform:uppercase;color:#555;letter-spacing:.05em;">Base</strong><br>${esc(o.baseNome || o.base || '—')}</div>
      <div style="padding:3px 6px;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;color:#555;letter-spacing:.05em;">${esc(labelCoord)}</strong><br><span style="background:#a7f3d0;padding:1px 4px;">${esc(nomeCoordPessoa)}</span></div>
    </div>

    <!-- LINHA TERCIÁRIA: designer + ficha técnica -->
    <div style="display:grid;grid-template-columns:1fr 1fr;border:1.5px solid #000;border-top:none;font-size:7.5pt;">
      <div style="padding:3px 6px;border-right:1px solid #000;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;text-transform:uppercase;color:#555;letter-spacing:.05em;">Designer</strong><br>${esc(nomeDesigner)}</div>
      <div style="padding:3px 6px;"><strong style="font-family:'IBM Plex Mono',monospace;font-size:7pt;text-transform:uppercase;color:#555;letter-spacing:.05em;">Ficha Técnica</strong><br>${esc(nomeFtec)}</div>
    </div>

    <!-- CORPO -->
    <div class="sheet-body">
      <div class="sheet-left">
        ${corBannerHtml}
        <div class="desenho-area">
          <div class="desenho-label">Desenho Técnico: ${esc(o.codigo || '—')}</div>
          ${imgHtml}
        </div>
        <!-- OBSERVAÇÕES — fica aqui, embaixo do desenho, porque e esta coluna
             que sobra espaco: o desenho tem a altura travada pela largura
             (~137mm) e nao cresce mais que isso. Antes ficava na coluna
             direita e a sobra daqui era so faixa branca.
             flex:1 + 20mm de piso: a caixa estica e vira area de escrita a
             mao. Ja foi flex:none com 20mm fixos, quando Componentes e
             Aviamentos ainda imprimiam e a folga era ~11mm; com os dois fora
             da folha impressa a folga virou 51.2mm e, sem ninguem pra usar,
             ela ia toda pra .desenho-area e virava 31.1mm de branco acima do
             desenho (medido em 8 OS reais).
             Os valores vao INLINE de proposito (pra sobreviver a um
             styles.css antigo em cache, ver o banner de cor) — mas atencao:
             inline vence a folha de estilo, entao mudar so o styles.css nao
             tem efeito nenhum aqui. Mexeu num, mexe no outro. -->
        <div style="background:#c9e8d0;padding:3px 6px;font-family:'IBM Plex Mono',monospace;font-size:7pt;font-weight:700;letter-spacing:.08em;text-transform:uppercase;text-align:center;border:1px solid #000;border-left:none;border-right:none;">Observações</div>
        <div class="obs-box" style="flex:1;min-height:20mm;display:flex;flex-direction:column;border-left:none;border-right:none;"><textarea class="obs-input" placeholder="Digite as observações..." style="flex:1;min-height:14mm;" onchange="salvarObsOS('${esc(o.id)}', this.value)">${esc(o.obs || '')}</textarea></div>
      </div>

      <div class="sheet-right">
        <!-- GRADE -->
        <table class="side-table tab-tecidos" style="table-layout:fixed;width:100%;">
          <!-- A 1ª coluna só carrega os rótulos "Tom 1/2/3" e ficava vazia nas
               linhas de grade e de totais, ocupando ~96px enquanto os tamanhos se
               espremiam em ~31px e o Total sobrava estreito. Larguras explícitas
               devolvem esse espaço: rótulo no tamanho do texto, tamanhos iguais
               entre si e Total com folga para "504" e para o cabeçalho. -->
          <colgroup>
            <col style="width:48px;">
            <col style="width:40px;"><col style="width:40px;"><col style="width:40px;"><col style="width:40px;">
            <col style="width:40px;"><col style="width:40px;"><col style="width:40px;">
            <col style="width:78px;">
          </colgroup>
          <thead>
            <tr><th colspan="9" class="subhead">Grade ${o.grade?.descricao?'· '+esc(o.grade.descricao):''}</th></tr>
            <tr>
              <th></th>
              <th>P</th><th>M</th><th>G</th><th>GG</th><th>G1</th><th>G2</th><th>G3</th><th>Total</th>
            </tr>
          </thead>
          <tbody>
            <tr style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:600;">
              <td></td>
              <td>${g.p>0?g.p:''}</td><td>${g.m>0?g.m:''}</td><td>${g.g>0?g.g:''}</td>
              <td>${g.gg>0?g.gg:''}</td><td>${g.g1>0?g.g1:''}</td><td>${g.g2>0?g.g2:''}</td><td>${g.g3>0?g.g3:''}</td>
              <td style="background:#fff59d;">${g.total>0?g.total:''}</td>
            </tr>
            ${(() => {
              const cam = o.enfesto?.camadas || 0;
              // Todos os números vêm de totaisPorTamanhoTomOS — a mesma função
              // que a folha de OE usa, pra as duas folhas não divergirem.
              const TT = totaisPorTamanhoTomOS(o);
              const multPrincipal = multiplicadorPecaOS(o);
              const t = (q) => (q > 0 && cam > 0) ? q * cam * multPrincipal : '';
              const totalGeral = TT.totalGeral;
              const ttTons = (o.progresso && o.progresso.totalTamanhoTons) || {};
              const sizeKeys = TT.keys;
              // Tons marcados em ordem (prefixo: 1, 1+2 ou 1+2+3). O ultimo
              // vira o "balanceador": cada celula dele recebe colTotal menos a
              // soma dos V dos editaveis, mantendo as somas das colunas iguais
              // a linha "Total geral" e a soma total = X.
              const tomsSel = TT.tons;
              const balancerTom = TT.balancer;
              // V uniforme por linha (mesmo numero em todas as celulas
              // visiveis). Digitar em uma celula propaga pra todas via DOM.
              const vTom = TT.vTom;
              const balancerCellVal = (k) => Math.max(0, TT.colTotal(k) - TT.somaEditaveis);
              const tomRow = (tom) => {
                const isChecked = tomsSel.includes(tom);
                const ck = isChecked ? 'checked' : '';
                let bloqueado = false;
                if (tom === 2 && !ttTons[1]) bloqueado = true;
                if (tom === 3 && (!ttTons[1] || !ttTons[2])) bloqueado = true;
                const disabledAttr = bloqueado ? 'disabled' : '';
                // Todos os tons com a MESMA cor do Tom 1 (sem desbotar): mesmo
                // bloqueado, o texto fica na cor normal — só o checkbox continua
                // desabilitado pra manter a sequência (Tom 2 exige Tom 1 etc.).
                const labelStyle = "display:flex;align-items:center;gap:4px;font-family:'IBM Plex Mono',monospace;font-size:7pt;font-weight:700;";
                // A linha sai de TT.linhas — a mesma estrutura que a folha de OE
                // usa. O Tom 1 aparece mesmo sem checkbox marcado (tonalidade
                // implícita) e, enquanto nada foi digitado, carrega a quantidade
                // cheia; nesse estado ele é só leitura, como o balanceador.
                const linhaTT = TT.linhas.find(L => L.tom === tom);
                const mostra = !!linhaTT;
                // O Tom 1 continua DIGITÁVEL no estado inicial — é digitar nele
                // que reparte a diferença pro balanceador. Só o balanceador e o
                // Tom 1 implícito (sem checkbox) são célula calculada.
                const editavel = !!linhaTT && linhaTT.editavel;
                let rowSum = 0;
                const cells = sizeKeys.map(k => {
                  const has = (g[k] || 0) > 0;
                  if (!mostra || !has) return `<td></td>`;
                  const val = linhaTT.cels[k] || 0;
                  rowSum += val;
                  if (!editavel) {
                    // Célula calculada (balanceador). Precisa da MESMA aparência das
                    // digitáveis — mono, negrito, 8pt, centralizada: sem isso ela
                    // herdava o estilo base da tabela e o número automático saía
                    // apagado e à esquerda, parecendo menos válido que os digitados.
                    return `<td data-tt-balancer-cell="${tom}" data-tt-balancer-tom="${tom}" data-tt-balancer-size="${k}" style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:8pt;">${val > 0 ? val : ''}</td>`;
                  }
                  return `<td style="padding:0;"><input type="number" min="0" value="${val > 0 ? val : ''}" data-tt-tom-input="${tom}" oninput="propagarValorTomTamanho(this, ${tom})" onchange="salvarValorTotalTamanhoTom('${esc(o.id)}', ${tom}, this.value)" style="width:100%;box-sizing:border-box;border:none;background:transparent;text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:8pt;padding:1px 2px;"></td>`;
                }).join('');
                const totalCell = !mostra
                  ? `<td style="background:#c9e8d0;"></td>`
                  : `<td style="background:#c9e8d0;" data-tt-row-total="${tom}">${rowSum > 0 ? rowSum : ''}</td>`;
                return `<tr style="background:#f4faf5;">
                  <td style="white-space:nowrap;padding:1px 4px;">
                    <label style="${labelStyle}">
                      <input type="checkbox" class="os-check" ${ck} ${disabledAttr} data-tt-tom="${tom}" onchange="togglarTotalTamanhoTom('${esc(o.id)}', this.dataset.ttTom, this.checked)" style="margin:0;">
                      Tom ${tom}
                    </label>
                  </td>
                  ${cells}
                  ${totalCell}
                </tr>`;
              };
              return `
                <tr><th colspan="9" class="subhead" style="background:#c9e8d0;font-size:6.5pt;">Total por tamanho</th></tr>
                <tr style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;background:#eaf6ed;">
                  <td></td>
                  <td>${t(g.p)}</td><td>${t(g.m)}</td><td>${t(g.g)}</td>
                  <td>${t(g.gg)}</td><td>${t(g.g1)}</td><td>${t(g.g2)}</td><td>${t(g.g3)}</td>
                  <td style="background:#c9e8d0;">${totalGeral > 0 ? totalGeral : ''}</td>
                </tr>
                ${tomRow(1)}${tomRow(2)}${tomRow(3)}`;
            })()}
          </tbody>
        </table>

        <!-- ENFESTO (mescla campo Base com tecidos/consumo + dados de enfesto) -->
        ${renderEnfestoBox(o)}

        <!-- ETAPAS -->
        <div class="etapas-list">
          <div class="titulo">Etapas de Produção</div>
          ${(() => {
            if (!o.etapas?.length) return `<em style="color:#999;">—</em>`;
            // Tarefas da etapa CORTE saem do enfesto DESTA OS, não do cadastro.
            // O cadastro de etapas é global (tinha "Fase 1|Fase 2|Fase 3" fixo),
            // então não conseguia acompanhar a grade: camiseta com 2 fases
            // imprimia uma "Fase 3" fantasma e moletom com 4 perdia a última.
            // Usando consumoEnfestoOS — a mesma fonte da tabela de Enfestos —
            // o checklist fica 1:1 com as linhas do enfesto, sempre.
            const RE_FASE_TAREFA = /^fase\s*\d+/i;
            // Fase de viés fica de fora: não é enfesto de corte com camadas —
            // é tira cortada em diagonal, não entra no checklist de Corte.
            // Sem o nome do tecido: a coluna do checklist tem ~metade da coluna
            // direita, e "Fase 4 · Malha Algodão" mais os campos de horário
            // quebravam em TRÊS linhas (15,5mm contra 4mm de uma linha), o que
            // sozinho estourava a folha. O tecido de cada fase já está na tabela
            // de Enfestos logo acima.
            const fasesCorte = consumoEnfestoOS(o).filter(L => !L.ehVies).map(L => ({
              nome: 'Fase ' + L.ordem,                       // chave do check: não muda
              hint: '',
              ordem: L.ordem                                 // liga os campos de horário
            }));
            // RESGATE DAS MARCAÇÕES ÓRFÃS. O check de tarefa é gravado pelo NOME
            // (progresso.tarefasCheck[etapa][tarefa]) mas a LISTA exibida vem do
            // cadastro global. Renomear ou excluir uma tarefa no cadastro não
            // apaga nada da OS — só tira da folha a linha que mostrava a marca,
            // e quem preencheu lê isso como "o checklist que eu marquei sumiu".
            // Aqui toda tarefa com marca NESTA OS volta para a lista, mesmo que
            // não exista mais no cadastro, sinalizada como fora dele.
            const marcadasDaEtapa = (nomeEtapa) =>
              Object.entries(((o.progresso || {}).tarefasCheck || {})[nomeEtapa] || {})
                .filter(([, v]) => !!v).map(([t]) => t);
            // Mantém a ordem salva na OS; busca as tarefas embutidas na etapa cadastrada
            const ordenadas = o.etapas.map(nome => {
              const cad = STATE.etapas.find(e => e.nome === nome);
              const cadTarefas = cad ? tarefasDaEtapa(cad).map(t => t.nome) : [];
              let tarefas;
              // Só a etapa de Corte é derivada, e só quando a OS tem enfesto.
              // As tarefas do cadastro que NÃO são "Fase N" (ex.: "Conferir
              // molde") continuam aparecendo, depois das fases.
              if (/corte/i.test(nome) && fasesCorte.length) {
                const extras = cadTarefas.filter(t => !RE_FASE_TAREFA.test(t));
                tarefas = [...fasesCorte, ...extras.map(t => ({ nome: t, hint: '' }))];
              } else {
                tarefas = cadTarefas.map(t => ({ nome: t, hint: '' }));
              }
              const naLista = new Set(tarefas.map(t => t.nome));
              marcadasDaEtapa(nome).forEach(t => {
                if (!naLista.has(t)) tarefas.push({ nome: t, hint: 'fora do cadastro', orfa: true });
              });
              return { nome, tarefas };
            });
            const prog = o.progresso || {};
            const etapaCk = (nomeEtapa) => {
              const checked = !!prog.etapasCheck?.[nomeEtapa];
              return `<input type="checkbox" class="os-check" ${checked?'checked':''}
                onchange="togglarChecklistEtapa('${esc(o.id)}', this.dataset.etapa, this.checked)"
                data-etapa="${esc(nomeEtapa)}">`;
            };
            const tarefaCk = (nomeEtapa, nomeTarefa) => {
              const checked = !!prog.tarefasCheck?.[nomeEtapa]?.[nomeTarefa];
              return `<input type="checkbox" class="os-check sub" ${checked?'checked':''}
                onchange="togglarChecklistTarefa('${esc(o.id)}', this.dataset.etapa, this.dataset.tarefa, this.checked)"
                data-etapa="${esc(nomeEtapa)}" data-tarefa="${esc(nomeTarefa)}">`;
            };
            // Campo Início/Fim de corte, exibido só na etapa "Corte". Par único
            // por OS em progresso.corteTempo. Texto livre (imprime como linha de
            // preencher e aceita digitação na tela).
            const ct = prog.corteTempo || {};
            const campoCorte = (campo) =>
              `<input type="text" inputmode="numeric" placeholder="--:--" value="${esc(ct[campo] || '')}" `
              + `data-corte-tempo="${campo}" `
              + `onchange="this.value=_horaFmt(this.value); salvarTempoCorte('${esc(o.id)}', '${campo}', this.value)" `
              + `style="width:48px;border:none;border-bottom:1px solid #888;background:transparent;text-align:center;`
              + `font-family:'IBM Plex Mono',monospace;font-size:8pt;padding:0 1px;">`;
            const temposCorte = `<span style="display:inline-flex;align-items:center;gap:4px;margin-left:8px;font-family:'IBM Plex Mono',monospace;font-size:7pt;color:#555;font-weight:400;">
                <span>Início</span>${campoCorte('ini')}<span>Fim</span>${campoCorte('fim')}
              </span>`;
            // Início/Fim de corte POR FASE, na própria linha do checklist. Grava
            // em progresso.enfestosTempos[ordem] com chaves PRÓPRIAS (corteIni/
            // corteFim), sem colidir com o Início/Fim de ENFESTO da tabela de
            // enfestos (enfIni/enfFim) — são operações diferentes na mesma fase.
            // Reaproveita salvarTempoEnfesto e o data-enf-tempo, que já é
            // sincronizado entre usuários em atualizarChecksFolha.
            const campoTempoFase = (ordem, campo) => {
              const tv = (prog.enfestosTempos || {})[ordem] || {};
              return `<input type="text" inputmode="numeric" placeholder="--:--" value="${esc(tv[campo] || '')}" `
                + `data-enf-tempo="${esc(String(ordem))}" data-enf-campo="${campo}" `
                + `onchange="this.value=_horaFmt(this.value); salvarTempoEnfesto('${esc(o.id)}', '${esc(String(ordem))}', '${campo}', this.value)" `
                + `style="width:40px;border:none;border-bottom:1px solid #999;background:transparent;text-align:center;`
                + `font-family:'IBM Plex Mono',monospace;font-size:7.5pt;padding:0 1px;">`;
            };
            // flex:none — quem cede largura quando aperta é o texto da fase (que
            // reflui em duas linhas), não os campos de horário, que ficariam
            // espremidos e impossíveis de preencher à mão.
            const tempoFase = (ordem) =>
              `<span style="display:inline-flex;align-items:center;gap:2px;flex:none;margin-left:auto;padding-left:4px;font-family:'IBM Plex Mono',monospace;font-size:6.5pt;color:#666;white-space:nowrap;">
                <span>Ini</span>${campoTempoFase(ordem, 'corteIni')}<span>Fim</span>${campoTempoFase(ordem, 'corteFim')}
              </span>`;
            return `<ul style="list-style:none;padding-left:0;margin:0;font-size:9pt;column-count:2;column-gap:16px;">
              ${ordenadas.map(e => `
                <li style="padding:4px 6px;border-bottom:1px dotted #d4d0c5;break-inside:avoid;-webkit-column-break-inside:avoid;page-break-inside:avoid;">
                  <div style="display:flex;align-items:center;flex-wrap:wrap;">
                    ${etapaCk(e.nome)}
                    <strong>${esc(e.nome)}</strong>${/corte/i.test(e.nome) ? temposCorte : ''}
                  </div>
                  ${e.tarefas.length ? `
                    <!-- font-size vive no styles.css (.etapas-list ul ul): inline
                         venceria a regra da folha densa da blusa moletom tricolor. -->
                    <ul class="tarefas-etapa" style="list-style:none;padding-left:24px;margin:3px 0 0 0;color:#555;">
                      ${e.tarefas.map(t => `
                        <li style="display:flex;align-items:center;padding:1px 0;">
                          ${tarefaCk(e.nome, t.nome)}
                          <span style="flex:1;min-width:0;${t.ordem != null ? 'white-space:nowrap;' : ''}">${esc(t.nome)}${t.hint ? `<span style="color:#8a8a8a;"> · ${esc(t.hint)}</span>` : ''}</span>${t.ordem != null ? tempoFase(t.ordem) : ''}
                        </li>`).join('')}
                    </ul>` : ''}
                </li>`).join('')}
            </ul>`;
          })()}
        </div>

      </div>
    </div>

    <!-- COMPONENTES E AVIAMENTOS — totais por tamanho -->
    ${renderComponentesDetalheBox(o) || ''}
    ${renderAviamentosDetalheBox(o) || ''}

    <div class="sheet-atencao"><strong>Atenção</strong> <span class="atencao-text">${esc(o.atencao || '')}</span></div>
  `;
}

/* ========================================================= */
/*              EXPORT / IMPORT / LIMPAR                     */
/* ========================================================= */
// Lista canônica de todos os arrays persistidos
const ALL_KEYS = ['tecidos','cores','materiais','modelos','colecoes','grades','desenhos',
                  'marcas','linhas','bases','blocos','equipe','funcoes','tarefas','etapas','componentes','ordens'];

function exportarDados() {
  const data = { exportadoEm: new Date().toISOString() };
  // Exporta TUDO (CAD_KEYS), inclusive estoqueMov (estoque de tecido) e os
  // movimentos de fase + osCounter — ALL_KEYS sozinho deixava o estoque de fora.
  CAD_KEYS.forEach(k => { data[k] = STATE[k]; });
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `os-gen-backup-${Date.now()}.json`;
  a.click();
  URL.revokeObjectURL(url);
  toast('Backup exportado', 'ok');
}

async function importarDados(e) {
  if (!exigirAdmin('importar dados')) { e.target.value = ''; return; }
  const file = e.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    ALL_KEYS.forEach(k => {
      if (Array.isArray(data[k])) STATE[k] = data[k];
    });
    for (const k of ALL_KEYS) {
      await saveState(k);
    }
    toast('Dados importados', 'ok');
    goto('home');
  } catch (err) {
    toast('Arquivo inválido', 'err');
  }
}

async function limparTudo() {
  if (!exigirAdmin('apagar todos os dados')) return;
  const resp = prompt(
    'Isso vai APAGAR todos os cadastros e OS de TODOS os usuários da equipe.\n\n' +
    'Esta ação é IRREVERSÍVEL — o último snapshot diário pode não ter tudo.\n\n' +
    'Para confirmar, digite exatamente a palavra APAGAR abaixo:'
  );
  if (resp === null) return;
  if ((resp || '').trim().toUpperCase() !== 'APAGAR') {
    toast('Palavra não conferiu — nada foi apagado.', 'err');
    return;
  }
  // Ação intencional: libera a trava anti-apagamento para este flush.
  _permitirFlushVazio = true;
  try {
    ALL_KEYS.forEach(k => STATE[k] = []);
    for (const k of ALL_KEYS) {
      await saveState(k);
    }
    // Garante que o flush pendente saia antes de rearmar a trava.
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
    await cloudFlush();
  } finally {
    _permitirFlushVazio = false;
  }
  toast('Tudo apagado', 'ok');
  goto('home');
}

/* ========================================================= */
/*              DADOS DE EXEMPLO (Dx7282)                    */
/* ========================================================= */
async function popularExemplo() {
  if (!exigirAdmin('popular dados de exemplo')) return;
  if (!confirm('Isso vai adicionar dados de exemplo aos cadastros. Continuar?')) return;

  // Marcas
  STATE.marcas.push(
    { id: uid(), nome: 'Dixie', desc: 'Marca principal' },
    { id: uid(), nome: 'Diverse', desc: 'Segunda marca' }
  );

  // Linhas
  STATE.linhas.push(
    { id: uid(), nome: 'Adulto', desc: '' },
    { id: uid(), nome: 'Infantil', desc: '' },
    { id: uid(), nome: 'Juvenil', desc: '' },
    { id: uid(), nome: 'Plus Size', desc: '' }
  );

  // Bases
  STATE.bases.push(
    { id: uid(), nome: 'BASE M MOLETOM', desc: 'molde padrão para moletons tam. M' },
    { id: uid(), nome: 'BASE M CAMISETA', desc: 'molde padrão para camisetas tam. M' },
    { id: uid(), nome: 'BASE P CALÇA', desc: 'molde base calça tam. P' }
  );

  // Blocos / revisões
  STATE.blocos.push(
    { id: uid(), nome: 'R1 BLOCO 1', desc: 'primeira revisão, primeiro bloco' },
    { id: uid(), nome: 'R1 BLOCO 2', desc: 'primeira revisão, segundo bloco' },
    { id: uid(), nome: 'R2 BLOCO 1', desc: 'segunda revisão' }
  );

  // Equipe
  STATE.equipe.push(
    { id: uid(), nome: 'Marcelo', funcao: 'Ambos' },
    { id: uid(), nome: 'Ana',     funcao: 'Designer' },
    { id: uid(), nome: 'Paula',   funcao: 'Ficha Técnica' }
  );

  // Tecidos
  STATE.tecidos.push(
    { id: uid(), nome: 'Moletom Bulk', desc: '65% algodão 35% poliéster', categoria: 'moletom' },
    { id: uid(), nome: '1/2 Malha', desc: 'Malha meia-felpa', categoria: 'malha' },
    { id: uid(), nome: 'Ribana Bulk', desc: 'Ribana para punho/barra', categoria: 'malha' },
    { id: uid(), nome: 'Moletom Peluciado', desc: 'Interior peluciado', categoria: 'moletom' },
    { id: uid(), nome: 'Tricoline', desc: 'Algodão fio tinto', categoria: 'outro' }
  );

  // Cores
  STATE.cores.push(
    { id: uid(), nome: 'Camel', hex: '#c9a961', codigo: 'AV.CO.129' },
    { id: uid(), nome: 'Palha', hex: '#e4d9b0', codigo: 'AV.IN.848' },
    { id: uid(), nome: 'Nut', hex: '#6b4423', codigo: 'AV.IL.35' },
    { id: uid(), nome: 'Preto', hex: '#1a1a1a', codigo: '' },
    { id: uid(), nome: 'Off-white', hex: '#f5f2ea', codigo: '' },
    { id: uid(), nome: 'Cinza Mescla', hex: '#9aa0a6', codigo: '' }
  );

  // Materiais
  STATE.materiais.push(
    { id: uid(), codigo: 'AV.IN.848', tipo: 'Cordão', desc: 'Cordão 1,30m palha' },
    { id: uid(), codigo: 'AV.CO.129', tipo: 'Trançador', desc: 'Trançador camel' },
    { id: uid(), codigo: 'AV.IL.35', tipo: 'Ilhós', desc: 'Ilhós metal nut' },
    { id: uid(), codigo: 'AV.EB.182', tipo: 'Etiqueta', desc: 'Etiqueta bordada' },
    { id: uid(), codigo: 'AV.TG.889', tipo: 'Tag', desc: 'Tag papel Dixie' }
  );

  // Modelos
  STATE.modelos.push(
    { id: uid(), nome: 'Moletom fechado básico', linha: 'Adulto' },
    { id: uid(), nome: 'Moletom aberto com zíper', linha: 'Adulto' },
    { id: uid(), nome: 'Calça jogger', linha: 'Adulto' },
    { id: uid(), nome: 'Camiseta regata', linha: 'Adulto' }
  );

  // Coleções
  STATE.colecoes.push(
    { id: uid(), nome: 'Inverno 2024', temporada: 'Outono-Inverno' },
    { id: uid(), nome: 'Verão 2024', temporada: 'Primavera-Verão' },
    { id: uid(), nome: 'Inverno 2025', temporada: 'Outono-Inverno' }
  );

  // Grades
  STATE.grades.push(
    { id: uid(), nome: 'Grade padrão 6 peças', tamanhos: { p:1, m:2, g:2, gg:1, g1:0, g2:0, g3:0 } },
    { id: uid(), nome: 'Grade ampliada 8 peças', tamanhos: { p:2, m:2, g:2, gg:1, g1:1, g2:0, g3:0 } },
    { id: uid(), nome: 'Grade plus 4 peças',     tamanhos: { p:0, m:0, g:0, gg:1, g1:1, g2:1, g3:1 } }
  );

  for (const k of ['tecidos','cores','materiais','modelos','colecoes','grades',
                   'marcas','linhas','bases','blocos','equipe']) {
    await saveState(k);
  }
  toast('Exemplos carregados — cadastre o desenho técnico em "Desenhos" enviando uma imagem', 'ok');
  goto('home');
}

/* ========================================================= */
/*                   INICIALIZAÇÃO                           */
/* ========================================================= */
(async function init() {
  await inicializarAuth();
  if (currentUser) {
    await loadState();
    await carregarPapel();
    aplicarPermissoesUI();
    await migrarEtapasOS();        // padroniza etapas das OSs (1×, admin)
    // Republica o snapshot p/ Contabilidade/Estoque-Confeccao ao ABRIR como admin
    // (reload): aqui o papel já está carregado — no init, loadState roda ANTES de
    // carregarPapel, então o republish do fim do loadState não pega o papel. Sem
    // isto, recarregar a página deixava o snapshot antigo (SKUs vazios) no ar.
    if (currentRole === 'admin' && typeof atualizarContabSnapshot === 'function') {
      atualizarContabSnapshot();
    }
    goto('home');
    // Tarefas em background — não bloqueiam a navegação
    snapshotDiario().catch(e => console.warn('snapshotDiario', e));
    // Snapshot de contingência base ao abrir (estado carregado, não-vazio).
    salvarSnapshotContingencia({ forcar: true }).catch(e => console.warn('snapshot base', e));
    if (currentRole === 'admin') {
      migrarImagensBase64().catch(e => console.warn('migrarImagensBase64', e));
    }
  }
})();

// Deixar disponível globalmente
window.goto = goto;
window.openCadastroModal = openCadastroModal;
window.closeModal = closeModal;
window.salvarCadastro = salvarCadastro;
window.excluirCadastro = excluirCadastro;
window.addTecidoRow = addTecidoRow;
window.renderEnfestoBlocos = renderEnfestoBlocos;
window.addVarianteRow = addVarianteRow;
window.addComponenteRow = addComponenteRow;
window.expandirCoresComponente = expandirCoresComponente;
window.addAviamentoRow = addAviamentoRow;
window.addEtapaCustomizada = addEtapaCustomizada;
window.aplicarGradePreset = aplicarGradePreset;
window.atualizarCalculosEnfesto = atualizarCalculosEnfesto;
window.calcularCamadasParaProducao = calcularCamadasParaProducao;
window.mostrarResponsabilidadesFuncao = mostrarResponsabilidadesFuncao;
window.atualizarCoresComponente = atualizarCoresComponente;
window.addFaseGradeRow = addFaseGradeRow;
window.removerFaseGrade = removerFaseGrade;
window.atualizarResponsabilidadesOS = atualizarResponsabilidadesOS;
window.onModeloChange = onModeloChange;
window.renderEtapasCad = renderEtapasCad;
window.addTarefaEtapaRow = addTarefaEtapaRow;
window.removerTarefaEtapa = removerTarefaEtapa;
window.copiarEtapasEntreDesenhos = copiarEtapasEntreDesenhos;
window.rodarCopiarEtapasParaTodos = rodarCopiarEtapasParaTodos;
window.recarregarDadosDoServidor = recarregarDadosDoServidor;
window.togglarChecklistEtapa = togglarChecklistEtapa;
window.togglarChecklistTarefa = togglarChecklistTarefa;
window.togglarChecklistEnfesto = togglarChecklistEnfesto;
window.salvarTomEnfesto = salvarTomEnfesto;
window.togglarTotalTamanhoTom = togglarTotalTamanhoTom;
window.salvarValorTotalTamanhoTom = salvarValorTotalTamanhoTom;
window.propagarValorTomTamanho = propagarValorTomTamanho;
window.renderComponentesCad = renderComponentesCad;
window.toggleUnidadesGrade = toggleUnidadesGrade;
window.aplicarVinculosDesenho = aplicarVinculosDesenho;
window.aplicarVinculosModelo = aplicarVinculosModelo;
window.previewDesenhoSelecionado = previewDesenhoSelecionado;
window.previewUploadImg = previewUploadImg;
window.reindexTecidos = reindexTecidos;
window.reindexVariantes = reindexVariantes;
window.salvarOS = salvarOS;
window.salvarEImprimir = salvarEImprimir;
window.imprimirEtiquetas = imprimirEtiquetas;
window.imprimirEtiquetasAtual = imprimirEtiquetasAtual;
window.salvarEImprimirEtiquetas = salvarEImprimirEtiquetas;
window.ajustarImpressaoParaA4 = ajustarImpressaoParaA4;
window.conectarPastaPdf = conectarPastaPdf;
window.desconectarPastaPdf = desconectarPastaPdf;
window.conectarPastaOe = conectarPastaOe;
window.desconectarPastaOe = desconectarPastaOe;
window.salvarPdfOeNaPasta = salvarPdfOeNaPasta;
window.conectarPastaBackup = conectarPastaBackup;
window.desconectarPastaBackup = desconectarPastaBackup;
window.escreverBackupJsonAgora = escreverBackupJsonAgora;
window.verOS = verOS;
window.editarOS = editarOS;
window.editarOsAtual = editarOsAtual;
window.excluirOS = excluirOS;
window.duplicarOS = duplicarOS;
window.abrirMovEstoque = abrirMovEstoque;
window.salvarMovEstoque = salvarMovEstoque;
window.excluirMovEstoque = excluirMovEstoque;
window.renderEstoqueCorte = renderEstoqueCorte;
window.renderFasePainel = renderFasePainel;
window.abrirMovFase = abrirMovFase;
window.salvarMovFase = salvarMovFase;
window.excluirMovFase = excluirMovFase;
window.darBaixaMaterialOS = darBaixaMaterialOS;
window.estornarBaixaMaterialOS = estornarBaixaMaterialOS;
window.exportarDados = exportarDados;
window.importarDados = importarDados;
window.limparTudo = limparTudo;
window.popularExemplo = popularExemplo;
window.abrirLogin = abrirLogin;
window.fecharLogin = fecharLogin;
window.trocarAbaAuth = trocarAbaAuth;
window.submeterAuth = submeterAuth;
window.sairConta = sairConta;
window.abrirRecuperacaoSenha = abrirRecuperacaoSenha;
window.enviarEmailRecuperacao = enviarEmailRecuperacao;
window.definirNovaSenha = definirNovaSenha;
window.sincCodigoDesenho = sincCodigoDesenho;
window.atualizarDatalistCodigos = atualizarDatalistCodigos;
window.renderFuncoes = renderFuncoes;
window.listarSnapshots = listarSnapshots;
window.restaurarSnapshot = restaurarSnapshot;
window.listarSnapshotsLocais = listarSnapshotsLocais;
window.restaurarSnapshotLocal = restaurarSnapshotLocal;
window.esconderAlertaSalvamento = esconderAlertaSalvamento;
window.setUserRole = setUserRole;
window.listarUsuariosComPapel = listarUsuariosComPapel;
window.duplicarCadastro = duplicarCadastro;
window.toggleFolderGrade = toggleFolderGrade;
window.moverEtapaForm = moverEtapaForm;
window.moverEtapaDesenho = moverEtapaDesenho;
