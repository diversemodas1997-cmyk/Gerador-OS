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
    // Backup local automatico (silencioso; falha nao bloqueia o save).
    // Funcao definida mais abaixo, perto da pasta de PDFs.
    if (typeof escreverBackupJson === 'function') {
      escreverBackupJson().catch(e => console.warn('backup local', e));
    }
  } catch (e) {
    console.error('cloudFlush', e);
    setSyncStatus('error');
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
        const activeBtn = document.querySelector('.nav-btn.active');
        const pagina = activeBtn && activeBtn.dataset.page ? activeBtn.dataset.page : 'home';
        if (pagina !== 'nova-os') goto(pagina);
        toast('Dados atualizados por outro usuário', 'ok');
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
  pararPolling();
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
      // Evita resetar formulario em edicao — so re-renderiza paginas seguras
      const activeBtn = document.querySelector('.nav-btn.active');
      const pagina = activeBtn?.dataset?.page || 'home';
      if (pagina !== 'nova-os') goto(pagina);
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
  else { el.textContent = '☁ Sincronizado'; }
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
const CAD_KEYS = ['tecidos','cores','materiais','modelos','colecoes','grades','desenhos','marcas','linhas','bases','blocos','equipe','funcoes','tarefas','etapas','componentes','ordens','osCounter'];

async function inicializarAuth() {
  if (!supa) return;
  const { data: { session } } = await supa.auth.getSession();
  if (session && session.user) {
    currentUser = session.user;
    await cloudLoad();
    iniciarRealtime();
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
      iniciarRealtime();
    } else if (event === 'SIGNED_OUT') {
      pararRealtime();
      pararPresenceOS();
      currentUser = null;
      cloudCache = null;
      currentRole = null;
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
  osCounter: 0,
  etapasPadrao: ['Corte', 'Termo (Frente)', 'Costura', 'Travetes (Bolso)', 'Acabamento', 'Estampa', 'Retirada de fios', 'Lavanderia'],
  componentesPadrao: ['Frente', 'Costas', 'Capuz', 'Forro do capuz', 'Mangas', 'Bolso canguru', 'Punho', 'Barra', 'Ribana', 'Cobre gola', 'Recorte lateral', 'Cordão', 'Ilhós', 'Etiqueta interna', 'Tag']
};

async function saveState(key) {
  try {
    await DB.set(key, JSON.stringify(STATE[key]));
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
  const keys = ['tecidos','cores','materiais','modelos','colecoes','grades','desenhos','marcas','linhas','bases','blocos','equipe','funcoes','tarefas','etapas','componentes','ordens'];
  for (const k of keys) {
    try {
      const r = await DB.get(k);
      if (r && r.value) {
        try { STATE[k] = JSON.parse(r.value); } catch { STATE[k] = []; }
      }
    } catch (e) { /* chave não existe ainda, ok */ }
  }
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
}

function uid() { return 'id_' + Date.now() + '_' + Math.floor(Math.random()*1000); }

/* ========================================================= */
/*                      NAVEGAÇÃO                            */
/* ========================================================= */
function goto(page) {
  // Bloqueia navegação a páginas de cadastro para usuários não-admin
  if (page && page.startsWith('cad-') && currentRole && currentRole !== 'admin') {
    toast('Apenas admin pode acessar cadastros', 'err');
    page = 'home';
  }
  document.querySelectorAll('section.page').forEach(s => s.classList.add('hidden'));
  const target = document.querySelector(`section.page[data-page="${page}"]`);
  if (target) target.classList.remove('hidden');
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const btn = document.querySelector(`.nav-btn[data-page="${page}"]`);
  if (btn) btn.classList.add('active');
  window.scrollTo(0, 0);

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
  if (page === 'nova-os') initOSForm();
  if (page === 'config') {
    atualizarPdfFolderStatus();
    atualizarBackupFolderStatus();
  }
}

document.querySelectorAll('.nav-btn').forEach(b => b.addEventListener('click', () => goto(b.dataset.page)));

/* ========================================================= */
/*                      TOAST                                */
/* ========================================================= */
function toast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show ' + type;
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
        <div class="field"><label>Composição / observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: 65% algodão 35% poliéster"></div>
      </div>`;
  }
  else if (tipo === 'cor') {
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Camel"></div>
        <div class="field"><label>Cor (hex)</label><input type="color" id="m-hex" value="${item.hex||'#c9a961'}"></div>
        <div class="field full"><label>Código (ex.: Linx)</label><input type="text" id="m-codigo" value="${esc(item.codigo||'')}" placeholder="Ex.: AV.CO.129"></div>
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
    const tiposPeca = [
      { v: '', lbl: '— sem categoria —' },
      { v: 'camiseta', lbl: 'Camiseta' },
      { v: 'blusa_moletom', lbl: 'Blusa Moletom' },
      { v: 'outro', lbl: 'Outro' }
    ];
    const variacoes = [
      { v: '', lbl: '— sem variação —' },
      { v: 'basica', lbl: 'Básica' },
      { v: 'bicolor', lbl: 'Bicolor' },
      { v: 'tricolor', lbl: 'Tricolor' }
    ];
    box.innerHTML = `
      <div class="form-grid cols-2">
        <div class="field full"><label>Nome *</label><input type="text" id="m-nome" value="${esc(item.nome||'')}" placeholder="Ex.: Grade padrão 6 peças"></div>
        <div class="field"><label>Tipo de peça</label>
          <select id="m-grade-tipopeca">
            ${tiposPeca.map(t => `<option value="${t.v}" ${item.tipoPeca===t.v?'selected':''}>${t.lbl}</option>`).join('')}
          </select>
        </div>
        <div class="field"><label>Variação</label>
          <select id="m-grade-variacao">
            ${variacoes.map(x => `<option value="${x.v}" ${item.variacao===x.v?'selected':''}>${x.lbl}</option>`).join('')}
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
                  const atual = porId.get(c.id) || {};
                  const marcado = porId.has(c.id);
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
}

function addFaseGradeRow(fase = {}) {
  const cont = document.getElementById('m-fases-container');
  if (!cont) return;
  const tecOpts = (selId) => '<option value="">— selecione —</option>' + STATE.tecidos.map(t =>
    `<option value="${esc(t.id)}" ${selId===t.id?'selected':''}>${esc(t.nome)}${t.categoria?' ('+esc(t.categoria)+')':''}</option>`).join('');
  const unidadesAtual = parseInt(fase.unidades) || 2;
  const unidadesOpts = [1, 2, 4, 6, 8].map(n =>
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
      <div class="field fase-unid-wrap"><label>Unidades da grade</label><select class="fase-unid">${unidadesOpts}</select><div class="field-hint">1 unidade da grade = N peças por camada (ribana moletom)</div></div>
      <div class="field"><label>Comprimento (m)</label><input type="number" step="0.01" class="fase-comp" value="${esc(fase.comp || '')}" placeholder="Ex.: 6,50"></div>
      <div class="field"><label>Largura (m)</label><input type="number" step="0.01" class="fase-larg" value="${esc(fase.larg || '')}" placeholder="Ex.: 1,80"></div>
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
  wrap.style.display = isTecidoRibanaMoletom(tec) ? '' : 'none';
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

// Handler do botao de Configuracoes: le o codigo digitado, confirma e dispara a copia.
async function rodarCopiarEtapasParaTodos() {
  const input = document.getElementById('copyEtapasOrigem');
  const codigo = (input?.value || '').trim();
  if (!codigo) { toast('Informe o codigo do desenho de origem', 'err'); return; }
  const origem = STATE.desenhos.find(d => (d.codigo || '').trim() === codigo);
  if (!origem) { toast(`Desenho "${codigo}" nao encontrado`, 'err'); return; }
  const etapas = Array.isArray(origem.etapasNomes) ? origem.etapasNomes : [];
  if (!etapas.length) { toast(`Desenho "${codigo}" nao tem etapas configuradas`, 'err'); return; }
  const total = STATE.desenhos.length - 1;
  const ok = confirm(
    `Copiar as ${etapas.length} etapas do desenho "${codigo}" para os outros ${total} desenhos?\n\n`
    + `Etapas: ${etapas.join(', ')}\n\n`
    + `Apenas o campo "etapas" sera sobrescrito nos demais desenhos. Esta acao nao pode ser desfeita automaticamente.`
  );
  if (!ok) return;
  await copiarEtapasEntreDesenhos(codigo);
}

// Utilitario admin: copia as etapasNomes (e a ordem) de um desenho de origem
// para todos os demais desenhos cadastrados. Uso: copiarEtapasEntreDesenhos('001').
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
  let alteradas = 0;
  STATE.desenhos.forEach(d => {
    if (d.id === origem.id) return;
    d.etapasNomes = [...etapasNomes];
    alteradas++;
  });
  await saveState('desenhos');
  toast(`Etapas de "${codigoOrigem}" aplicadas a ${alteradas} desenhos`, 'ok');
  if (typeof renderDesenhos === 'function') renderDesenhos();
  return { origem: codigoOrigem, etapas: etapasNomes, alteradas };
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
  }
  else if (tipo === 'cor') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.hex = v('m-hex');
    item.codigo = v('m-codigo');
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
    item.fases = Array.from(document.querySelectorAll('#m-fases-container .fase-grade-bloco')).map((b, i) => ({
      ordem: i + 1,
      nome: b.querySelector('.fase-nome')?.value || '',
      tecidoId: b.querySelector('.fase-tec')?.value || '',
      unidades: parseInt(b.querySelector('.fase-unid')?.value) || 2,
      comp: b.querySelector('.fase-comp')?.value || '',
      larg: b.querySelector('.fase-larg')?.value || ''
    }));
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
    item.componentes = Array.from(document.querySelectorAll('.m-componente-chk:checked')).map(chk => {
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
  if (!STATE.tecidos.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhum tecido cadastrado.</td></tr>`; return; }
  const catLabel = { malha: 'Malha algodão · máx 80', moletom: 'Moletom · máx 36', outro: 'Outro' };
  tb.innerHTML = STATE.tecidos.map(t => `
    <tr>
      <td><strong>${esc(t.nome)}</strong></td>
      <td>${esc(t.desc)}</td>
      <td><span class="badge">${esc(catLabel[t.categoria] || '—')}</span></td>
      ${acoesCell('tecido', t.id)}
    </tr>`).join('');
}
function renderCores() {
  const tb = document.getElementById('tbl-cores');
  if (!STATE.cores.length) { tb.innerHTML = `<tr><td colspan="3" class="empty">Nenhuma cor cadastrada.</td></tr>`; return; }
  tb.innerHTML = STATE.cores.map(c => `
    <tr><td><span class="color-swatch" style="background:${esc(c.hex)}"></span><strong>${esc(c.nome)}</strong></td>
    <td><span class="badge">${esc(c.codigo)||'—'}</span></td>${acoesCell('cor', c.id)}</tr>`).join('');
}
function renderMateriais() {
  const tb = document.getElementById('tbl-materiais');
  if (!STATE.materiais.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhum material cadastrado.</td></tr>`; return; }
  tb.innerHTML = STATE.materiais.map(m => `
    <tr><td><span class="badge">${esc(m.codigo)}</span></td><td>${esc(m.desc)}</td>
    <td>${esc(m.tipo)||'—'}</td>${acoesCell('material', m.id)}</tr>`).join('');
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

function renderGrades() {
  const tb = document.getElementById('tbl-grades');
  if (!STATE.grades.length) { tb.innerHTML = `<tr><td colspan="4" class="empty">Nenhuma grade cadastrada.</td></tr>`; return; }

  const labelsTipoPeca = { camiseta: 'Camiseta', blusa_moletom: 'Blusa Moletom', outro: 'Outro', '': 'Sem categoria' };
  const labelsVariacao = { basica: 'Básica', bicolor: 'Bicolor', tricolor: 'Tricolor', '': 'Sem variação' };
  const ordemTipoPeca = ['camiseta', 'blusa_moletom', 'outro', ''];
  const ordemVariacao = ['basica', 'bicolor', 'tricolor', ''];

  // Agrupa por tipoPeca → variacao
  const grupos = {};
  for (const g of STATE.grades) {
    const tp = g.tipoPeca || '';
    const vr = g.variacao || '';
    grupos[tp] = grupos[tp] || {};
    grupos[tp][vr] = grupos[tp][vr] || [];
    grupos[tp][vr].push(g);
  }

  const renderGradeRow = (g) => {
    const t = g.tamanhos || {};
    const dist = ['p','m','g','gg','g1','g2','g3']
      .filter(x => t[x] > 0).map(x => `${x.toUpperCase()}:${t[x]}`).join(' · ');
    const total = Object.values(t).reduce((a,b)=>a+(b||0),0);
    const nFases = Array.isArray(g.fases) ? g.fases.length : 0;
    const fasesBadge = nFases > 0 ? ` <span class="badge" style="background:#fff8e1">${nFases} fase${nFases>1?'s':''}</span>` : '';
    return `<tr><td style="padding-left:48px;"><strong>${esc(g.nome)}</strong>${fasesBadge}</td>
      <td><code style="font-size:11px">${dist||'—'}</code></td>
      <td><span class="badge">${total}</span></td>${acoesCell('grade', g.id)}</tr>`;
  };

  let html = '';
  for (const tp of ordemTipoPeca) {
    if (!grupos[tp]) continue;
    const tpPath = 'tp:' + tp;
    const tpOpen = pastasGradeExpandidas.has(tpPath);
    const chevTop = tpOpen ? '▼' : '▶';
    const totalNoGrupo = Object.values(grupos[tp]).reduce((a, v) => a + v.length, 0);
    html += `<tr class="grade-folder grade-folder-top" onclick="toggleFolderGrade('${tpPath}')"><td colspan="4">
      <span class="folder-chev">${chevTop}</span> 📁 ${esc(labelsTipoPeca[tp] || tp)}
      <span class="folder-count">(${totalNoGrupo})</span>
    </td></tr>`;
    if (!tpOpen) continue;

    for (const vr of ordemVariacao) {
      const gs = grupos[tp][vr];
      if (!gs || !gs.length) continue;
      const vrPath = tpPath + '|var:' + vr;
      const vrOpen = pastasGradeExpandidas.has(vrPath);
      const chevSub = vrOpen ? '▼' : '▶';
      html += `<tr class="grade-folder grade-folder-sub" onclick="event.stopPropagation(); toggleFolderGrade('${vrPath}')"><td colspan="4">
        <span class="folder-chev">${chevSub}</span> ↳ ${esc(labelsVariacao[vr] || vr)}
        <span class="folder-count">(${gs.length})</span>
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
    // Aplica componentes padrão do desenho — nova estrutura com tecido+cor+qtd por componente
    const compsDesenho = Array.isArray(d.componentes) && d.componentes.length
      ? d.componentes
      : (d.componentesIds || []).map(id => ({
          componenteId: id,
          nome: (STATE.componentes.find(x => x.id === id) || {}).nome || '',
          tecidoId: d.tecidoPadraoId || '',
          corId: d.corPrincipalId || '',
          qtdPorPeca: 1
        }));
    if (compsDesenho.length) {
      const cont = document.getElementById('componentes-rows');
      if (cont) {
        cont.innerHTML = '';
        compsDesenho.forEach(c => {
          // Lookup robusto: por ID e, se falhar, por nome
          const cad = STATE.componentes.find(x => x.id === c.componenteId)
                   || (c.nome ? STATE.componentes.find(x => x.nome === c.nome) : null);
          // Prioriza a cor escolhida NO DESENHO pra este componente; cor1Id do cadastro é fallback
          const corPrincipal = c.corId || cad?.cor1Id || '';
          addComponenteRow({
            nome: c.nome || cad?.nome || '',
            material: c.tecidoId ? 'T:' + c.tecidoId : '',
            cor: corPrincipal,
            qtdPorPeca: c.qtdPorPeca != null ? c.qtdPorPeca : 1
          });
        });
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
          if (c1 && d.corPrincipalId)   c1.value = d.corPrincipalId;
          if (c2 && d.corSecundariaId)  c2.value = d.corSecundariaId;
          if (c3 && d.corTerciariaId)   c3.value = d.corTerciariaId;
          aplicou = true;
        }
      }
    }
    // Aplica tecido/cores do desenho nas linhas de Tecidos (1 linha por cor)
    const coresDoDesenho = [d.corPrincipalId, d.corSecundariaId, d.corTerciariaId].filter(Boolean);
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
    bloco.innerHTML = `
      <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;color:var(--ink);margin-bottom:6px;letter-spacing:.08em;">
        ENFESTO ${i+1}${labelDisplay ? ` · <span style="color:var(--ink-2);font-weight:500;">${esc(labelDisplay)}</span>` : ''}
      </div>
      <div class="form-grid cols-3">
        <div class="field"><label>Comprimento (m)</label><input type="number" step="0.01" class="enf-comp" data-idx="${i}" value="${esc(p.comp||'')}" placeholder="Ex.: 6,50"></div>
        <div class="field"><label>Largura (m)</label><input type="number" step="0.01" class="enf-larg" data-idx="${i}" value="${esc(p.larg||'')}" placeholder="Ex.: 1,80"></div>
        <div class="field"><label>Camadas</label><input type="number" min="0" step="1" class="enf-camadas" data-idx="${i}" value="${esc(p.camadas||'')}" placeholder="—" oninput="atualizarCalculosEnfesto()"></div>
      </div>`;
    cont.appendChild(bloco);
  }
}

function lerEnfestoBlocos() {
  const cont = document.getElementById('f-enfestos-blocos');
  if (!cont) return [];
  return Array.from(cont.querySelectorAll('.enfesto-bloco')).map((b, i) => ({
    ordem: i + 1,
    nomeTecido: b.dataset.nomeTecido || '',
    nomeCor: b.dataset.nomeCor || '',
    comp: parseFloat(b.querySelector('.enf-comp').value) || 0,
    larg: parseFloat(b.querySelector('.enf-larg').value) || 0,
    camadas: parseInt(b.querySelector('.enf-camadas')?.value) || 0
  }));
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
      const cores = [
        desenhoAtual?.corPrincipalId,
        desenhoAtual?.corSecundariaId,
        desenhoAtual?.corTerciariaId
      ];
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

// Ribana moletom: tecido cuja categoria efetiva é ribana E o nome contém "moletom".
// Usa o multiplicador "unidades da grade" cadastrado por fase em vez do MULTIPLICADOR_PECAS.ribana.
function isTecidoRibanaMoletom(t) {
  if (!t) return false;
  if (categoriaEfetivaTecido(t) !== 'ribana') return false;
  return (t.nome || '').toLowerCase().includes('moletom');
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
  if (cat === 'moletom') return 'blusa_moletom';
  return cat; // 'camiseta' / 'outro' / ''
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
  if (!cat && !tipoModelo && !variacao) return STATE.grades;
  const keep = new Set(extraIds.filter(Boolean));
  return STATE.grades.filter(g => {
    if (keep.has(g.id)) return true;
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
          // Multiplicador por label de fase ribana: ribana moletom usa "unidades" da fase;
          // demais ribanas usam o multiplicador padrão (2).
          const multPorLabelRib = {};
          papeis.forEach((p, idx) => {
            if (!(p.papel || '').startsWith('ribana_')) return;
            const fase = fasesGrade[idx];
            const tec = STATE.tecidos.find(t => t.id === fase?.tecidoId);
            multPorLabelRib[p.label] = isTecidoRibanaMoletom(tec)
              ? (parseInt(fase?.unidades) || multRib)
              : multRib;
          });
          ribanaPorFase = grupos
            .filter(g => g.qty > 0)
            .map(g => {
              const mult = multPorLabelRib[g.label] || multRib;
              const camadas = gradeTotal > 0 ? Math.ceil(referencia / (gradeTotal * mult)) : 0;
              return {
                label: g.label,
                total: referencia * g.qty,
                detalhes: g.detalhes,
                camadas,
                mult
              };
            });
          // Se tiver componentes sem match, agrupa num fallback "Ribana (outros)"
          if (sobra.length) {
            const qtyTot = sobra.reduce((s, x) => s + x.qty, 0);
            const camadasOutros = gradeTotal > 0 ? Math.ceil(referencia / (gradeTotal * multRib)) : 0;
            ribanaPorFase.push({
              label: 'Ribana (outros)',
              total: referencia * qtyTot,
              detalhes: sobra.map(x => `${x.nome} ×${x.qty}`),
              camadas: camadasOutros,
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
          const refBlusas = totalMoletom || totalForro;
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
                  (1 camada = ${m} peça${m===1?'':'s'}/tamanho; ${refBlusas} peças ÷ ${gradeTotal*m})
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
    let val;
    if (papel.papel === 'moletom') {
      // Enfesto moletom: todos componentes moletom na mesma camada → 1 camada = 1 blusa
      val = camadasPrincipal;
    } else if (papel.papel === 'forro_capuz') {
      // Enfesto forro: camadas = metade das camadas de moletom
      val = Math.max(1, Math.ceil(camadasPrincipal / 2));
    } else if ((papel.papel || '').startsWith('ribana_')) {
      // Enfesto ribana:
      //  - Ribana moletom: regra simples — camadas = camadasMoletom / unidades.
      //    Ignora qtdPorBlusa do componente (a relacao com moletom e direta:
      //    1x=igual, 2x=metade, 4x=um quarto, etc.).
      //  - Ribana padrao: usa qtdPorBlusa do componente, multiplicador fixo (2).
      const fase = fases[i] || {};
      const tecFase = STATE.tecidos.find(t => t.id === fase.tecidoId);
      if (isTecidoRibanaMoletom(tecFase)) {
        const unidades = parseInt(fase.unidades) || multRib;
        val = Math.max(1, Math.ceil(camadasPrincipal / unidades));
      } else {
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
  await aplicarRegraConjugadaSeAplicavel(data);
  toast('OS ' + data.os + ' salva', 'ok');
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

function ajustarImpressaoParaA4() {
  const sheet = document.querySelector('.sheet');
  if (!sheet) return;
  // Limpa ajustes inline anteriores (incluindo styles de medição que
  // possam ter sobrado de uma chamada interrompida).
  sheet.style.removeProperty('zoom');
  sheet.style.transform = '';
  sheet.style.transformOrigin = '';
  sheet.style.width = '';
  sheet.style.height = '';
  sheet.style.padding = '';
  sheet.style.minHeight = '';

  // Mede o conteudo SIMULANDO o estado de impressao — zoom: 1 (anula o
  // 1.20 da tela), width: 200mm fixo, padding: 0, min-height: auto.
  // Sem isso, scrollHeight reflete o layout de tela (zoom amplificado +
  // padding) e o scale calculado nao bate com o que o print engine usa,
  // o que faz conteudo extrapolar mesmo apos o ajuste.
  sheet.style.setProperty('zoom', '1');
  sheet.style.width = '200mm';
  sheet.style.padding = '0';
  sheet.style.minHeight = 'auto';
  void sheet.offsetHeight; // forca reflow pra leitura correta

  const pxPerMm = 3.7795275591;
  const maxHpx = 287 * pxPerMm; // A4 util com margem 5mm
  const natH = sheet.scrollHeight;

  // Restaura estado screen (estilos voltam pras regras CSS base).
  sheet.style.padding = '';
  sheet.style.minHeight = '';
  sheet.style.width = '';
  sheet.style.removeProperty('zoom');

  // Se o conteudo (a 200mm de largura) ultrapassa 287mm de altura, aplica
  // zoom < 1 inline com !important pra sobrepor o zoom: 1 do @media print.
  // zoom afeta LAYOUT (diferente de transform: scale, que e so visual e
  // mantem o layout box no tamanho original), entao tabelas, fontes e
  // quebras de linha encolhem proporcionalmente — o conteudo cabe sem
  // reflow imprevisivel. 1% de margem de seguranca evita borderline.
  if (natH > maxHpx) {
    const scale = (maxHpx / natH) * 0.99;
    sheet.style.setProperty('zoom', scale.toFixed(4), 'important');
  }
}

window.addEventListener('beforeprint', ajustarImpressaoParaA4);
window.addEventListener('afterprint', function() {
  const sheet = document.querySelector('.sheet');
  if (!sheet) return;
  sheet.style.removeProperty('zoom');
  sheet.style.transform = '';
  sheet.style.transformOrigin = '';
  sheet.style.width = '';
  sheet.style.height = '';
  sheet.style.padding = '';
  sheet.style.minHeight = '';
});

/* ========================================================= */
/*                    LISTA DE OS                            */
/* ========================================================= */
function renderListaOS() {
  const tb = document.getElementById('tbl-os');
  if (!STATE.ordens.length) { tb.innerHTML = `<tr><td colspan="7" class="empty">Nenhuma OS cadastrada ainda.</td></tr>`; return; }
  tb.innerHTML = STATE.ordens.slice().reverse().map(o => `
    <tr>
      <td><strong>${esc(o.os)||'—'}</strong></td>
      <td><span class="badge">${esc(o.codigo)||'—'}</span></td>
      <td>${esc(o.modeloNome)||'—'}</td>
      <td>${esc(o.colecaoNome)||'—'}</td>
      <td>${esc(formatDate(o.data))}</td>
      <td>${o.grade?.total||0} pç</td>
      <td class="col-actions row-actions">
        <button class="edit" onclick="verOS('${o.id}')">visualizar</button>
        <button class="edit" onclick="editarOS('${o.id}')">editar</button>
        <button class="edit" onclick="duplicarOS('${o.id}')">duplicar</button>
        <button class="del admin-only" onclick="excluirOS('${o.id}')">excluir</button>
      </td>
    </tr>`).join('');
}

function formatDate(iso) {
  if (!iso) return '—';
  const [y,m,d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

let printOsAtual = null;

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
    // Em modo de edicao, lista TODAS as grades cadastradas — nao filtra por
    // desenho/modelo/variacao. Motivo: o usuario pode estar editando pra
    // trocar tipo da grade (ex.: bicolor -> tricolor) e o filtro estrito
    // esconderia justamente as opcoes que ele quer ver. Tambem restaura
    // gradeId salvo (antes ficava sempre vazio na edicao).
    const gradeEl = document.getElementById('f-grade-preset');
    if (gradeEl) {
      fillSelect('f-grade-preset', STATE.grades, 'nome', '— nenhuma —');
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
        if (fase) nomeTecido = fase.tecidoNome || '';
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

  return `
    <table class="side-table" style="border-top:none;width:100%;">
      <thead>
        <tr><th colspan="${4 + colsTam.length + 1}" class="subhead" style="background:#c9e8d0;">Componentes — totais por tamanho</th></tr>
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

  return `
    <table class="side-table" style="border-top:none;width:100%;">
      <thead>
        <tr><th colspan="${3 + colsTam.length + 1}" class="subhead" style="background:#ffe0b2;">Aviamentos — totais por tamanho</th></tr>
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

  const camadas = e.camadas || 0;
  const fmt = n => n ? Number(n).toFixed(2).replace('.',',') : '—';

  // Cor/tecido de cada fase: vem de o.fases por ordem
  const fasesPorOrdem = {};
  (o.fases || []).forEach(f => { if (f?.ordem) fasesPorOrdem[f.ordem] = f; });

  // Linhas: blocos do enfesto se houver, ou pseudo-blocos derivados de tecidos
  // (OS sem fases/enfesto cadastrado mas com tecidos preenchidos manualmente)
  const linhas = blocos.length
    ? blocos.map((b, i) => ({ b, i }))
    : tecs.map((t, i) => ({ b: { ordem: i+1, nomeTecido: t.tecidoNome, nomeCor: t.corNome }, i }));

  const linhasEnfestos = linhas.map(({ b, i }) => {
    const ord = b.ordem || (i+1);
    const fase = fasesPorOrdem[ord] || {};
    let nomeEnf = b.nomeTecido || fase.tecidoNome || '';
    let cor = b.nomeCor || fase.corNome || '';
    if (!cor && nomeEnf.includes(' · ')) {
      const parts = nomeEnf.split(' · ');
      nomeEnf = parts[0];
      cor = parts.slice(1).join(' · ');
    }
    // Tecido real cadastrado (Moletom Bulk, Ribana Bulk, etc.) — fallback p/ tecido da OS
    const tecidoReal = fase.tecidoNome || tecs[i]?.tecidoNome || '';
    const corReal = cor || tecs[i]?.corNome || '';
    const camBloco = b.camadas || camadas || 0;
    // Comp/Larg priorizam o valor da fase ao vivo da grade (gViva já
    // foi remontado em renderPrintSheet a partir de STATE.grades). Cai pro
    // valor salvo no bloco do enfesto so como fallback — assim alteracoes
    // posteriores na grade refletem na impressao da OS sem precisar
    // re-aplicar a grade dentro de cada OS antiga.
    const compEf = (parseFloat(fase.comp) > 0 ? fase.comp : b.comp) || '';
    const largEf = (parseFloat(fase.larg) > 0 ? fase.larg : b.larg) || '';
    return `<tr>
      <td style="text-align:center;"><span style="display:inline-block;width:11px;height:11px;border:1.5px solid #000;vertical-align:middle;"></span></td>
      <td style="text-align:center;font-weight:700;">${ord}</td>
      <td>${esc(nomeEnf) || '—'}</td>
      <td>${esc(tecidoReal) || '—'}</td>
      <td>${esc(corReal) || '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${compEf ? fmt(compEf)+' m' : '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${largEf ? fmt(largEf)+' m' : '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;">${camBloco || '—'}</td>
      <td>&nbsp;</td>
    </tr>`;
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
        <tr><th colspan="9" class="subhead" style="background:#c9e8d0;">Enfesto${linhas.length>1?'s':''}</th></tr>
        <tr>
          <th style="font-size:6.5pt;white-space:nowrap;">✓</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Fase</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Enfesto</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Tecido</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Cor</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Compr.</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Largura</th>
          <th style="font-size:6.5pt;white-space:nowrap;">CAM</th>
          <th style="font-size:6.5pt;white-space:nowrap;">Consumo</th>
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

  const g = o.grade || {};
  const vars_ = o.variantes || [];
  const comps = ordenarComponentesPorFase(o.componentes || [], o);
  const avs = o.aviamentos || [];

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


  // Aviamentos
  const aviamentosHtml = avs.length
    ? avs.map(a => {
        const mat = STATE.materiais.find(m => m.id === a.material);
        return `<div>
          <span class="code">${esc(mat?.codigo||'—')}</span><br>
          <span style="font-size:6.5pt;color:#444;">${esc(mat?.desc||'')}</span>
          ${a.app?`<br><em style="font-size:6.5pt;">${esc(a.app)}</em>`:''}
        </div>`;
      }).join('')
    : `<div style="grid-column: 1/-1; color:#999; padding:4px 6px; font-style:italic;">Nenhum aviamento cadastrado</div>`;

  // Lista curta de componentes removida — info já aparece em "Componentes — totais por tamanho"

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

  document.getElementById('print-sheet').innerHTML = `
    <!-- CABEÇALHO -->
    <div class="sheet-header">
      <div class="cell brand-cell">${esc(o.griffeNome || o.griffe || 'MARCA')}</div>
      <div class="cell"><span class="mini">Coleção</span>${esc(o.colecaoNome || '—')}</div>
      <div class="cell"><span class="mini">${esc(o.blocoNome || o.bloco || 'R1 BLOCO 1')}</span></div>
      <div class="cell"><span class="mini">Data</span>${esc(formatDate(o.data))}</div>
      <div class="cell des-cell" style="flex-direction:column;align-items:center;justify-content:center;">
        <span class="mini">OS Nº:</span>
        <span style="font-size:13pt;letter-spacing:.05em;">${esc(o.os || '—')}</span>
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
        <div class="desenho-area">
          <div class="desenho-label">Desenho Técnico: ${esc(o.codigo || '—')}</div>
          ${imgHtml}
        </div>
      </div>

      <div class="sheet-right">
        <!-- GRADE -->
        <table class="side-table tab-tecidos">
          <thead>
            <tr><th colspan="8" class="subhead">Grade ${o.grade?.descricao?'· '+esc(o.grade.descricao):''}</th></tr>
            <tr>
              <th>P</th><th>M</th><th>G</th><th>GG</th><th>G1</th><th>G2</th><th>G3</th><th>Total</th>
            </tr>
          </thead>
          <tbody>
            <tr style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:600;">
              <td>${g.p>0?g.p:''}</td><td>${g.m>0?g.m:''}</td><td>${g.g>0?g.g:''}</td>
              <td>${g.gg>0?g.gg:''}</td><td>${g.g1>0?g.g1:''}</td><td>${g.g2>0?g.g2:''}</td><td>${g.g3>0?g.g3:''}</td>
              <td style="background:#fff59d;">${g.total>0?g.total:''}</td>
            </tr>
            ${(() => {
              const cam = o.enfesto?.camadas || 0;
              // multPrincipal: 1 camada produz quantas unidades por slot
              // de grade. Moletom = 1 (1 camada = 1 blusa). Malha sem
              // moletom (camiseta) = 2. Sem isso, o total por tamanho
              // sai pela metade pra camiseta e o usuario ve 'peças-alvo'
              // entrar como total geral.
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
              const t = (q) => (q > 0 && cam > 0) ? q * cam * multPrincipal : '';
              const totalGeral = (g.total || 0) * cam * multPrincipal;
              return `
                <tr><th colspan="8" class="subhead" style="background:#c9e8d0;font-size:6.5pt;">Total por tamanho</th></tr>
                <tr style="text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;background:#eaf6ed;">
                  <td>${t(g.p)}</td><td>${t(g.m)}</td><td>${t(g.g)}</td>
                  <td>${t(g.gg)}</td><td>${t(g.g1)}</td><td>${t(g.g2)}</td><td>${t(g.g3)}</td>
                  <td style="background:#c9e8d0;">${totalGeral > 0 ? totalGeral : ''}</td>
                </tr>`;
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
            // Mantém a ordem salva na OS; busca as tarefas embutidas na etapa cadastrada
            const ordenadas = o.etapas.map(nome => {
              const cad = STATE.etapas.find(e => e.nome === nome);
              return { nome, tarefas: cad ? tarefasDaEtapa(cad).map(t => t.nome) : [] };
            });
            const checkbox = `<span style="display:inline-block;width:10px;height:10px;border:1.5px solid #000;margin-right:8px;vertical-align:middle;flex-shrink:0;"></span>`;
            const subCheckbox = `<span style="display:inline-block;width:8px;height:8px;border:1px solid #000;margin-right:5px;vertical-align:middle;flex-shrink:0;"></span>`;
            return `<ul style="list-style:none;padding-left:0;margin:0;font-size:9pt;">
              ${ordenadas.map(e => `
                <li style="padding:4px 6px;border-bottom:1px dotted #d4d0c5;">
                  <div style="display:flex;align-items:center;">
                    ${checkbox}
                    <strong>${esc(e.nome)}</strong>
                  </div>
                  ${e.tarefas.length ? `
                    <ul style="list-style:none;padding-left:24px;margin:3px 0 0 0;font-size:8.5pt;color:#555;">
                      ${e.tarefas.map(t => `
                        <li style="display:flex;align-items:center;padding:1px 0;">
                          ${subCheckbox}
                          <span>${esc(t)}</span>
                        </li>`).join('')}
                    </ul>` : `
                    <div style="padding-left:24px;font-size:8.5pt;color:#999;font-style:italic;">sem tarefa definida</div>`}
                </li>`).join('')}
            </ul>`;
          })()}
        </div>

        <!-- AVIAMENTOS -->
        <div style="background:#c9e8d0;padding:3px 6px;font-family:'IBM Plex Mono',monospace;font-size:7pt;font-weight:700;letter-spacing:.08em;text-transform:uppercase;text-align:center;border:1px solid #000;border-top:none;">Aviamentos</div>
        <div class="aviamentos-grid">${aviamentosHtml}</div>

        <div class="obs-box"><strong>Observações</strong>${esc(o.obs || '')}</div>
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
  ALL_KEYS.forEach(k => { data[k] = STATE[k]; });
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
  ALL_KEYS.forEach(k => STATE[k] = []);
  for (const k of ALL_KEYS) {
    await saveState(k);
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
    goto('home');
    // Tarefas em background — não bloqueiam a navegação
    snapshotDiario().catch(e => console.warn('snapshotDiario', e));
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
window.ajustarImpressaoParaA4 = ajustarImpressaoParaA4;
window.conectarPastaPdf = conectarPastaPdf;
window.desconectarPastaPdf = desconectarPastaPdf;
window.conectarPastaBackup = conectarPastaBackup;
window.desconectarPastaBackup = desconectarPastaBackup;
window.escreverBackupJsonAgora = escreverBackupJsonAgora;
window.verOS = verOS;
window.editarOS = editarOS;
window.editarOsAtual = editarOsAtual;
window.excluirOS = excluirOS;
window.duplicarOS = duplicarOS;
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
window.setUserRole = setUserRole;
window.listarUsuariosComPapel = listarUsuariosComPapel;
window.duplicarCadastro = duplicarCadastro;
window.toggleFolderGrade = toggleFolderGrade;
window.moverEtapaForm = moverEtapaForm;
window.moverEtapaDesenho = moverEtapaDesenho;
