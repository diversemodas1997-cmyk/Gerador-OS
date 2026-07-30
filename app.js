/* ========================================================= */
/*                  ESTADO E PERSISTÊNCIA                    */
/* ========================================================= */
const SUPA_URL = 'https://ckkqrjkhorvaahyazqsr.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNra3Fyamtob3J2YWFoeWF6cXNyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4MTY2MjMsImV4cCI6MjA5MjM5MjYyM30.yT3Tb6KKx4sDNJXetwIoA77WudWUqQ2gCgT7JLi0iT8';
const supa = (window.supabase && typeof window.supabase.createClient === 'function')
  ? window.supabase.createClient(SUPA_URL, SUPA_KEY)
  : null;

// Identidade do DISPOSITIVO (aba), NÃO da conta. Dois computadores logados com
// o MESMO usuário precisam se distinguir. O filtro de sync ("essa gravação fui
// eu que fiz?") não pode comparar o id da CONTA (updated_by): com login
// compartilhado os dois têm o mesmo id, então o 2º computador achava que a
// gravação do 1º era própria e a descartava — a OS nunca aparecia lá. Este id é
// único por aba (sessionStorage sobrevive ao F5; some ao fechar a aba). Vai
// junto no blob (campo _device) e é o que comparamos na chegada de cada update.
const DEVICE_ID = (() => {
  try {
    let id = sessionStorage.getItem('deviceId');
    if (!id) {
      id = (window.crypto && crypto.randomUUID)
        ? crypto.randomUUID()
        : 'dev-' + Math.random().toString(36).slice(2) + Date.now().toString(36);
      sessionStorage.setItem('deviceId', id);
    }
    return id;
  } catch (e) {
    // sessionStorage indisponível (modo restrito): id em memória, ainda por aba.
    return 'dev-' + Math.random().toString(36).slice(2) + Date.now().toString(36);
  }
})();

let cloudCache = null;
// A ÚLTIMA leitura do servidor falhou? Enquanto true, não é seguro salvar: o
// cloudCache pode estar vazio por causa da falha (não porque o usuário apagou).
// Sem este flag, o seed/migração do loadState tentava gravar logo após um load
// falho e disparava a trava "gravação bloqueada" em loop a cada reload — quando
// o problema real é de CARREGAMENTO (ex.: sessão expirada).
let _cloudLoadErro = false;
// Uma gravação está em andamento agora? Junto do saveTimer (edição pendente no
// debounce), serve para o realtime/polling NÃO reler o servidor no meio de uma
// edição local — senão o cloudLoad sobrescrevia o checklist/horário que o
// usuário acabou de marcar e ainda não foi salvo, revertendo na tela.
let _flushing = false;
// Chaves que ESTE dispositivo alterou desde a última gravação. No flush, só
// estas sobrescrevem o servidor; as demais adotam o valor do servidor. É o que
// impede um cache desatualizado de apagar o que outro dispositivo gravou numa
// chave que este nem tocou (causa raiz das perdas recorrentes de cadastro).
const _dirtyKeys = new Set();
// Quantas ancoras do organizador foram convertidas na migracao do 📌 (log unico).
let _opFixoMigradoAgora = null;
// BASE do merge de três vias: o valor de cada chave como o servidor tinha na
// última vez que este dispositivo sincronizou. Comparando BASE × LOCAL sabe-se o
// que ESTE dispositivo mudou de fato — e só isso sobe por cima do servidor.
// Sem a base, uma chave suja subia inteira: quem tinha uma cópia velha de
// `ordens` apagava a OS que outro dispositivo tinha acabado de criar.
let _baseline = {};
let currentUser = null;
let currentRole = null; // 'admin' | 'usuario' | null
let saveTimer = null;
// Reagenda uma gravação que FALHOU (rede oscilou). Sem isto a edição ficaria só
// na memória (chave suja) até o próximo save, e um recarregamento a apagaria.
let _retryTimer = null;
let inRecoveryFlow = false;
// Compras de materiais vindas do programa de Contabilidade (tabela própria
// compras_materiais no Supabase). O Gerador-OS só LÊ — entram como ENTRADAS
// no estoque. Não fazem parte do blob shared_data (fonte separada).
let comprasCache = [];
let comprasChannel = null;
// Catálogo de SKUs publicado pelo Estoque-Confeccao (tabela skus_catalogo).
// O Gerador-OS só LÊ — usado no dropdown de SKU dos cadastros de Desenho/Modelo.
let catalogoSkus = [];

// Traz o estado do servidor para o cache local SEM apagar edições ainda não
// gravadas. Adota do servidor toda chave que ESTE dispositivo não alterou desde
// a última gravação (_dirtyKeys); as sujas ficam intactas até subirem. Assim um
// recarregamento de fundo (realtime/polling) nunca reverte o que o usuário
// acabou de mudar e ainda não foi salvo — a brecha que apagava edições do
// planejamento de operações quando uma segunda sessão/aba gravava por cima.
function _adotarServidorPreservandoEdicoes(srvData) {
  const srv = (srvData && typeof srvData === 'object') ? srvData : {};
  if (!cloudCache) cloudCache = {};
  Object.keys(srv).forEach(k => {
    if (_dirtyKeys.has(k)) return;
    cloudCache[k] = srv[k];
    _baseline[k] = srv[k];        // chave limpa: o servidor passa a ser a base
  });
  // Chaves que sumiram do servidor e que não editamos saem também (paridade com
  // a substituição total antiga), preservando as sujas e o carimbo de device.
  Object.keys(cloudCache).forEach(k => {
    if (k === '_device' || _dirtyKeys.has(k)) return;
    if (!(k in srv)) { delete cloudCache[k]; delete _baseline[k]; }
  });
}

// Merge de três vias de uma LISTA DE REGISTROS (base × nosso × servidor), por id.
// Parte do que está no servidor e aplica por cima só o que ESTE dispositivo
// mudou de verdade: registro criado ou editado aqui entra; registro que sumiu
// daqui (exclusão intencional) sai; registro que este dispositivo nem tocou fica
// como está no servidor — inclusive os que ele nunca viu, que é o caso da OS
// criada em outro aparelho enquanto esta aba estava aberta com a lista velha.
// Devolve a lista mesclada em JSON, ou null quando não é lista de registros com
// id (aí quem chama mantém o comportamento antigo).
function _mergeListaPorRegistro(baseStr, localStr, srvStr) {
  const parse = s => {
    if (typeof s !== 'string') return null;
    try { const v = JSON.parse(s); return Array.isArray(v) ? v : null; } catch (e) { return null; }
  };
  const local = parse(localStr), srv = parse(srvStr);
  if (!local || !srv) return null;
  const comId = arr => arr.every(x => x && typeof x === 'object' && !Array.isArray(x) && x.id != null);
  if (!comId(local) || !comId(srv)) return null;
  const base = parse(baseStr) || [];
  const porId = new Map();
  srv.forEach(r => porId.set(String(r.id), r));            // ponto de partida: o servidor
  const baseTxt = new Map(base.map(r => [String(r.id), JSON.stringify(r)]));
  const idsLocais = new Set(local.map(r => String(r.id)));
  local.forEach(r => {                                     // criado ou editado aqui: manda
    const id = String(r.id);
    const antes = baseTxt.get(id);
    if (antes === undefined || antes !== JSON.stringify(r)) porId.set(id, r);
  });
  baseTxt.forEach((_, id) => { if (!idsLocais.has(id)) porId.delete(id); });   // apagado aqui: sai
  // Ordem: a da lista local (é a que o usuário vê); o que só existe no servidor
  // entra no fim, preservado.
  const saida = [];
  local.forEach(r => {
    const id = String(r.id);
    if (porId.has(id)) { saida.push(porId.get(id)); porId.delete(id); }
  });
  porId.forEach(r => saida.push(r));
  return JSON.stringify(saida);
}

async function cloudLoad() {
  if (!supa || !currentUser) return;
  let { data, error } = await supa
    .from('shared_data')
    .select('data')
    .eq('id', 'main')
    .maybeSingle();
  if (error) {
    // Causa mais comum de reincidência: access token EXPIRADO. Renova a sessão
    // UMA vez e relê antes de desistir — recupera sozinho o caso mais frequente
    // (o refresh token ainda vale). Só depois trata como erro real.
    try {
      const { data: sess } = await supa.auth.refreshSession();
      if (sess && sess.session) {
        if (sess.user) currentUser = sess.user;
        ({ data, error } = await supa
          .from('shared_data').select('data').eq('id', 'main').maybeSingle());
      }
    } catch (e) { /* refresh falhou (refresh token morto) — cai no erro abaixo */ }
  }
  if (error) {
    console.error('cloudLoad', error);
    _cloudLoadErro = true;
    cloudCache = {};
    _baseline = {};   // leitura falhou: sem base confiavel para o merge
    // Visibilidade do problema: costuma ser sessão expirada ou RLS. Sem o toast,
    // a falha silenciosa esconde o motivo (e o seed subsequente virava o popup
    // enganoso de "gravação bloqueada").
    setTimeout(() => toast(
      `Falha ao ler dados do servidor (${error.code || 'erro'}). ` +
      `A sessão pode ter expirado — saia e entre de novo, ou verifique a conexão.`,
      'err'
    ), 50);
    return;
  }
  _cloudLoadErro = false;
  _adotarServidorPreservandoEdicoes(data && data.data);
}

async function cloudFlush() {
  if (!supa || !currentUser || !cloudCache) return;
  // Se a ÚLTIMA leitura falhou, NÃO salvar: o cloudCache pode estar vazio pela
  // falha, e o seed/migração do loadState tentam gravar logo em seguida. Sem
  // este atalho, o fluxo caía na trava anti-apagamento e mostrava "gravação
  // bloqueada" — mensagem enganosa, pois o problema é de CARREGAMENTO. Aqui a
  // mensagem é a certa e o reload/re-login (que renova a sessão) resolve.
  if (_cloudLoadErro) {
    setSyncStatus('error');
    mostrarAlertaSalvamento('carregamento',
      'Não foi possível CARREGAR seus dados do servidor (a sessão pode ter expirado ou está sem conexão). '
      + 'Não edite nada agora, para não sobrescrever. Clique em "Recarregar agora"; se continuar, saia e entre de novo.');
    return;
  }
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
  // TRAVA ANTI-APAGAMENTO da EXPEDIÇÃO: um flush que zera as cargas de expedição
  // — mas ainda tem OS/desenhos, então escapa da trava acima — é a assinatura da
  // sobrescrita por cache velho que apagou as OEs. Só dispara se ESTE dispositivo
  // JÁ viu cargas nesta sessão (não bloqueia quem nunca usou expedição) e agora
  // tenta gravar zero; então confere o servidor e bloqueia se lá ainda houver
  // cargas. Ações intencionais (limpar/restaurar) liberam via _permitirFlushVazio.
  if (!_permitirFlushVazio && _appJaTeveExpedicao && _contarItens(cloudCache, 'expedicaoCargas') === 0) {
    let servidorTemCargas = false;
    try {
      const { data } = await supa.from('shared_data').select('data').eq('id', 'main').maybeSingle();
      const d = (data && data.data) || {};
      servidorTemCargas = _contarItens(d, 'expedicaoCargas') > 0;
    } catch (e) { console.warn('checagem anti-apagamento expedição', e); }
    if (servidorTemCargas) {
      console.error('cloudFlush BLOQUEADO: gravar 0 cargas de expedição sobre servidor que ainda tem OEs.');
      setSyncStatus('error');
      mostrarAlertaSalvamento('bloqueio',
        'A tela está sem nenhuma OE (carga de expedição), mas o servidor ainda tem as suas. '
        + 'Para evitar apagar as OEs, a gravação foi bloqueada e NADA foi sobrescrito. '
        + 'Não continue editando — clique em "Recarregar agora" para trazer as OEs de volta.');
      toast('⛔ Gravação bloqueada — suas OEs no servidor estão protegidas. Recarregue a página.', 'err');
      return;
    }
  }
  setSyncStatus('saving');
  _flushing = true;
  try {
    // MERGE POR CHAVE (concorrência otimista) — correção definitiva da perda
    // recorrente de cadastros. Relê o servidor e só sobrescreve as chaves que
    // ESTE dispositivo alterou (_dirtyKeys). As chaves que NÃO tocamos adotam o
    // valor do servidor — assim um cache velho não apaga o que outro dispositivo
    // gravou numa chave que este nem mexeu. Ações intencionais de limpar/
    // restaurar já marcam TODAS as chaves como dirty (via saveState), então
    // sobrescrevem tudo normalmente.
    const adotadas = [], mescladas = [];
    try {
      const { data: srv } = await supa.from('shared_data').select('data').eq('id', 'main').maybeSingle();
      const servidor = (srv && srv.data && typeof srv.data === 'object') ? srv.data : null;
      if (servidor) {
        Object.keys(servidor).forEach(k => {
          if (k === '_device') return;
          if (!_dirtyKeys.has(k)) {                            // não editamos: adota o do servidor
            if (cloudCache[k] !== servidor[k]) { cloudCache[k] = servidor[k]; adotadas.push(k); }
            return;
          }
          // Editamos ESTA chave. Se ela é uma lista de registros, sobe só o que
          // mudou aqui (merge por registro) — do contrário a lista inteira deste
          // dispositivo apagaria o que outro criou enquanto esta aba estava
          // aberta. Não sendo lista de registros, vale o nosso, como antes.
          if (cloudCache[k] === servidor[k]) return;
          const mesclado = _mergeListaPorRegistro(_baseline[k], cloudCache[k], servidor[k]);
          if (mesclado != null && mesclado !== cloudCache[k]) {
            cloudCache[k] = mesclado;
            mescladas.push(k);
          }
        });
      }
    } catch (e) { console.warn('merge re-leitura', e); /* segue com o cache local */ }
    cloudCache._device = DEVICE_ID; // carimba ESTE dispositivo antes de gravar
    const { error } = await supa.from('shared_data').upsert({
      id: 'main',
      data: cloudCache,
      updated_at: new Date().toISOString(),
      updated_by: currentUser.id
    }, { onConflict: 'id' });
    if (error) throw error;
    _dirtyKeys.clear();
    // O que acabou de subir vira a nova BASE: daqui pra frente, só o que mudar
    // em relação a isto é considerado edição deste dispositivo.
    _baseline = Object.assign({}, cloudCache);
    if (_retryTimer) { clearTimeout(_retryTimer); _retryTimer = null; }  // subiu: cancela retry pendente
    setSyncStatus('ok');
    if (!_blobEstaVazio(cloudCache)) _appJaTeveDados = true;
    if (_contarItens(cloudCache, 'expedicaoCargas') > 0) _appJaTeveExpedicao = true;
    // Adotamos mudanças do servidor em chaves que não editamos, ou mesclamos
    // registros de outro dispositivo numa que editamos: reflete no STATE pra não
    // re-sobrescrever depois. O realtime/polling do outro cuida do re-render.
    if (adotadas.length || mescladas.length) {
      try { await loadState(); } catch (e) { console.warn('loadState pós-merge', e); }
    }
    // Snapshot de contingência (local + pasta) do estado recém-salvo.
    salvarSnapshotContingencia();
    // Snapshot DIÁRIO no servidor. Ele rodava só ao ABRIR o app: uma aba deixada
    // aberta desde ontem, ou um dia inteiro sem recarregar, passava sem nenhuma
    // cópia no servidor — enquanto o backup de contingência acima seguia
    // gravando a cada save e dava a impressão de que tudo estava coberto.
    // Agora a primeira gravação de cada dia também garante o snapshot; o guarda
    // _snapDiarioDiaOk impede consulta ao servidor nos saves seguintes.
    snapshotDiario().catch(e => console.warn('snapshotDiario', e));
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
    // As chaves sujas continuam marcadas (não foram limpas). Reagenda a subida
    // com folga (sem martelar offline) — sem isto a edição ficava órfã na
    // memória e um recarregamento a apagaria. _adotarServidorPreservandoEdicoes
    // segura as sujas até esta tentativa (ou o próximo save) conseguir subir.
    if (!_retryTimer) {
      _retryTimer = setTimeout(() => { _retryTimer = null; if (_dirtyKeys.size) cloudFlush(); }, 8000);
    }
  } finally {
    _flushing = false;
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
        // Ignora só o eco da gravação DESTE dispositivo. Comparar o id da conta
        // (updated_by) fazia o 2º computador do MESMO login descartar a mudança
        // do 1º achando que era própria — e a OS nunca aparecia lá.
        if (payload.new.data && payload.new.data._device === DEVICE_ID) return;
        // Enquanto há edição local pendente/salvando, NÃO relê o servidor: o
        // cloudLoad sobrescreveria o checklist/horário que o usuário acabou de
        // marcar e ainda não foi salvo, revertendo na tela. O polling reaplica a
        // mudança remota assim que a edição estiver salva.
        if (saveTimer || _flushing) return;
        // NÃO confiar no payload.new.data: o Realtime TRUNCA payloads grandes, e
        // o campo `data` pode chegar AUSENTE ou vazio. Usá-lo direto zerava o
        // cloudCache ({}), esvaziava a tela (as OEs "sumiam") e travava todo
        // save com "gravação bloqueada". Trata o realtime só como SINAL e relê o
        // estado COMPLETO via REST (sem limite de tamanho), que é a verdade.
        await cloudLoad();
        if (_cloudLoadErro) return; // leitura falhou: cloudLoad já avisou
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
      // Mesmo critério do realtime: só pula se foi ESTE dispositivo que gravou
      // (e não qualquer sessão do mesmo login).
      if (data.data && data.data._device === DEVICE_ID) {
        lastSeenUpdatedAt = data.updated_at;
        return;
      }
      // Enquanto há edição local pendente/salvando, adia a aplicação: reler
      // agora reverteria o checklist/horário recém-marcado. Não atualiza o
      // marcador, então o próximo poll reaplica a mudança remota quando a
      // edição já estiver salva.
      if (saveTimer || _flushing) return;
      // Mudanca de outro usuario: aplica, mas preservando o que ESTE dispositivo
      // ainda não gravou (chaves sujas) — senão reverteria a edição local pendente.
      _adotarServidorPreservandoEdicoes(data.data);
      _cloudLoadErro = false; // chegou dado bom do servidor
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

// Gravar ao SAIR/ESCONDER a página: uma edição feita e a aba trocada/minimizada/
// fechada dentro dos 800ms do debounce ficaria sem subir. Ao esconder a página
// sobe o pendente na hora — só age se houver algo por gravar (chaves sujas).
function _flushPendentesAoSair() {
  if (!cloudCache || !_dirtyKeys.size || _flushing) return;
  if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try { cloudFlush(); } catch (e) { /* melhor esforço ao sair */ }
}
if (typeof document !== 'undefined') {
  // visibilitychange (trocar de aba / minimizar) mantém a página viva, então o
  // envio assíncrono conclui; pagehide é a última cartada ao fechar.
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') _flushPendentesAoSair();
  });
  window.addEventListener('pagehide', _flushPendentesAoSair);
}

function setSyncStatus(status) {
  const el = document.getElementById('authSync');
  if (!el) return;
  el.classList.remove('saving', 'error');
  if (status === 'saving') { el.textContent = '☁ Salvando...'; el.classList.add('saving'); }
  else if (status === 'error') { el.textContent = '☁ Erro ao salvar'; el.classList.add('error'); }
  else { el.textContent = '☁ Sincronizado'; esconderAlertaSalvamento(); }
}

// Banner de aviso no topo do conteúdo. tipo 'bloqueio' (trava anti-apagamento),
// 'carregamento' (falha ao LER os dados — sessão/conexão) ou 'erro' (falha de
// gravação por qualquer motivo).
function mostrarAlertaSalvamento(tipo, msg) {
  const box = document.getElementById('alertaSalvamento');
  if (!box) return;
  const ic = document.getElementById('alertaSalvamentoIcone');
  const tit = document.getElementById('alertaSalvamentoTitulo');
  const m = document.getElementById('alertaSalvamentoMsg');
  box.classList.remove('erro', 'bloqueio');
  // 'carregamento' usa o mesmo visual grave (vermelho) do bloqueio.
  box.classList.add(tipo === 'erro' ? 'erro' : 'bloqueio');
  if (ic) ic.textContent = tipo === 'erro' ? '⚠' : '⛔';
  if (tit) tit.textContent =
    tipo === 'carregamento' ? 'Não foi possível carregar seus dados do servidor'
    : tipo === 'bloqueio' ? 'Gravação bloqueada — seus dados no servidor estão protegidos'
    : 'Falha ao salvar no servidor';
  if (m) m.textContent = msg || '';
  box.classList.remove('hidden');
}

function esconderAlertaSalvamento() {
  const box = document.getElementById('alertaSalvamento');
  if (box) box.classList.add('hidden');
}

// Recarrega FORÇANDO versão nova: acrescenta um parâmetro à URL, então o
// navegador não serve o index.html cacheado (que apontaria pro app.js antigo).
// O location.reload() comum às vezes reabria a mesma versão em cache — por isso
// "Recarregar agora" parecia "não mudar nada".
function recarregarForcado() {
  try {
    const u = new URL(window.location.href);
    u.searchParams.set('_r', String(Date.now()));
    window.location.replace(u.toString());
  } catch (e) {
    window.location.reload();
  }
}

/* Snapshot diário: grava cópia do blob atual em shared_data_backups uma vez por dia. */
// Dia (LOCAL) cujo snapshot já foi resolvido nesta sessão — evita consultar o
// servidor a cada gravação.
let _snapDiarioDiaOk = null;

async function snapshotDiario() {
  if (!supa || !currentUser || !cloudCache) return;
  // Data LOCAL, não UTC. Com toISOString() qualquer sessão depois das 21h em
  // Brasília gravava a data de AMANHÃ: o slot do dia seguinte era consumido com
  // os dados da véspera e, no dia seguinte, o "já existe" fazia o app pular —
  // o dia inteiro acabava sem cópia nenhuma no servidor.
  const hoje = _expIso(new Date());
  if (_snapDiarioDiaOk === hoje) return;
  const { data: existente } = await supa
    .from('shared_data_backups')
    .select('id')
    .eq('snapshot_date', hoje)
    .maybeSingle();
  if (existente) { _snapDiarioDiaOk = hoje; return; }
  const { error } = await supa.from('shared_data_backups').insert({
    snapshot_date: hoje,
    data: cloudCache,
    created_by: currentUser.id
  });
  if (error) { console.warn('snapshot falhou', error); return; }
  _snapDiarioDiaOk = hoje;
  // Retenção: 30 dias (também em data local, para casar com o que foi gravado)
  const corte = new Date();
  corte.setDate(corte.getDate() - 30);
  await supa.from('shared_data_backups').delete().lt('snapshot_date', _expIso(corte));
}

async function listarSnapshots() {
  if (!exigirEdicao('ver os snapshots do servidor')) return;
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
        <td class="col-actions">
          <button class="btn small" onclick="baixarSnapshot(${s.id}, '${esc(s.snapshot_date)}')">Baixar</button>
          <button class="btn small danger" onclick="restaurarSnapshot(${s.id}, '${esc(s.snapshot_date)}')">Restaurar</button>
        </td>
      </tr>`).join('')}
    </tbody></table>`;
}

// Baixa um snapshot diário como arquivo, SEM aplicá-lo. "Restaurar" é
// tudo-ou-nada: sobrescreve o estado atual inteiro com o do dia escolhido, o que
// joga fora todo o trabalho feito depois daquele ponto. Quando o que se perdeu é
// só uma parte (o progresso de algumas OSs, por exemplo), o caminho seguro é
// abrir o snapshot, comparar e reimportar apenas o que falta — e para isso é
// preciso poder LER o snapshot sem detonar o presente.
// O arquivo sai no mesmo formato do "Importar JSON" (chaves com arrays reais).
async function baixarSnapshot(id, dataStr) {
  if (!supa) return;
  if (!exigirAdmin('baixar snapshots')) return;
  toast('Baixando snapshot...', '');
  const { data, error } = await supa
    .from('shared_data_backups')
    .select('data, snapshot_date, created_at')
    .eq('id', id)
    .maybeSingle();
  if (error || !data) { toast('Erro ao ler o snapshot: ' + ((error && error.message) || 'não encontrado'), 'err'); return; }
  // No blob cada chave é uma STRING JSON; o import espera arrays reais.
  const bruto = data.data || {};
  const saida = { __snapshot: { data: data.snapshot_date, criadoEm: data.created_at } };
  Object.keys(bruto).forEach(k => {
    const v = bruto[k];
    if (typeof v === 'string') { try { saida[k] = JSON.parse(v); } catch { saida[k] = v; } }
    else saida[k] = v;
  });
  const blob = new Blob([JSON.stringify(saida, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `snapshot-${dataStr}.json`;
  a.click();
  URL.revokeObjectURL(url);
  toast(`Snapshot de ${dataStr} baixado`, 'ok');
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

// QUEM ESCREVE É SÓ O ADMIN. Todo perfil que não é admin usa o programa em
// modo LEITURA: consulta e imprime, não cadastra nem edita. Um só administra a
// fábrica; o resto da casa lê a OS, a agenda e a OE na tela ou no papel.
//
// Esta função existia para abrir uma exceção ao planejamento da expedição
// (admin E usuario). A exceção acabou — ela continua existindo com o nome
// próprio porque é o que os pontos de expedição chamam, e porque a mensagem de
// recusa fica no vocabulário daquela tela.
function exigirEdicao(acao) {
  if (currentRole === 'admin') return true;
  toast(currentRole === 'usuario'
    ? `Seu acesso é de leitura — apenas o admin pode ${acao}`
    : `Faça login como admin para ${acao}`, 'err');
  return false;
}

// O papel guardado em `currentRole` serve à TELA — decidir o que mostrar. Ele
// NÃO pode ser a tranca de nada que importe: é uma variável de JavaScript no
// navegador de quem usa, e quem abre o console muda o valor em um segundo.
// Antes de mexer em permissão, o papel é lido de novo NO SERVIDOR.
//
// E nem isso é a tranca de verdade: a tranca está na função `set_user_role` do
// Supabase, que confere sozinha se quem chamou é admin (ver
// supabase-admin-roles.sql). Sem ela, qualquer usuário logado promove a si
// mesmo chamando a função direto, sem passar por esta tela. O que se faz aqui é
// não deixar a tela mentir e não gastar uma ida ao servidor à toa.
async function _ehAdminNoServidor() {
  if (!supa || !currentUser) return false;
  try {
    const { data, error } = await supa
      .from('user_roles').select('role').eq('user_id', currentUser.id).maybeSingle();
    if (error) return false;
    return ((data && data.role) || '') === 'admin';
  } catch (e) { return false; }
}

async function setUserRole(novoPapel) {
  if (!supa) return;
  if (!exigirAdmin('gerenciar usuários')) return;
  if (!(await _ehAdminNoServidor())) {
    // O papel em memória dizia admin e o servidor discorda: ou a conta foi
    // rebaixada nesta sessão, ou alguém mexeu na variável. Nos dois casos, a
    // tela volta a mostrar o que a conta realmente é.
    await carregarPapel();
    aplicarPermissoesUI();
    toast('Apenas admin pode gerenciar usuários', 'err');
    return;
  }
  const email = (document.getElementById('roleEmail').value || '').trim().toLowerCase();
  if (!email) { toast('Informe o e-mail do usuário', 'err'); return; }
  if (novoPapel === 'admin' && !confirm(
    `Promover ${email} a ADMIN?\n\nAdmin pode criar e apagar tudo — cadastros, OS, expedição — `
    + 'e também mudar o papel dos outros, inclusive o seu.')) return;
  const { error } = await supa.rpc('set_user_role', { user_email: email, new_role: novoPapel });
  if (error) { toast('Erro: ' + error.message, 'err'); return; }
  toast(`${email} agora é ${novoPapel}`, 'ok');
  document.getElementById('roleEmail').value = '';
  listarUsuariosComPapel();
}

async function listarUsuariosComPapel() {
  if (!supa) return;
  if (!exigirAdmin('ver os usuários e os papéis')) return;
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
  _cloudLoadErro = false; // restauramos dado bom
  cloudCache._device = DEVICE_ID; // este dispositivo é o autor da restauração
  const { error: upErr } = await supa.from('shared_data').upsert({
    id: 'main', data: cloudCache,
    updated_at: new Date().toISOString(),
    updated_by: currentUser.id
  }, { onConflict: 'id' });
  if (upErr) { toast('Erro ao gravar: ' + upErr.message, 'err'); return; }
  _baseline = Object.assign({}, cloudCache);   // restauracao e a nova base do merge
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
      _dirtyKeys.add(key);
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
      _dirtyKeys.add(key);
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
      _baseline = {};
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
    await migrarLimpezaDesenho0023();  // remove componentes duplicados do 0023 (1×, admin)
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
  _baseline = {};
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

// Chaves que mudam o conteúdo da OE. Gancho no saveState em vez de espalhar a
// chamada por cada função que mexe no plano (alocar, mover, excluir, cancelar,
// remarcar, recalcular volumes, marcar Ensaque na OS…): todo caminho passa por
// aqui, inclusive os que vierem depois.
const _CHAVES_OE = ['expedicaoCargas', 'expedicaoJanelas', 'expedicaoExcecoes'];

async function saveState(key) {
  try {
    // O cadastro de funções mudou: o índice de horas marcadas tem que ser refeito
    // (ele é lido milhares de vezes por clique e por isso vive em cache).
    if (key === 'funcoes') _opHoraFixaVersao++;
    await DB.set(key, JSON.stringify(STATE[key]));
    // Republica o snapshot de estoque p/ a Contabilidade quando muda algo
    // que altera os saldos. Best-effort; não bloqueia nem quebra o save.
    if (_CHAVES_CONTAB_SNAPSHOT.includes(key) && typeof construirContabSnapshot === 'function') {
      atualizarContabSnapshot();
    }
    // Mexeu no plano de expedição: regrava a OE na pasta, como a OS regrava a
    // folha e a etiqueta ao ser salva. Best-effort; sem pasta conectada não faz
    // nada e nunca bloqueia o save.
    if (_CHAVES_OE.includes(key) && typeof agendarAutoSaveOE === 'function') {
      agendarAutoSaveOE();
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

// Operador de esteira de corte: a esteira é AUTOMÁTICA, então a mesma pessoa
// toca duas operações ao mesmo tempo. Por isso as operações deste posto não
// entram na detecção de sobreposição (não são conflito). Basta o nome conter
// "operador" e "esteira" — o "Coordenador … Esteira de corte" acima não bate,
// pois é "coordenador", não "operador".
function ehFuncaoOperadorEsteira(nome) {
  const n = _normFuncaoNome(nome);
  return n.includes('operador') && n.includes('esteira');
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
  if (STATE.expedicaoCargas && STATE.expedicaoCargas.length) _appJaTeveExpedicao = true;
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
  // Seed defaults (etapas/componentes) SÓ em conta genuinamente NOVA. Se a conta
  // já tem dados (OS, funções, tecidos, modelos, grades, equipe) mas as etapas
  // vieram vazias, isso é um GLITCH de carregamento — semear os defaults e SALVAR
  // apagaria as etapas CUSTOM do servidor (causa raiz de "o programa apagou as
  // etapas cadastradas"). Nesse caso não mexe: a próxima carga boa traz de volta.
  const _contaPopulada = _appJaTeveDados
    || (STATE.funcoes && STATE.funcoes.length)
    || (STATE.tecidos && STATE.tecidos.length)
    || (STATE.modelos && STATE.modelos.length)
    || (STATE.grades && STATE.grades.length)
    || (STATE.equipe && STATE.equipe.length);
  if (!_contaPopulada) {
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
  }
  // Migração: operações planejadas no formato antigo (uma linha por OS, com
  // peças e sem duração) viram jornada de posto.
  try {
    if (migrarOperacoesParaJornada()) {
      await saveState('operacoes');
      // O marcador da migração do 📌 mora em meta: sem gravar, ela rodaria de
      // novo na próxima abertura e desfaria um "horário fixo" recém-marcado.
      if (_opFixoMigradoAgora != null) {
        await saveState('meta');
        if (_opFixoMigradoAgora > 0) {
          console.info(`Operações: ${_opFixoMigradoAgora} âncora(s) do organizador deixaram de aparecer como "horário fixo do usuário".`);
        }
        _opFixoMigradoAgora = null;
      }
    }
  } catch (e) { console.warn('migrarOperacoesParaJornada', e); }
  // Padroniza o NÚMERO da OS em quatro dígitos. "340" e "0340" são o mesmo
  // número para quem trabalha, mas eram strings diferentes para o programa: duas
  // OS distintas no cadastro e dois arquivos distintos na pasta de PDFs.
  // Só normaliza quando o resultado NÃO colide com outra OS — se colidir, é
  // duplicata de verdade e quem decide o que fazer é o usuário, avisado abaixo.
  try {
    const canon = o => _numeroOSCanonico(o.os);
    const usados = new Map();
    (STATE.ordens || []).forEach(o => {
      const k = canon(o);
      if (!usados.has(k)) usados.set(k, []);
      usados.get(k).push(o);
    });
    let ajustadas = 0;
    usados.forEach((lista, k) => {
      if (k === 'sem-numero' || lista.length !== 1) return;
      const o = lista[0];
      if (String(o.os || '').trim() !== k) { o.os = k; ajustadas++; }
    });
    if (ajustadas) {
      await saveState('ordens');
      console.info(`Números de OS padronizados em 4 dígitos: ${ajustadas}.`);
    }
    // Duplicatas de verdade: dois registros para o mesmo número. Não se resolve
    // sozinho — apagar a errada é decisão de quem toca a fábrica, e as duas
    // podem ter movimento de estoque, expedição e operação amarrados.
    const dup = Array.from(usados.entries()).filter(([k, a]) => k !== 'sem-numero' && a.length > 1);
    if (dup.length) {
      console.warn('OS com NÚMERO REPETIDO (o programa não escolhe qual vale — resolva em OS Salvas): '
        + dup.map(([k, a]) => `${k} → ${a.map(o => `${o.modeloNome || 'sem modelo'} de ${o.data}`).join(' | ')}`).join(' · '));
    }
  } catch (e) { console.warn('normalizarNumerosOS', e); }
  // Reparo: OS cujas camadas foram lançadas na FOLHA depois de a OS ter sido
  // criada sem elas. A folha passou a mostrar as peças certas, mas os
  // componentes continuaram com o número congelado no cadastro — zero. Quem lê o
  // número gravado (o estoque de corte e o seletor de OS da expedição) enxergava
  // a OS vazia, e ela sumia da lista de alocação sem nenhum aviso.
  // Só mexe em OS com camadas > 0 e componentes zerados: é exatamente o
  // descasamento, e não há como confundir com OS que ainda não tem quantidade.
  try {
    const consertadas = [], divergentes = [];
    (STATE.ordens || []).forEach(o => {
      const camadas = parseInt((o.enfesto || {}).camadas, 10) || 0;
      if (!(camadas > 0)) return;
      const soma = (o.componentes || []).reduce((s, c) => s + (Number(c.qtdTotal) || 0), 0);
      if (soma > 0) {
        // Componentes com número, mas de OUTRA quantidade de camadas: a folha
        // baixou (ou subiu) as camadas depois de a OS ser salva. Aqui NÃO se
        // conserta sozinho — mexer nisso mudaria o estoque de uma OS já
        // produzida, e essa é uma decisão de quem toca a fábrica. Só avisa.
        const copia = JSON.parse(JSON.stringify(o));
        if (recomputarQuantidadesOS(copia)) {
          const novo = (copia.componentes || []).reduce((s, c) => s + (Number(c.qtdTotal) || 0), 0);
          divergentes.push(`${o.os} (gravado ${soma}, pelas ${camadas} camadas seria ${novo})`);
        }
        return;
      }
      if (recomputarQuantidadesOS(o)) consertadas.push(o.os);
    });
    if (consertadas.length) {
      await saveState('ordens');
      console.info(`OS com quantidade recomposta das camadas do enfesto: ${consertadas.join(', ')}`);
    }
    if (divergentes.length) {
      console.warn('OS com componentes de uma quantidade de camadas antiga (não alteradas — reabra e salve a OS para acertar): '
        + divergentes.join(' · '));
    }
  } catch (e) { console.warn('recomputarQuantidadesOS', e); }
  // Cadastro de operações POR FUNÇÃO: traz para cada função as operações que ela
  // já executa no planejamento (sem tempo) e converte quem ainda estava no
  // formato de texto livre.
  try {
    // A foto inclui `opsSincronizadas` — é a memória do que já foi oferecido a
    // cada função, e ela precisa ser GRAVADA na primeira vez. Sem isso a linha de
    // base seria refeita a cada abertura e a operação apagada pelo usuário
    // voltaria, que é justamente o que esta memória existe para impedir.
    const foto = () => JSON.stringify((STATE.funcoes || []).map(f => [f.operacoes || null, f.opsSincronizadas || null]));
    const antes = foto();
    const novas = sincronizarOperacoesDasFuncoes();
    if (antes !== foto()) {
      await saveState('funcoes');
      if (novas) console.info(`Funções: ${novas} operação(ões) do planejamento cadastradas, sem tempo.`);
    }
  } catch (e) { console.warn('sincronizarOperacoesDasFuncoes', e); }
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

// Limpeza ÚNICA do desenho 0023 (Blusa Moletom Vermelha): a cópia do desenho
// mostarda deixou componentes redundantes "Frente/Mangas/Costas Blusa Moletom
// Básica" (Mostarda) duplicando os novos "Frente/Costas/Mangas". Remove os
// redundantes SÓ quando o nome-base existe no desenho — pra não apagar nada útil.
// Roda 1× (flag em meta) no admin; propaga pelo sync normal.
async function migrarLimpezaDesenho0023() {
  if (currentRole !== 'admin' || !Array.isArray(STATE.desenhos)) return;
  STATE.meta = STATE.meta || {};
  if (STATE.meta.limpezaDesenho0023V1) return;
  const d = STATE.desenhos.find(x => String(x.codigo || '').trim() === '0023');
  let removidos = 0;
  if (d && Array.isArray(d.componentes)) {
    const nomeDe = c => (c.nome || (STATE.componentes.find(x => x.id === c.componenteId) || {}).nome || '').trim();
    const nomesBase = new Set(d.componentes.map(nomeDe).map(n => n.toLowerCase()));
    const antes = d.componentes.length;
    d.componentes = d.componentes.filter(c => {
      const m = nomeDe(c).match(/^(Frente|Mangas|Costas)\s+Blusa\s+Moletom\s+B[aá]sica$/i);
      // remove só se o nome-base (Frente/Mangas/Costas) também existe no desenho
      return !(m && nomesBase.has(m[1].toLowerCase()));
    });
    removidos = antes - d.componentes.length;
  }
  STATE.meta.limpezaDesenho0023V1 = true;
  try {
    if (removidos) { await saveState('desenhos'); toast(`Desenho 0023 limpo: ${removidos} componente(s) duplicado(s) removido(s)`, 'ok'); }
    await saveState('meta');
  } catch (e) { console.warn('migrarLimpezaDesenho0023', e); }
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
  const paginaAnterior = document.querySelector('section.page:not(.hidden)')?.dataset?.page;
  // As telas de CADASTRO são de LEITURA para todo mundo — aqui não há rota
  // fechada. Quem está no chão precisa consultar a ficha do tecido, a grade, a
  // composição do desenho e quem faz cada etapa; barrar a tela inteira só fazia
  // essa gente vir perguntar. O que o não-admin não pode é ESCREVER: os botões
  // de novo/editar/duplicar/excluir são `admin-only` e cada função de escrita se
  // defende sozinha (salvarCadastro, excluirCadastro, duplicarCadastro).
  // Diferente do formulário de OS logo abaixo, que é escrita pura e não tem
  // versão de leitura — por isso aquele continua fechado na rota.
  // O FORMULÁRIO DA OS é a mesma página para criar e para editar, e é escrita
  // pura. Esconder os botões que levam até ele não basta: dá para chegar pelo
  // card do Início, pelo "+ Nova OS" da lista, pelo "editar" de uma OS e pelo
  // endereço. Fechar aqui, na rota, fecha todos os caminhos de uma vez — e
  // quem só consulta continua vendo a OS pela folha, que é onde ela se lê.
  if (page === 'nova-os' && currentRole && currentRole !== 'admin') {
    toast('Seu acesso é de leitura — apenas o admin cria ou edita OS', 'err');
    page = 'lista-os';
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
  // Saiu da folha de OE com um auto-save adiado? Agora grava — a seção já está
  // escondida, então a captura acontece fora da tela, sem perturbar nada.
  if (paginaAnterior === 'print-expedicao' && page !== 'print-expedicao' && _oeSaveAdiado) {
    _oeSaveAdiado = false;
    salvarPdfOeNaPasta({ silent: true }).catch(e => console.warn('auto-save OE ao sair', e));
  }

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
  if (page === 'print-operacoes') renderPrintPlanoOperacoes();
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
    atualizarExportFolderStatus();
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
  // Este modal é a FICHA do cadastro, e a ficha é de leitura para todo mundo:
  // é onde se vê a composição de um desenho, os tamanhos de uma grade, a
  // gramatura de um tecido. Para quem não é admin ele abre inerte (CSS em
  // styles.css, bloco de PAPÉIS) e sem o botão Salvar.
  // Só a ficha EM BRANCO é barrada: "novo" não tem o que ler.
  if (!editId && !exigirEdicao('criar cadastros')) return;
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
  // "Editar" só quem edita. Para quem consulta o título é "Ficha do/da …" —
  // dizer "Editar Tecido" numa tela que não deixa digitar é mentir para o
  // usuário e gerar chamado.
  const verbo = currentRole === 'admin' ? (editId ? 'Editar ' : 'Novo ') : 'Ficha · ';
  title.textContent = verbo + titles[tipo];

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
        <div class="field"><label>Excedente de enfesto (cm)</label><input type="number" min="0" step="1" id="m-excedente" value="${esc(item.excedente ?? '')}" placeholder="${EXCEDENTE_ENFESTO_PADRAO_CM}"><div class="field-hint">Quanto este tecido ganha de sobra no <b>comprimento</b> ao ser enfestado: a diferença entre a medida de <b>cortar</b> (a do risco do CAD) e a de <b>enfestar</b> (a que se cadastra na fase). É a ponta que a enfestadeira segura e a folga para o corte não morrer na borda. A <b>largura</b> não recebe nada — ela é a do tecido. Em branco vale ${EXCEDENTE_ENFESTO_PADRAO_CM} cm.</div></div>
        <div class="field full"><label>Composição / observação</label><input type="text" id="m-desc" value="${esc(item.desc||'')}" placeholder="Ex.: 65% algodão 35% poliéster"></div>
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
    const optsTp = opcoesPastaGrade('pasta', item.tipoPeca);
    const optsVr = opcoesPastaGrade('subpasta', item.variacao);

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
      </div>
      ${_gradeTemposHtml(item)}`;
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
        <div class="field full">
          <label>Operações desta função, o tempo de cada uma e o horário fixo</label>
          <div id="m-func-ops"></div>
          <button type="button" class="add-row-btn" onclick="addOperacaoFuncaoRow()" style="margin-top:6px;">+ Adicionar operação</button>
          <div class="field-hint">São as operações que este posto executa. O <b>tempo</b> é o que a operação costuma levar — ele vem preenchido no planejamento quando você escolher esta função e esta operação. Deixe em branco enquanto não souber.</div>
          <div class="field-hint"><b>Todo dia às</b> é para a operação de <b>hora marcada</b>: café, almoço, preparação das máquinas, limpeza do fim do expediente. Preenchida, ela entra sozinha em todo dia que você planejar, sempre nessa hora, e nem o encadeamento do posto nem o botão <b>Organizar o dia</b> a tiram do lugar — a fila do posto passa a se encaixar em volta dela. Em branco, a operação entra na fila normalmente.</div>
          <div class="field-hint">Essa operação é <b>independente das OS</b>: ela é da jornada, não do lote. Entra <b>uma vez por dia</b> — não uma por OS nem uma por fase do enfesto —, sem referência a OS nenhuma, e alocar mais OS no mesmo dia não a repete.</div>
        </div>
      </div>`;
    // As linhas são montadas depois do innerHTML, como as fases da grade.
    setTimeout(() => {
      const lista = _opsDaFuncao(item);
      if (lista.length) lista.forEach(o => addOperacaoFuncaoRow(o));
      else addOperacaoFuncaoRow();
    }, 0);
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

// Leva para o cadastro de cada FUNÇÃO as operações que ela já executa no
// planejamento, sem tempo (o tempo é o usuário que decide). Serve para o
// cadastro nascer preenchido e para uma operação NOVA, criada no plano, aparecer
// ali na próxima abertura em vez de o cadastro ficar congelado.
//
// Cada função guarda em `opsSincronizadas` os nomes que já foram oferecidos a
// ela. Um nome só é acrescentado UMA VEZ na vida. Sem essa memória, a rotina
// olhava só o que estava na lista naquele momento e repunha tudo o que o plano
// usava: apagar uma operação no cadastro não colava (ela voltava na abertura
// seguinte) e renomear criava uma duplicata, porque o nome antigo continuava no
// plano e era reinserido. Com 47 operações por OS vindas da cascata, quase todo
// nome estava no plano — na prática o campo tinha virado somente-leitura.
// Devolve quantas foram acrescentadas.
function sincronizarOperacoesDasFuncoes() {
  if (!Array.isArray(STATE.funcoes) || !STATE.funcoes.length) return 0;
  const usadas = new Map();   // chave da função → Map(nome normalizado → nome)
  (STATE.operacoes || []).forEach(op => {
    const nome = String(op.operacao || '').trim();
    if (!nome) return;
    const chave = op.funcaoId || _normNome(_opFuncaoNome(op));
    if (!chave) return;
    if (!usadas.has(chave)) usadas.set(chave, new Map());
    const m = usadas.get(chave);
    if (!m.has(_normNome(nome))) m.set(_normNome(nome), nome);
  });
  let novas = 0;
  STATE.funcoes.forEach(f => {
    const lista = _opsDaFuncao(f).slice();
    const jaTem = new Set(lista.map(o => _normNome(o.nome)));
    const doPlano = usadas.get(f.id) || usadas.get(_normNome(f.nome)) || new Map();
    // Primeira vez nesta função: a linha de base é TUDO o que o plano já usa
    // hoje. Assim a memória começa valendo na mesma abertura em que é criada, e
    // uma operação apagada antes desta regra não volta mais uma última vez.
    const primeiraVez = !Array.isArray(f.opsSincronizadas);
    const jaOferecidas = new Set(primeiraVez ? Array.from(doPlano.keys()) : f.opsSincronizadas);
    // O que está no cadastro AGORA também conta como oferecido — inclusive o que
    // o usuário digitou à mão. Assim, apagar qualquer linha cola para sempre,
    // mesmo que aquele nome apareça no plano só mais tarde.
    jaTem.forEach(norm => jaOferecidas.add(norm));
    let mudouLista = false;
    doPlano.forEach((nome, norm) => {
      if (jaTem.has(norm)) { jaOferecidas.add(norm); return; }
      // Já foi oferecida um dia e não está mais aqui: o usuário tirou de
      // propósito (ou renomeou). Respeita a decisão dele.
      if (jaOferecidas.has(norm)) return;
      lista.push({ nome, duracaoMin: 0, horaFixa: '' });   // sem tempo: quem sabe é a casa
      jaTem.add(norm);
      jaOferecidas.add(norm);
      mudouLista = true;
      novas++;
    });
    f.opsSincronizadas = Array.from(jaOferecidas);
    // Grava quando entrou nome novo e quando a função ainda estava no formato
    // antigo (texto em `acoes`) — é o que a converte para a lista com tempo, sem
    // perder nada. Sem o `mudouLista`, uma novidade em QUALQUER função reescrevia
    // a lista de TODAS elas.
    if (!Array.isArray(f.operacoes) || mudouLista) {
      f.operacoes = lista;
      f.acoes = lista.map(o => o.nome).join('\n');
    }
  });
  return novas;
}

// As operações de uma função, no formato novo [{nome, duracaoMin}]. Função ainda
// no formato antigo (uma ação por linha em `acoes`) é lida como lista sem tempo —
// ninguém perde o que já cadastrou e ninguém precisa redigitar.
function _opsDaFuncao(f) {
  if (f && Array.isArray(f.operacoes)) return f.operacoes;
  return String((f && f.acoes) || '').split(/\r?\n/)
    .map(s => s.trim()).filter(Boolean)
    .map(nome => ({ nome, duracaoMin: 0 }));
}

// Tempo cadastrado para uma operação de uma função (0 quando não há).
function _tempoOperacaoCadastrada(funcaoId, nomeOperacao) {
  const f = (STATE.funcoes || []).find(x => x.id === funcaoId);
  if (!f) return 0;
  const alvo = _normNome(nomeOperacao);
  if (!alvo) return 0;
  const achou = _opsDaFuncao(f).find(o => _normNome(o.nome) === alvo);
  return (achou && Number(achou.duracaoMin)) || 0;
}

// Uma linha do editor de operações da função: nome + tempo + horário fixo.
// O horário fixo é opcional e quer dizer "esta operação acontece TODO DIA nesta
// hora" — café, almoço, preparação das máquinas, limpeza do fim do expediente.
// Preenchido, o planejamento do dia já nasce com ela na hora certa, e nem o
// encadeamento do posto nem o "Organizar o dia" a arrastam de lugar.
function addOperacaoFuncaoRow(op = {}) {
  const cont = document.getElementById('m-func-ops');
  if (!cont) return;
  const dur = Math.max(0, Math.round(Number(op.duracaoMin) || 0));
  const div = document.createElement('div');
  div.className = 'func-op-row';
  div.innerHTML = `
    <input type="text" class="func-op-nome" value="${esc(op.nome || '')}" placeholder="Ex.: Enfesto, Movimentação de unidades cortadas" autocomplete="off" oninput="_funcOpTravarTempoDeEnfesto(this.closest('.func-op-row'))">
    <input type="number" class="func-op-h" min="0" step="1" value="${dur ? Math.floor(dur / 60) || '' : ''}" placeholder="0" title="Horas">
    <span class="u">h</span>
    <input type="number" class="func-op-m" min="0" max="59" step="5" value="${dur ? dur % 60 : ''}" placeholder="0" title="Minutos">
    <span class="u">min</span>
    <span class="u as">todo dia às</span>
    <input type="time" class="func-op-fixo" value="${esc(op.horaFixa || '')}" title="Horário fixo: esta operação entra em todo dia planejado nesta hora e não é reencaixada na fila do posto. Em branco = entra na fila, como as demais.">
    <button type="button" class="btn small danger" title="Remover esta operação" onclick="this.closest('.func-op-row').remove()">✕</button>
    <div class="func-op-nota"></div>`;
  cont.appendChild(div);
  _funcOpTravarTempoDeEnfesto(div);
}

// O ENFESTO não tem tempo cadastrável. Quanto ele leva é do TRABALHO, não do
// posto: depende do modelo, da grade, de qual fase é e de quantas camadas ela
// tem — "Corpo Parte 3" (5,73 m) não leva o mesmo que "Corpo Parte 2" (1,13 m),
// e um número só no operador de enfestadeira daria o mesmo tempo aos dois.
// O planejamento apura esse tempo do histórico de horários lançados nas folhas.
// Por isso os campos de duração desta linha ficam travados e o cadastro explica
// de onde o número vem.
function _funcOpTravarTempoDeEnfesto(row) {
  if (!row) return;
  const nome = row.querySelector('.func-op-nome')?.value || '';
  const ehEnfesto = _opEhEnfesto({ operacao: nome });
  const nota = row.querySelector('.func-op-nota');
  ['.func-op-h', '.func-op-m'].forEach(sel => {
    const el = row.querySelector(sel);
    if (!el) return;
    el.disabled = ehEnfesto;
    if (ehEnfesto) el.value = '';
    el.title = ehEnfesto ? 'O tempo do enfesto é apurado do histórico, não cadastrado aqui' : (sel === '.func-op-h' ? 'Horas' : 'Minutos');
  });
  row.classList.toggle('sem-tempo', ehEnfesto);
  if (nota) {
    nota.innerHTML = ehEnfesto
      ? 'Tempo <b>apurado automaticamente</b>: a média desta fase, para este tipo de roupa e esta grade, sai dos horários de enfesto lançados nas folhas de OS. Não se cadastra aqui porque o mesmo posto enfesta panos de tamanhos muito diferentes.'
      : '';
  }
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
      <div class="field"><label>Tecido</label><select class="fase-tec" onchange="toggleUnidadesGrade(this); atualizarSugestaoBobinas(this)">${tecOpts(fase.tecidoId)}</select></div>
      <div class="field fase-unid-wrap"><label>Unidades da grade</label><select class="fase-unid">${unidadesOpts}</select><div class="field-hint fase-unid-hint">1 unidade da grade = N peças por camada (ribana). Ex.: 2x para Barra+Punhos moletom, 10x ou 20x para Gola.</div></div>
      <div class="field"><label>Comprimento (m)</label><input type="number" step="0.01" class="fase-comp" value="${esc(fase.comp || '')}" oninput="atualizarSugestaoBobinas(this)" placeholder="Ex.: 6,50"></div>
      <div class="field"><label>Largura (m)</label><input type="number" step="0.01" class="fase-larg" value="${esc(fase.larg || '')}" placeholder="Ex.: 1,80"></div>
      <div class="field full"><label>Bobinas previstas (consumo esperado)</label><input type="text" class="fase-bobinas" value="${esc(fase.bobinas != null && fase.bobinas !== '' ? String(fase.bobinas).replace('.', ',') : '')}" oninput="this.dataset.sug=''" placeholder="Ex.: 14  ·  1/2  ·  0"><div class="field-hint">Quantas bobinas deste tecido esta grade costuma consumir nesta fase. Aparece na coluna "Consumo" da folha de OS. Aceita fração (1/2) e zero.<span class="fase-bobinas-sug"></span></div></div>
    </div>`;
  cont.appendChild(div);
  toggleUnidadesGrade(div.querySelector('.fase-tec'));
  atualizarSugestaoBobinas(div.querySelector('.fase-tec'));
  renumerarFasesGrade();
}

// Preenche as bobinas pela regra da malha algodão assim que o tecido e o
// comprimento estão postos. NÃO pisa no que foi digitado à mão: só escreve no
// campo vazio ou no que a própria regra escreveu antes (marcado em data-sug).
// Quem digita apaga a marca (oninput do campo) e a regra não mexe mais ali.
function atualizarSugestaoBobinas(el) {
  const bloco = el?.closest?.('.fase-grade-bloco');
  if (!bloco) return;
  const comp = bloco.querySelector('.fase-comp')?.value;
  const tecId = bloco.querySelector('.fase-tec')?.value;
  const inp = bloco.querySelector('.fase-bobinas');
  const dica = bloco.querySelector('.fase-bobinas-sug');
  const sug = sugestaoBobinasFase(tecId, comp);
  if (dica) {
    dica.innerHTML = sug == null ? ''
      : ` <b style="color:var(--accent);">${esc(textoRegraBobinas(comp))}.</b>`;
  }
  if (!inp || sug == null) return;
  const atual = inp.value.trim();
  if (atual === '' || atual === inp.dataset.sug) {
    inp.value = String(sug);
    inp.dataset.sug = String(sug);
  }
}

// Quem mostra o campo "Unidades da grade" são DUAS fases, pelo mesmo motivo de
// fundo: uma camada delas rende mais de uma unidade da grade, então elas se
// enfestam com menos camadas que o tecido principal.
//
//   RIBANA — 2x para barra+punho, 10x ou 20x para gola.
//   FORRO  — o forro de capuz é a fase de malha numa grade que tem moletom. O
//            programa SEMPRE enfestou o forro pela metade das camadas do
//            moletom, o que é exatamente "2x"; só que esse 2 era fixo no
//            código, sem onde dizer outra coisa. Agora é o mesmo campo da
//            ribana, com 2x de padrão — quem rende diferente cadastra.
//
// A conta é do formulário inteiro, não de um bloco só: ser forro depende de
// existir uma fase de moletom em OUTRA linha, então trocar um tecido qualquer
// pode revelar (ou esconder) o campo de outra fase.
function toggleUnidadesGrade(_selectEl) {
  atualizarUnidadesDasFases();
}

const DICA_UNID_RIBANA = '1 unidade da grade = N peças por camada (ribana). Ex.: 2x para Barra+Punhos moletom, 10x ou 20x para Gola.';
const DICA_UNID_FORRO  = 'Quantas unidades da grade uma camada de forro rende. 2x é o padrão da casa (o forro enfesta com metade das camadas do moletom).';

function atualizarUnidadesDasFases() {
  const blocos = Array.from(document.querySelectorAll('#m-fases-container .fase-grade-bloco'));
  if (!blocos.length) return;
  const fases = blocos.map(b => ({
    nome: b.querySelector('.fase-nome')?.value || '',
    tecidoId: b.querySelector('.fase-tec')?.value || ''
  }));
  const papeis = calcularPapeisFases(fases);
  blocos.forEach((b, i) => {
    const wrap = b.querySelector('.fase-unid-wrap');
    if (!wrap) return;
    const tec = STATE.tecidos.find(t => t.id === fases[i].tecidoId);
    const ehForro = (papeis[i] || {}).papel === 'forro_capuz';
    wrap.style.display = (isTecidoRibana(tec) || ehForro) ? '' : 'none';
    const dica = wrap.querySelector('.fase-unid-hint');
    if (dica) dica.textContent = ehForro ? DICA_UNID_FORRO : DICA_UNID_RIBANA;
  });
}

function removerFaseGrade(btn) {
  const bloco = btn.closest('.fase-grade-bloco');
  if (bloco) bloco.remove();
  renumerarFasesGrade();
  // Tirar a fase de moletom faz as de malha deixarem de ser forro.
  atualizarUnidadesDasFases();
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
  if (!exigirEdicao('copiar etapas para todos os desenhos')) return;
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
  // Quantas OSs já emitidas mudam junto — o desenho não muda sozinho, ele leva
  // as OSs dele. Dizer isso ANTES é a diferença entre uma cópia consentida e uma
  // surpresa em 60 ordens de serviço.
  let osAfetadas = 0;
  alvos.forEach(d => {
    osAfetadas += _osDoDesenho(d).filter(o => {
      const final = _etapasFinaisOS(o, etapas);
      const atual = o.etapas || [];
      return !(atual.length === final.length && atual.every((e, i) => e === final[i]));
    }).length;
  });

  const ok = confirm(
    `Copiar as ${etapas.length} etapas do desenho "${codigo}" para os outros ${alvos.length} desenhos do modelo "${modeloLabel}"?\n\n`
    + `Etapas: ${etapas.join(', ')}\n\n`
    + (osAfetadas
      ? `Junto com os desenhos, ${osAfetadas} OS já emitida(s) recebem essa lista de etapas. Nenhuma marcação é removida: etapa já marcada que não estiver na lista continua na OS, no fim.\n\n`
      : `Nenhuma OS já emitida muda com isso.\n\n`)
    + `Só desenhos do MESMO modelo/variação (${modeloLabel}) são afetados — outros modelos, e as variações bicolor/tricolor, ficam intactos. As OSs do próprio desenho de origem não mudam, já que ele não é alterado. Esta ação não pode ser desfeita automaticamente.`
  );
  if (!ok) return;
  await copiarEtapasEntreDesenhos(codigo);
}

// Etapas de uma OS que NÃO podem sair da lista: as que já têm marca — a própria
// etapa marcada, ou qualquer tarefa dela. Tirar da lista uma etapa marcada
// apagaria da folha um trabalho que já foi registrado; é exatamente assim que os
// checklists "sumiam" antes.
function _etapasComMarcaOS(o) {
  const prog = (o && o.progresso) || {};
  const marcadas = new Set();
  Object.entries(prog.etapasCheck || {}).forEach(([n, v]) => { if (v) marcadas.add(n); });
  Object.entries(prog.tarefasCheck || {}).forEach(([n, obj]) => {
    if (obj && typeof obj === 'object' && Object.values(obj).some(Boolean)) marcadas.add(n);
  });
  return marcadas;
}

// Propaga as etapas de produção do desenho técnico para as OSs já emitidas com
// ele. Antes isso era manual — reabrir cada OS e reselecionar o desenho — e na
// prática ninguém fazia: mudava-se a etapa no desenho e as OSs em andamento
// seguiam imprimindo a lista velha.
// Etapa que já tem marca nunca é removida: vai para o fim da lista, preservando
// o que foi feito. Devolve quantas OSs mudaram.
async function propagarEtapasDesenhoParaOS(desenho, etapasAntes) {
  if (!desenho) return 0;
  const novas = (Array.isArray(desenho.etapasNomes) ? desenho.etapasNomes : []).filter(Boolean);
  if (!novas.length) return 0;
  // Sem mudança na lista do desenho, nenhuma OS é tocada.
  if (Array.isArray(etapasAntes)
      && etapasAntes.length === novas.length
      && etapasAntes.every((n, i) => n === novas[i])) return 0;

  const cod = (desenho.codigo || '').trim();
  // OS antiga pode não ter desenhoId — aí casa pelo código, como na lista de OS.
  const alvos = (STATE.ordens || []).filter(o =>
    (o.desenhoId && o.desenhoId === desenho.id)
    || (!o.desenhoId && cod && (o.codigo || '').trim() === cod));

  let n = 0;
  alvos.forEach(o => {
    const final = _etapasFinaisOS(o, novas);
    const atual = o.etapas || [];
    if (atual.length === final.length && atual.every((e, i) => e === final[i])) return;
    o.etapas = final;
    n++;
  });
  if (n) await saveState('ordens');
  return n;
}

// Lista final de etapas de uma OS ao receber as do desenho: as do desenho, mais
// as que já têm marca e sumiram de lá — estas vão para o fim, nunca somem.
function _etapasFinaisOS(o, novas) {
  const marcadas = _etapasComMarcaOS(o);
  const preservar = (o.etapas || []).filter(e => marcadas.has(e) && !novas.includes(e));
  return [...novas, ...preservar];
}

// OSs que um desenho alcança: pelo vínculo direto e, nas OSs antigas sem ele,
// pelo código do desenho.
function _osDoDesenho(desenho) {
  const cod = (desenho.codigo || '').trim();
  return (STATE.ordens || []).filter(o =>
    (o.desenhoId && o.desenhoId === desenho.id)
    || (!o.desenhoId && cod && (o.codigo || '').trim() === cod));
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
  const mudados = [];
  STATE.desenhos.forEach(d => {
    if (d.id === origem.id) return;
    if (chaveModeloDesenho(d) !== chaveOrigem) return;   // só mesmo modelo/variação
    const antes = Array.isArray(d.etapasNomes) ? [...d.etapasNomes] : null;
    d.etapasNomes = [...etapasNomes];
    mudados.push({ desenho: d, antes });
    alteradas++;
  });
  await saveState('desenhos');
  // Cada desenho que mudou leva junto as OSs já emitidas com ele — mesma regra
  // da edição avulsa do desenho.
  let osTocadas = 0;
  for (const m of mudados) {
    try { osTocadas += await propagarEtapasDesenhoParaOS(m.desenho, m.antes); }
    catch (e) { console.warn('propagarEtapasDesenhoParaOS', e); }
  }
  toast(`Etapas de "${codigoOrigem}" aplicadas a ${alteradas} desenho(s) do modelo "${modeloLabel}"`
        + (osTocadas ? ` e a ${osTocadas} OS já emitida(s)` : ''), 'ok');
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

// Cores da PEÇA de uma OS, na ordem canônica do desc do desenho: junta Cor 1, 2 e
// 3 de TODAS as variantes, tira o tecido do nome (corNomeCurto) e colapsa as
// repetições — preto na malha + preto na ribana é "Preto" uma vez só. O número de
// cores é limitado ao que o DESENHO tem: uma peça de uma cor cujo enfesto usa
// moletom + forro + ribana herda 3 cores na variante e sairia tricolor à toa.
// Ponto único de verdade: o banner da folha de OS e a linha da OS na folha de OE
// leem daqui, então as duas folhas dizem a mesma cor.
function coresDaPecaOS(o) {
  const desenho = (o && o.desenhoId) ? (STATE.desenhos || []).find(x => x.id === o.desenhoId) : null;
  const nCores = desenho
    ? (ordemCoresPorDesc(desenho).length
        || [desenho.corPrincipalId, desenho.corSecundariaId, desenho.corTerciariaId].filter(Boolean).length
        || 3)
    : 3;
  return ordenarCoresNomesPorDesc([...new Set(
    (((o || {}).variantes) || [])
      .flatMap(v => [v.cor1Nome, v.cor2Nome, v.cor3Nome].slice(0, nCores))
      .filter(c => c && c !== '—')
      .map(corNomeCurto)
  )], desenho);
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
  if (!exigirEdicao('criar ou editar cadastros')) return;
  const { tipo, editId } = cadastroContext;
  const list = pluralize(tipo);
  const v = id => document.getElementById(id)?.value || '';
  let item = editId ? STATE[list].find(x => x.id === editId) : { id: uid() };
  if (!item) item = { id: uid() };
  // `item` É o objeto guardado no STATE — os blocos abaixo o alteram no lugar.
  // Por isso a foto das etapas do desenho tem que ser tirada AGORA: é ela que
  // diz, no fim, se a lista mudou e se as OSs já emitidas precisam acompanhar.
  const etapasAntes = (tipo === 'desenho' && editId && Array.isArray(item.etapasNomes))
    ? [...item.etapasNomes] : null;

  if (tipo === 'tecido') {
    if (!v('m-nome')) return toast('Nome obrigatório', 'err');
    item.nome = v('m-nome');
    item.desc = v('m-desc');
    item.categoria = v('m-categoria');
    item.peso = parseFloat(String(v('m-peso')).replace(',', '.')) || 0;
    // Excedente de enfesto em CENTÍMETROS, como se fala no chão. Vazio é
    // diferente de zero: vazio significa "usa o padrão da casa", e zero
    // significa "este tecido não leva sobra nenhuma" — por isso não cai num
    // `|| 0`, que apagaria a diferença entre os dois.
    const exc = String(v('m-excedente')).trim().replace(',', '.');
    item.excedente = exc === '' ? '' : (Math.max(0, parseFloat(exc)) || 0);
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
    // Operações da função, cada uma com o seu tempo. `acoes` continua sendo
    // gravado (só os nomes, uma por linha) porque é dele que saem as sugestões de
    // operação no planejamento e a coluna da tabela — assim nada que já lê `acoes`
    // precisa mudar junto.
    item.operacoes = Array.from(document.querySelectorAll('#m-func-ops .func-op-row')).map(row => {
      const nome = (row.querySelector('.func-op-nome')?.value || '').trim();
      const h = parseInt(row.querySelector('.func-op-h')?.value, 10) || 0;
      const m = parseInt(row.querySelector('.func-op-m')?.value, 10) || 0;
      // Horário fixo: só entra se for hora de verdade. Campo em branco significa
      // "entra na fila do posto", que é o comportamento de sempre.
      const fixa = (row.querySelector('.func-op-fixo')?.value || '').trim();
      // Enfesto não guarda duração: ela é apurada do histórico a cada
      // planejamento. Zerar aqui impede que um número antigo, cadastrado antes
      // desta regra, continue disputando com o tempo apurado.
      const ehEnfesto = _opEhEnfesto({ operacao: nome });
      return {
        nome,
        duracaoMin: ehEnfesto ? 0 : Math.max(0, h) * 60 + Math.max(0, m),
        horaFixa: _opMin(fixa) == null ? '' : fixa
      };
    }).filter(o => o.nome);
    item.acoes = item.operacoes.map(o => o.nome).join('\n');
    // `etapasIds` (a antiga marcação de etapas da função) não é mais editada
    // aqui: quem diz o que o posto faz é a lista de operações acima. O valor
    // antigo fica gravado como estava, sem uso — não se apaga o que o usuário
    // cadastrou um dia só porque a tela mudou.
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

  // Mudou a lista de etapas do desenho técnico? As OSs já emitidas com ele
  // acompanham sozinhas — era o passo manual que ninguém lembrava de fazer.
  if (tipo === 'desenho' && editId) {
    try {
      const n = await propagarEtapasDesenhoParaOS(item, etapasAntes);
      if (n) toast(`Etapas atualizadas em ${n} OS já emitida(s)`, 'ok');
    } catch (e) { console.warn('propagarEtapasDesenhoParaOS', e); }
  }

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
  if (!exigirEdicao('excluir cadastros')) return;
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

// Ações de uma linha de cadastro. Mesmo desenho da lista de OS: a ficha é de
// TODO MUNDO (o "ver" abre o modal em leitura, com os campos inertes), e o que
// escreve — editar, duplicar, excluir — é admin-only. "ver" e "editar" abrem a
// MESMA função: é o body.is-admin que decide qual dos dois aparece, e não o
// valor de currentRole no instante do render (o papel chega do servidor depois).
function acoesCell(tipo, id) {
  return `<td class="col-actions row-actions">
    <button class="edit leitura-only" onclick="openCadastroModal('${tipo}','${id}')">ver</button>
    <button class="edit admin-only" onclick="openCadastroModal('${tipo}','${id}')">editar</button>
    <button class="edit admin-only" onclick="duplicarCadastro('${tipo}','${id}')">duplicar</button>
    <button class="del admin-only" onclick="excluirCadastro('${tipo}','${id}')">excluir</button>
  </td>`;
}

async function duplicarCadastro(tipo, id) {
  if (!exigirEdicao('duplicar cadastros')) return;
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
  if (!STATE.tecidos.length) { tb.innerHTML = `<tr><td colspan="6" class="empty">Nenhum tecido cadastrado.</td></tr>`; return; }
  const catLabel = { malha: 'Malha algodão · máx 80', moletom: 'Moletom · máx 36', outro: 'Outro' };
  tb.innerHTML = STATE.tecidos.map(t => {
    // Quem não cadastrou aparece com o padrão em cinza — assim a coluna não
    // mente dizendo "—" para um tecido que na prática recebe 15 cm.
    const proprio = !(t.excedente === '' || t.excedente == null);
    return `
    <tr>
      <td><strong>${esc(t.nome)}</strong></td>
      <td>${esc(t.desc)}</td>
      <td><span class="badge">${esc(catLabel[t.categoria] || '—')}</span></td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${t.peso ? esc(t.peso) + ' g/m²' : '—'}</td>
      <td style="text-align:center;font-family:'IBM Plex Mono',monospace;${proprio ? '' : 'color:var(--ink-3);'}">${
        esc(excedenteEnfestoCm(t.id))} cm${proprio ? '' : ' <span style="font-size:10px;">(padrão)</span>'}</td>
      ${acoesCell('tecido', t.id)}
    </tr>`;
  }).join('');
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

// Índice de uma fase pelo id — as regras do lote parcial falam de duas fases
// nominadas (corte e costurando), e ler o índice daqui evita amarrar a posição.
function _faseIdxPorId(id) { return FASES_ESTOQUE.findIndex(f => f.id === id); }

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
  const iCorte = _faseIdxPorId('corte');
  const iCostura = _faseIdxPorId('costurando');
  (STATE.ordens || []).forEach(o => {
    const entrou = _faseEntrouOS(o, fase.entrada);
    // Modelo sobreposto: a OS "saiu" desta fase se o volume está em OUTRA fase
    // agora (a última etapa marcada não é a desta fase).
    const atual = faseAtualOS(o);
    // LOTE PARCIAL: o embarque também move peça. Os pacotes que já foram numa
    // carga de IDA saem do Estoque de corte e entram em Costurando na hora, sem
    // esperar a etapa Costura — e o que NÃO foi na carga continua no saldo do
    // corte, disponível para a próxima expedição. Só vale enquanto a OS, pelas
    // etapas, ainda está no corte: marcada a etapa seguinte, o modelo sobreposto
    // já leva o lote inteiro para a fase dela e a fração deixa de fazer sentido.
    const fracEmb = (atual === iCorte && (idx === iCorte || idx === iCostura))
      ? _expEmbarcadoOS(o).fracao : 0;
    const entradaParcial = idx === iCostura && !entrou && fracEmb > 0;
    if (!entrou && !entradaParcial) return;
    const saiu = entrou && atual !== idx;
    // Quais OSs listar na coluna OS desta linha:
    //  - padrão: só as que estão ATUALMENTE nesta fase (compõem o saldo).
    //  - fase.osTodasEntradas: TODAS as que entraram na fase (etapa marcada),
    //    mesmo que já tenham avançado (ex.: Costurando lista toda OS com Costura).
    //  - entrada parcial: a OS embarcou parte do lote, então compõe o saldo aqui
    //    mesmo sem a etapa marcada.
    const listarOS = entradaParcial || (fase.osTodasEntradas ? true : (atual === idx));
    const numOS = (o.os || '').toString().trim();
    componentesPorTecidoCorOS(o).forEach(it => {
      const cur = pegar(it.tecidoNome, it.corNome);
      const embarcado = Math.round(it.qtd * fracEmb);
      if (entradaParcial) {
        cur.entrada += embarcado;            // em Costurando entra só o que embarcou
      } else {
        cur.entrada += it.qtd;
        if (saiu) cur.saida += it.qtd;
        else if (idx === iCorte && embarcado > 0) cur.saida += embarcado;
      }
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
  const idxCorte = _faseIdxPorId('corte');
  const idxCostura = _faseIdxPorId('costurando');
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
  // O lote parcial muda a conta destas duas fases — dizer a regra aqui evita que
  // um saldo "quebrado" (parte da OS) pareça erro de contagem.
  const notaParcial = fase.id === 'corte'
    ? ' <b>Lote parcial:</b> os pacotes já alocados numa carga de <b>ida</b> contam como <b>saída</b> aqui e entram em <b>Costurando</b>; o que não foi na carga fica no saldo, disponível para a próxima expedição.'
    : (fase.id === 'costurando'
      ? ' <b>Lote parcial:</b> a OS entra aqui na proporção dos pacotes já embarcados na <b>ida</b>, mesmo antes de a etapa Costura ser marcada.'
      : '');
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
        <b>OS</b> = números das OS que estão nesta fase agora (várias separadas por vírgula).${notaParcial}
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

  // OSs atualmente NESTA fase. Com lote parcial, a OS que embarcou parte do lote
  // aparece nas DUAS: no corte com o que sobrou e em Costurando com o que saiu.
  const pacotes = (STATE.ordens || []).map(o => {
    const total = componentesPorTecidoCorOS(o).reduce((s, it) => s + it.qtd, 0);
    const fAtual = faseAtualOS(o);
    const frac = (fAtual === idxCorte && (faseIdx === idxCorte || faseIdx === idxCostura))
      ? _expEmbarcadoOS(o).fracao : 0;
    const embarcado = Math.round(total * frac);
    return {
      osId: o.id, osNumero: o.os || '', modelo: o.modeloNome || '', data: o.data || '',
      total, faseIdx: fAtual, embarcado, parcial: frac > 0 && frac < 1
    };
  }).filter(p => p.total > 0 && (p.faseIdx === faseIdx
      || (faseIdx === idxCostura && p.faseIdx === idxCorte && p.embarcado > 0)))
    .sort((a, b) => String(b.osNumero).localeCompare(String(a.osNumero), undefined, { numeric: true }));
  // Quantas peças da OS contam NESTA fase: no corte, o lote menos o que embarcou;
  // em Costurando (sem a etapa marcada), só o que embarcou.
  const pecasNaFase = p => {
    if (faseIdx === idxCostura && p.faseIdx === idxCorte) return p.embarcado;
    if (faseIdx === idxCorte && p.embarcado > 0) return Math.max(0, p.total - p.embarcado);
    return p.total;
  };
  const seloParcial = p => {
    if (!(p.embarcado > 0) || !p.parcial) return '';
    return faseIdx === idxCostura
      ? ' <span class="badge" style="background:#e6eefb;">parcial · embarcada</span>'
      : ` <span class="badge" style="background:#fdf0d5;">${fmt(p.embarcado)} pç já embarcadas</span>`;
  };
  const pacotesHtml = pacotes.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">OSs atualmente em ${esc(fase.titulo)}</h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Cada OS avança de fase automaticamente conforme as etapas do checklist são marcadas. Quando só parte do lote embarca numa carga, a OS conta nas duas fases: as peças que foram, em Costurando; as que ficaram, no Estoque de corte.</div>
      <table class="table">
        <thead><tr><th>OS</th><th>Modelo</th><th>Data</th><th style="text-align:right;">Peças</th><th class="col-actions">Ação</th></tr></thead>
        <tbody>
          ${pacotes.map(p => `
            <tr>
              <td><strong>${esc(p.osNumero) || '—'}</strong></td>
              <td>${esc(p.modelo) || '—'}${seloParcial(p)}</td>
              <td style="white-space:nowrap;">${esc(formatDate(p.data))}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${fmt(pecasNaFase(p))} pç</td>
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
  // Limite = capacidade do transporte, cadastrada UMA vez em Unidades e carga.
  // É o mesmo caminhão em toda janela, então não há limite por janela: o valor
  // cadastrado ali vale para todas as pernas e é isso que sai na folha de OE.
  const volMin = _expNum(cfg.volMin, 0);
  const volMax = _expNum(cfg.volMax, 0);
  const itens = _expCargasDa(oc.janela.id, oc.dataOrig, perna).map(c => {
    const o = (STATE.ordens || []).find(x => x.id === c.osId);
    const pecasOS = o ? _expPecasOS(o) : 0;
    // Lote parcial: as peças da carga são só a parte do lote que vai NELA — o
    // número ao lado dos volumes tem que falar da carga, não da OS inteira.
    const comp = (o && Array.isArray(c.pacotes)) ? _expPecasDaComposicao(o, c.pacotes) : null;
    return {
      carga: c,
      os: o,
      osNumero: o ? (o.os || '—') : '(OS excluída)',
      modelo: o ? (o.modeloNome || '') : '',
      pecas: comp ? Math.round(pecasOS * comp.fracao) : pecasOS,
      pecasOS,
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

/* =============== LOTE PARCIAL: pacotes por tamanho × tonalidade ===============
   A OS é decomposta na MESMA base das etiquetas: 1 pacote por vaga de tamanho
   (regra camiseta/moletom de _tamanhosDaGradeExpandido) para CADA tonalidade
   efetiva, + 1 pacote de reposição. Uma carga pode levar só PARTE desses
   pacotes; o que sobra fica "a alocar" numa próxima expedição. */

// Lista canônica de pacotes da OS: [{tam,tom}] (tom = número 1..3 ou null) na
// ordem tom→tamanho, mais o flag de reposição. `total` inclui a reposição.
function _expPacotesCanonicos(o) {
  const tamanhosBase = _tamanhosDaGradeExpandido(o);      // ['P','M',...] (camiseta repete a vaga)
  const tons = tonsEfetivos(((o || {}).progresso || {}).totalTamanhoTons || {});
  const nTons = Math.max(1, tons.length);
  const itens = [];
  for (let ti = 0; ti < nTons; ti++) {
    const tom = tons[ti] != null ? tons[ti] : null;
    tamanhosBase.forEach(tam => itens.push({ tam, tom }));
  }
  const temReposicao = itens.length > 0;
  return { itens, temReposicao, nTons, total: itens.length + (temReposicao ? 1 : 0) };
}

// Chave estável de um pacote (tamanho|tom) — agrupa as vagas iguais.
function _expChavePacote(p) { return `${p.tam}|${p.tom == null ? '-' : p.tom}`; }

// Conta pacotes de uma lista [{tam,tom}] num Map chave→{tam,tom,qtd}.
function _expContarPacotes(lista) {
  const m = new Map();
  (lista || []).forEach(p => {
    const k = _expChavePacote(p);
    const e = m.get(k) || { tam: p.tam, tom: p.tom, qtd: 0 };
    e.qtd++; m.set(k, e);
  });
  return m;
}

// Rótulo curto de um pacote pra tela ("G · tom 2", "G", "reposição").
function _expRotuloPacote(p) {
  if (p && p.rep) return 'reposição';
  const tam = (p && p.tam) || '?';
  return (p && p.tom != null) ? `${tam} · tom ${p.tom}` : tam;
}

// Ocorrências CANCELADAS como chaves 'janelaId|data': carga numa expedição
// cancelada não conta como alocada (aquele lote não vai sair).
function _expCancelSet() {
  const s = new Set();
  (STATE.expedicaoExcecoes || []).forEach(e => { if (e.tipo === 'cancelada') s.add(e.janelaId + '|' + e.data); });
  return s;
}

// O que já foi alocado de uma OS nas cargas NÃO canceladas (menos a carga em
// edição). Devolve a contagem por tamanho×tom, se a reposição já saiu, e se
// existe carga "cheia" ANTIGA (sem composição por pacote) — nesse caso a OS é
// tratada como já resolvida, pra não inventar pendência em dados antigos.
//
// Ida e volta descrevem o MESMO pacote (sai cortado, volta costurado), então as
// duas pernas NÃO se somam: o alocado de cada pacote é o MAIOR entre elas. Somar
// fazia um lote parcial parecer resolvido logo que a volta era preenchida — os
// pacotes que ficaram para trás desapareciam da lista "a alocar". Com `perna`,
// a conta olha só aquela perna (é o que o seletor de pacotes precisa: a volta
// pode levar de novo exatamente o que a ida levou).
function _expAlocadoOS(osId, exceptCargaId, perna = null) {
  const cancel = _expCancelSet();
  const pernas = new Map();
  (STATE.expedicaoCargas || []).forEach(c => {
    if (c.osId !== osId || c.id === exceptCargaId) return;
    if (cancel.has(c.janelaId + '|' + c.data)) return;
    const pn = c.perna === 'volta' ? 'volta' : 'ida';
    if (perna && pn !== perna) return;
    let e = pernas.get(pn);
    if (!e) { e = { contagem: new Map(), reposicao: false, legacyFull: false }; pernas.set(pn, e); }
    if (Array.isArray(c.pacotes)) {
      c.pacotes.forEach(p => {
        const k = _expChavePacote(p);
        const x = e.contagem.get(k) || { tam: p.tam, tom: p.tom, qtd: 0 };
        x.qtd++; e.contagem.set(k, x);
      });
      if (c.reposicao) e.reposicao = true;
    } else if ((Number(c.volumes) || 0) > 0) {
      e.legacyFull = true;   // carga antiga (só número): cobre a OS inteira
    }
  });
  const contagem = new Map();
  let reposicao = false, legacyFull = false;
  pernas.forEach(e => {
    if (e.reposicao) reposicao = true;
    if (e.legacyFull) legacyFull = true;
    e.contagem.forEach((x, k) => {
      const cur = contagem.get(k);
      if (!cur || x.qtd > cur.qtd) contagem.set(k, { tam: x.tam, tom: x.tom, qtd: x.qtd });
    });
  });
  return { contagem, reposicao, legacyFull };
}

// Remanescente de uma OS: quanto do lote total ainda não foi para NENHUMA
// expedição. `faltam` lista os pacotes de tamanho que restam (com qtd) e
// `repRestante` diz se a reposição ainda espera. `parcial` = tem alocado E tem
// sobra (o caso que o painel de remanescentes destaca).
function _expRemanescenteOS(o, exceptCargaId) {
  const vazio = { total: 0, alocado: 0, restante: 0, faltam: [], repRestante: false, legacyFull: false, parcial: false };
  if (!o) return vazio;
  const canon = _expPacotesCanonicos(o);
  if (!canon.total) return vazio;
  const aloc = _expAlocadoOS(o.id, exceptCargaId);
  const totKey = _expContarPacotes(canon.itens);
  const faltam = [];
  let restante = 0;
  if (!aloc.legacyFull) {
    totKey.forEach((e, k) => {
      const usados = (aloc.contagem.get(k) || {}).qtd || 0;
      const falta = Math.max(0, e.qtd - usados);
      if (falta > 0) { faltam.push({ tam: e.tam, tom: e.tom, qtd: falta }); restante += falta; }
    });
  }
  const repRestante = canon.temReposicao && !aloc.legacyFull && !aloc.reposicao;
  if (repRestante) restante += 1;
  const total = canon.total;
  const alocado = total - restante;
  return { total, alocado, restante, faltam, repRestante, legacyFull: aloc.legacyFull, parcial: alocado > 0 && restante > 0 };
}

// Texto curto dos pacotes que faltam: "G, GG · tom 1 · G, GG · tom 2 · +rep".
function _expFaltamTexto(rem) {
  const partes = (rem.faltam || []).map(f => `${f.qtd > 1 ? f.qtd + '× ' : ''}${_expRotuloPacote(f)}`);
  if (rem.repRestante) partes.push('reposição');
  return partes.join(' · ') || '—';
}

/* ---- peças de cada pacote: a ponte entre o lote parcial e o estoque ---- */

const _EXP_TAM_KEY = { P: 'p', M: 'm', G: 'g', GG: 'gg', G1: 'g1', G2: 'g2', G3: 'g3' };

// Quantas peças tem UM pacote de cada chave tamanho|tom. O pacote é uma VAGA de
// tamanho: na camiseta o mesmo tamanho pode ter várias vagas e as peças daquele
// tamanho se repartem entre elas. Os números por tamanho × tom vêm de
// totaisPorTamanhoTomOS — a mesma fonte da folha de OS e da folha de OE —, então
// o pacote da folha e o pacote do estoque contam a mesma peça. Enquanto a divisão
// entre tonalidades não foi digitada na OS, a coluna do tamanho é repartida em
// partes iguais entre os tons (é o melhor palpite e mantém a soma fechando).
function _expPecasPacoteOS(o) {
  const TT = totaisPorTamanhoTomOS(o);
  const vagas = new Map();
  _tamanhosDaGradeExpandido(o).forEach(t => vagas.set(t, (vagas.get(t) || 0) + 1));
  const tons = TT.tons.length ? TT.tons : [null];
  const mapa = new Map();
  let total = 0;
  vagas.forEach((nVagas, tam) => {
    const k = _EXP_TAM_KEY[tam];
    const col = k ? TT.colTotal(k) : 0;
    tons.forEach(tom => {
      const L = TT.linhas.find(x => x.tom === tom);
      const cel = (L && !TT.semDigitacao) ? (Number(L.cels[k]) || 0) : col / tons.length;
      const pecas = nVagas > 0 ? cel / nVagas : 0;
      mapa.set(_expChavePacote({ tam, tom }), pecas);
      total += pecas * nVagas;
    });
  });
  return { mapa, total, de: p => mapa.get(_expChavePacote(p)) || 0 };
}

// Peças de uma composição de carga ([{tam,tom}]) e a fração que ela representa do
// lote inteiro. É a fração que move o saldo entre Estoque de corte e Costurando.
function _expPecasDaComposicao(o, lista) {
  const vazio = { pecas: 0, total: 0, fracao: 0 };
  if (!o || !Array.isArray(lista)) return vazio;
  const pp = _expPecasPacoteOS(o);
  const pecas = lista.reduce((s, p) => s + pp.de(p), 0);
  return { pecas, total: pp.total, fracao: pp.total > 0 ? Math.min(1, pecas / pp.total) : 0 };
}

// Quanto do lote de uma OS já EMBARCOU: as peças dos pacotes alocados em cargas
// de IDA não canceladas. A ida é que leva o corte para a costura — a volta traz o
// mesmo pacote de volta e contá-la aqui tiraria a peça do estoque duas vezes.
// Carga antiga (só o número de volumes, sem composição) embarca o lote inteiro.
function _expEmbarcadoOS(o) {
  const vazio = { pecas: 0, total: 0, fracao: 0, parcial: false };
  if (!o) return vazio;
  const cancel = _expCancelSet();
  const cargas = (STATE.expedicaoCargas || []).filter(c =>
    c.osId === o.id && c.perna !== 'volta' && !cancel.has(c.janelaId + '|' + c.data));
  if (!cargas.length) return vazio;
  if (cargas.some(c => !Array.isArray(c.pacotes) && (Number(c.volumes) || 0) > 0)) {
    return { pecas: 0, total: 0, fracao: 1, parcial: false };   // carga cheia antiga
  }
  const pp = _expPecasPacoteOS(o);
  let pecas = 0;
  cargas.forEach(c => (c.pacotes || []).forEach(p => { pecas += pp.de(p); }));
  if (!(pp.total > 0)) return vazio;
  const fracao = Math.min(1, pecas / pp.total);
  return { pecas, total: pp.total, fracao, parcial: fracao > 0 && fracao < 1 };
}

// Aviso na tela quando o volume GRAVADO na carga não bate com a regra
// (tamanhos × tonalidades + 1). A propagação cobre as cargas futuras a partir do
// momento em que a tonalidade muda, mas não alcança as que já estavam gravadas
// com número velho antes disso — nem um ajuste manual que ficou defasado. Aqui
// elas ficam visíveis, em vez de irem caladas para a OE.
// Só avisa: quem decide é o usuário, que pode ter posto o número à mão de propósito.
function _expBadgeVolumeDivergente(item) {
  if (!item || !item.os || !(item.volumes > 0)) return '';
  if (item.carga && Array.isArray(item.carga.pacotes)) return '';  // lote parcial: volume vem dos pacotes, não da regra cheia
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
    if (Array.isArray(c.pacotes)) return;   // lote parcial: o volume vem dos pacotes, não da regra cheia
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

  // NÃO aloca nada. A OS entra numa OE quando o usuário a aloca no
  // planejamento da expedição, e sai quando ele a tira de lá — as duas coisas
  // com a janela e a perna escolhidas por ele. Marcar "Ensaque" é gesto do
  // checklist de PRODUÇÃO: diz que o lote está pronto, não em qual caminhão vai.
  // (Antes o tique criava a carga sozinho, na próxima janela: a OE do dia
  // amanhecia com OS que ninguém planejou ali, e quem marcou nem sabia.)
  if (checked && !STATE.expedicaoCargas.some(c => c.osId === os.id)) {
    toast(`OS ${num} pronta para expedir — aloque na OE pelo planejamento da Expedição`, 'ok');
  }
}

function moverCargaExp(cargaId) {
  const c = (STATE.expedicaoCargas || []).find(x => x.id === cargaId);
  if (!c) return;
  abrirModalExpCarga(c.janelaId, c.data, c.perna, c.osId, cargaId);
}

// ---- Busca de OE existente no planejamento ----
// Varre as ocorrências de expedição num intervalo amplo (±1 ano) e lista as que
// TÊM OS alocada e batem com o termo (nº de OS, modelo, data dd/mm ou nome da
// janela). Clicar num resultado leva o plano para aquele dia. A busca só mexe no
// dropdown de resultados — não re-renderiza o plano — pra não perder o foco.
let _oeBuscaTimer = null;
function buscarOePlano(q) {
  if (_oeBuscaTimer) clearTimeout(_oeBuscaTimer);
  _oeBuscaTimer = setTimeout(() => _oeBuscaExec(q), 180);
}
function _oeBuscaExec(q) {
  const box = document.getElementById('oe-busca-results');
  if (!box) return;
  const termo = (q || '').trim().toLowerCase();
  if (termo.length < 2) { box.style.display = 'none'; box.innerHTML = ''; return; }
  const hoje = _expHoje();
  const ocs = ocorrenciasExpedicao(_expAddDias(hoje, -365), _expAddDias(hoje, 365))
    .filter(oc => !oc.cancelada);
  const matches = [];
  ocs.forEach(oc => {
    const itens = resumoPernaExpedicao(oc, 'ida').itens.concat(resumoPernaExpedicao(oc, 'volta').itens);
    if (!itens.length) return;
    const dataBr = formatDate(oc.data);
    const hay = (oc.data + ' ' + dataBr + ' ' + (oc.janela.nome || '') + ' ' +
      itens.map(i => (i.osNumero || '') + ' ' + (i.modelo || '')).join(' ')).toLowerCase();
    if (hay.includes(termo)) {
      matches.push({
        anchor: oc.data, dataBr, dow: _EXP_DIAS_CURTO[_expData(oc.data).getDay()],
        janela: oc.janela.nome || 'Janela sem nome', n: itens.length
      });
    }
  });
  if (!matches.length) {
    box.innerHTML = '<div style="padding:10px 12px;color:var(--ink-3);font-size:13px;">Nenhuma OE encontrada com esse termo.</div>';
    box.style.display = 'block';
    return;
  }
  matches.sort((a, b) => String(b.anchor).localeCompare(String(a.anchor)));
  box.innerHTML = matches.slice(0, 25).map(m =>
    `<div class="oe-busca-item" onmousedown="irParaOeDia('${esc(m.anchor)}')" style="padding:8px 12px;cursor:pointer;border-bottom:1px solid var(--line-2);font-size:13px;">`
    + `<strong>${esc(m.dow)} ${esc(m.dataBr)}</strong> · ${esc(m.janela)} <span style="color:var(--ink-3);">· ${m.n} OS</span></div>`
  ).join('');
  box.style.display = 'block';
}
// Leva o plano para o dia da OE escolhida (modo diário + âncora na data). Usa
// onmousedown no item pra disparar ANTES do onblur do input fechar o dropdown.
function irParaOeDia(data) {
  expPlanoModo = 'dia';
  expPlanoAncora = data;
  try {
    sessionStorage.setItem('gos:exp:modo', 'dia');
    sessionStorage.setItem('gos:exp:ancora', data);
  } catch (e) {}
  const box = document.getElementById('oe-busca-results');
  if (box) { box.style.display = 'none'; box.innerHTML = ''; }
  renderExpedicaoPlano();
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
        <div style="display:flex;gap:6px;">
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
    const linhas = r.itens.length ? r.itens.map(i => {
      // Lote parcial: mostra o que ainda falta alocar desta OS (soma de todas as
      // cargas não canceladas). Só quando a expedição não foi cancelada.
      const rem = (!oc.cancelada && i.os) ? _expRemanescenteOS(i.os) : null;
      const parcialBadge = rem && rem.restante > 0
        ? ` <span class="exp-badge baixo" title="Faltam ${rem.restante} de ${rem.total} volume(s) desta OS: ${esc(_expFaltamTexto(rem))}">parcial · faltam ${rem.restante}</span>`
        : '';
      return `
      <div class="exp-os-row">
        <span class="num">${esc(i.osNumero)}</span>
        <span class="mod">${esc(i.modelo) || '—'}</span>
        <span class="qtd">${fmt(i.pecas)} pç</span>
        <span class="vol">${i.volumes > 0 ? fmt(i.volumes) + ' vol' : '<span class="exp-badge baixo" title="Ninguém disse quantos volumes esta OS ocupa">vol?</span>'}${_expBadgeVolumeDivergente(i)}${parcialBadge}${
          i.carga.origem === 'ensaque' ? ' <span class="exp-badge info" title="Entrou nesta OE ao ser marcada como ensacada no checklist da OS, não pelo planejamento da expedição">pelo ensaque</span>' : ''}${
          i.carga.feita ? ' <span class="exp-badge ok" title="Marcada como feita no quadrinho da folha de OE">feita</span>' : ''}</span>
        <span><button title="Mudar o dia e o horário em que esta OS será expedida" onclick="moverCargaExp('${esc(i.carga.id)}')">⇄</button><button class="admin-only" title="Tirar esta OS da carga" onclick="excluirCargaExp('${esc(i.carga.id)}')">×</button></span>
      </div>`;
    }).join('') : '<div class="exp-vazio">Nenhuma OS alocada.</div>';
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
        ${oc.cancelada ? '' : `<div style="margin-top:8px;display:flex;gap:6px;">
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
        <div style="display:flex;gap:6px;">
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

  // OSs alocadas PELA METADE: parte do lote já foi para alguma expedição, mas
  // sobraram pacotes esperando. É o rastro do lote parcial — sem esta lista, os
  // pacotes que ficaram para trás sumiriam da vista.
  const remanescentes = (STATE.ordens || [])
    .map(o => ({ o, rem: _expRemanescenteOS(o) }))
    .filter(x => x.rem.parcial)
    .sort((a, b) => String(b.o.os || '').localeCompare(String(a.o.os || ''), undefined, { numeric: true }));
  const remanescentesHtml = remanescentes.length ? `
    <div class="card">
      <h2 style="margin:0 0 8px;font-size:14px;">OSs com pacotes a alocar <span class="exp-badge baixo">${remanescentes.length}</span></h2>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Estas OSs foram alocadas <b>em parte</b>: já entraram em alguma expedição, mas sobraram pacotes (tamanho × tonalidade) esperando embarcar. As peças desses pacotes continuam no <b>Estoque de corte</b> — só o que embarcou passou para <b>Costurando</b>. Use <b>alocar restante</b> para pôr o que falta numa expedição — já vem com os pacotes que sobraram marcados.</div>
      <table class="table">
        <thead><tr><th>OS</th><th>Modelo</th><th style="text-align:right;">Alocado</th><th>Faltam</th><th class="col-actions">Ações</th></tr></thead>
        <tbody>
          ${remanescentes.map(({ o, rem }) => `
            <tr>
              <td><strong>${esc(o.os) || '—'}</strong></td>
              <td>${esc(o.modeloNome) || '—'}</td>
              <td style="text-align:right;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${fmt(rem.alocado)}/${fmt(rem.total)}</td>
              <td style="font-size:12px;"><span class="exp-badge baixo">${fmt(rem.restante)}</span> ${esc(_expFaltamTexto(rem))}</td>
              <td class="col-actions row-actions">
                <button onclick="verOS('${esc(o.id)}')">ver OS</button>
                <button class="edit" onclick="abrirModalExpCarga('','','ida','${esc(o.id)}')">alocar restante</button>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';

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
                <button class="edit" onclick="abrirModalExpCarga('','','ida','${esc(o.id)}')">alocar</button>
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
    <div class="card">
      <div class="card-title">Janelas de expedição cadastradas</div>
      <div class="muted" style="font-size:12px;margin-bottom:8px;">Uma janela <b>semanal</b> se repete nos dias marcados; uma de <b>data fixa</b> acontece uma vez só. O limite de volume é único para todas — <b>${esc(_expLimitesTexto(_expNum(cfg.volMin, 0), _expNum(cfg.volMax, 0)))}</b>, definido em <b>Unidades e carga</b>.</div>
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
              <td>${esc(_expLimitesTexto(_expNum(cfg.volMin, 0), _expNum(cfg.volMax, 0)))}</td>
              <td>${j.ativo === false ? '<span class="exp-badge vazio">inativa</span>' : '<span class="exp-badge ok">ativa</span>'}</td>
              <td class="col-actions row-actions">
                <button class="edit" onclick="abrirModalExpJanela('${esc(j.id)}')">editar</button>
                <button class="del admin-only" onclick="excluirJanelaExp('${esc(j.id)}')">excluir</button>
              </td>
            </tr>`).join('') : '<tr><td colspan="7" class="empty">Nenhuma janela cadastrada.</td></tr>'}
        </tbody>
      </table>
    </div>`;

  const buscaHtml = `
    <div class="exp-busca no-print" style="position:relative;margin-bottom:8px;">
      <input id="oe-busca-input" type="search" autocomplete="off" placeholder="🔎 Buscar OE existente por OS, data (dd/mm) ou janela…"
        oninput="buscarOePlano(this.value)" onfocus="buscarOePlano(this.value)"
        onblur="setTimeout(function(){var b=document.getElementById('oe-busca-results');if(b)b.style.display='none';},200)"
        style="width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid var(--line);border-radius:6px;font-size:13px;background:var(--paper);color:var(--ink);">
      <div id="oe-busca-results" style="position:absolute;left:0;right:0;top:100%;z-index:30;background:var(--paper);color:var(--ink);border:1px solid var(--line);border-top:none;border-radius:0 0 6px 6px;max-height:300px;overflow:auto;display:none;box-shadow:0 8px 20px rgba(0,0,0,.14);"></div>
    </div>`;
  cont.innerHTML = toolbar + buscaHtml + comoFunciona + resumo + (ocs.length ? cards : vazio) + remanescentesHtml + pendentesHtml + janelasHtml;
}

/* ---------------- modais ---------------- */

let _expModalCtx = null;

function _expCampoNum(id, label, valor, hint) {
  return `<div class="field"><label>${label}</label><input type="number" min="0" step="1" id="${id}" value="${valor === '' || valor == null ? '' : esc(valor)}">${hint ? `<div class="field-hint">${hint}</div>` : ''}</div>`;
}

function abrirModalExpJanela(editId = null) {
  if (!exigirEdicao('cadastrar janelas de expedição')) return;
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
      <div class="field"><label>Situação</label><select id="ej-ativo"><option value="1" ${!j || j.ativo !== false ? 'selected' : ''}>Ativa</option><option value="0" ${j && j.ativo === false ? 'selected' : ''}>Inativa (não gera expedições)</option></select></div>
      <div class="field full"><label>Observação</label><input type="text" id="ej-obs" value="${esc(j ? (j.obs || '') : '')}" placeholder="Ex.: motorista da tarde"></div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">Toda expedição é interna, de <b>ida e volta</b> entre ${esc(cfg.unidadeA)} e ${esc(cfg.unidadeB)}. O limite de volume (mín/máx) é único, cadastrado em <b>Unidades e carga</b> (hoje: ${esc(_expLimitesTexto(_expNum(cfg.volMin, 0), _expNum(cfg.volMax, 0)))}) e vale para todas as janelas.</div>`;
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
  if (!exigirEdicao('alocar OS na expedição')) return;
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
        <select id="ec-ocorrencia" onchange="_expAtualizarSugestaoVolumes()">${opts || '<option value="">— nenhuma expedição planejada —</option>'}</select>
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
      <div class="field full" id="ec-pacotes-wrap" style="display:none;">
        <label>Pacotes desta OS nesta carga</label>
        <div id="ec-pacotes" class="ec-pacotes"></div>
        <div class="field-hint">Marque os pacotes (tamanho × tonalidade) que vão nesta carga. Desmarque para deixar o restante para outra expedição — o que sobrar fica listado como <b>a alocar</b>.</div>
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

// Monta o seletor de pacotes da OS (tamanho × tonalidade + reposição) no modal
// de carga. Cada linha por tonalidade traz um contador por tamanho, limitado ao
// que AINDA sobra pra alocar (o total da grade menos o que já foi pra outras
// cargas). Por padrão marca tudo que resta — alocar a OS inteira segue sendo um
// clique; tirar pacotes é opcional. Devolve true se montou o seletor (a OS tem
// grade), false se caiu no modo "número manual" (OS sem grade).
// Perna escolhida no modal de carga. O alocado é contado POR PERNA: a volta traz
// de volta os mesmos pacotes que a ida levou, então o que está na ida não pode
// bloquear a composição da volta (e vice-versa).
function _expPernaDoForm() {
  const p = String(document.getElementById('ec-ocorrencia')?.value || '').split('|')[2];
  return p === 'volta' ? 'volta' : 'ida';
}

function _expMontarSeletorPacotes(o) {
  const wrap = document.getElementById('ec-pacotes-wrap');
  const box = document.getElementById('ec-pacotes');
  const campo = document.getElementById('ec-volumes');
  if (!wrap || !box) return false;
  const canon = o ? _expPacotesCanonicos(o) : { total: 0 };
  if (!o || !canon.total) {           // sem grade → some o seletor, volume manual
    wrap.style.display = 'none';
    box.innerHTML = '';
    if (campo) { campo.readOnly = false; campo.classList.remove('is-auto'); }
    return false;
  }
  const editId = (_expModalCtx && _expModalCtx.editId) || '';
  const cargaEdit = editId ? (STATE.expedicaoCargas || []).find(c => c.id === editId) : null;
  const editouPacotes = cargaEdit && Array.isArray(cargaEdit.pacotes);
  const alocOutros = _expAlocadoOS(o.id, editId, _expPernaDoForm());   // já alocado em OUTRAS cargas da mesma perna
  const proprio = editouPacotes ? _expContarPacotes(cargaEdit.pacotes) : null;
  const totKey = _expContarPacotes(canon.itens);

  // Agrupa as chaves por tonalidade preservando a ordem canônica (tom→tamanho).
  const ordem = [];
  canon.itens.forEach(p => { const k = _expChavePacote(p); if (!ordem.includes(k)) ordem.push(k); });
  const porTom = new Map();   // tom → [{k,tam,tom,max,def}]
  ordem.forEach(k => {
    const e = totKey.get(k);
    const usadosOutros = (alocOutros.contagem.get(k) || {}).qtd || 0;
    const max = Math.max(0, e.qtd - usadosOutros);        // teto: o que sobra pra esta carga
    const def = editouPacotes ? Math.min(max, (proprio.get(k) || {}).qtd || 0) : max;
    const tomKey = e.tom == null ? '-' : e.tom;
    if (!porTom.has(tomKey)) porTom.set(tomKey, []);
    porTom.get(tomKey).push({ k, tam: e.tam, tom: e.tom, max, def });
  });

  let idx = 0;
  const blocos = Array.from(porTom.entries()).map(([tomKey, arr]) => {
    const titulo = tomKey === '-' ? 'Pacotes' : `Tom ${tomKey}`;
    const cels = arr.map(it => {
      const id = 'ecpk-' + (idx++);
      const esgotado = it.max <= 0;
      return `<label class="ec-pk-cel ${esgotado ? 'off' : ''}" title="${esgotado ? 'Já alocado em outra expedição' : 'Máximo ' + it.max}">
        <span class="t">${esc(it.tam)}</span>
        <input type="number" class="ec-pk-in" id="${id}" min="0" max="${it.max}" step="1"
          value="${it.def}" data-tam="${esc(it.tam)}" data-tom="${it.tom == null ? '' : it.tom}"
          ${esgotado ? 'disabled' : ''} oninput="_expRecalcVolSeletor()">
        <span class="mx">/${it.max}</span>
      </label>`;
    }).join('');
    return `<div class="ec-pk-tom"><div class="ec-pk-tom-h">${esc(titulo)}</div><div class="ec-pk-sizes">${cels}</div></div>`;
  });

  // Reposição: 1 pacote por OS (ribana/viés). Marcada por padrão quando ainda
  // não saiu em outra carga.
  let repHtml = '';
  if (canon.temReposicao) {
    const repOutros = alocOutros.reposicao;
    const repDef = editouPacotes ? !!cargaEdit.reposicao : !repOutros;
    const repDisabled = repOutros && !(editouPacotes && cargaEdit.reposicao);
    repHtml = `<div class="ec-pk-tom"><div class="ec-pk-tom-h">Reposição</div><div class="ec-pk-sizes">
      <label class="ec-pk-cel ${repDisabled ? 'off' : ''}" title="${repDisabled ? 'Já foi em outra expedição' : 'Pacote de reposição / ribana'}">
        <input type="checkbox" id="ecpk-rep" ${repDef ? 'checked' : ''} ${repDisabled ? 'disabled' : ''} onchange="_expRecalcVolSeletor()">
        <span class="t">1 pacote</span>
      </label></div></div>`;
  }

  box.innerHTML = blocos.join('') + repHtml;
  wrap.style.display = '';
  if (campo) { campo.readOnly = true; campo.classList.add('is-auto'); }
  _expRecalcVolSeletor();
  return true;
}

// Soma os contadores do seletor de pacotes → escreve em #ec-volumes e atualiza a
// linha de info. O campo de volumes fica só-leitura enquanto o seletor manda.
function _expRecalcVolSeletor() {
  const box = document.getElementById('ec-pacotes');
  const campo = document.getElementById('ec-volumes');
  if (!box) return;
  let n = 0;
  box.querySelectorAll('.ec-pk-in').forEach(inp => {
    let v = parseInt(inp.value) || 0;
    const max = parseInt(inp.max) || 0;
    if (v < 0) v = 0; if (v > max) { v = max; inp.value = String(max); }
    n += v;
  });
  const rep = document.getElementById('ecpk-rep');
  if (rep && rep.checked && !rep.disabled) n += 1;
  if (campo) campo.value = n > 0 ? String(n) : '';
  const info = document.getElementById('ec-info');
  const osId = document.getElementById('ec-os')?.value || '';
  const o = osId ? (STATE.ordens || []).find(x => x.id === osId) : null;
  if (info && o) {
    const canon = _expPacotesCanonicos(o);
    const editId = (_expModalCtx && _expModalCtx.editId) || '';
    const alocOutros = _expAlocadoOS(o.id, editId, _expPernaDoForm());
    // "restante depois desta carga" = total − (já em outras) − (o que marquei aqui)
    let jaOutros = 0; alocOutros.contagem.forEach(e => jaOutros += e.qtd); if (alocOutros.reposicao) jaOutros += 1;
    const sobra = Math.max(0, canon.total - jaOutros - n);
    info.innerHTML = `OS <b>${esc(o.os || '—')}</b> · ${esc(o.modeloNome || 'sem modelo')} · lote total <b>${canon.total}</b> volume(s).`
      + ` Nesta carga: <b>${n}</b>.`
      + (sobra > 0 ? ` <span class="exp-badge baixo">restam ${sobra} a alocar</span>` : ' <span class="exp-badge ok">OS completa</span>');
  }
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
  if (!o) {
    _expMontarSeletorPacotes(null);
    if (info) info.textContent = 'Selecione a OS para ver as peças.';
    return;
  }
  // Com grade: o seletor de pacotes assume (marca o que resta, calcula o volume
  // e escreve a info). Sem grade: cai no número manual, como antes.
  if (_expMontarSeletorPacotes(o)) return;
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
  if (!exigirEdicao('alocar OS na expedição')) return;
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
        <input type="checkbox" class="ev-os" value="${esc(i.carga.osId)}" data-vol="${i.volumes}" data-carga="${esc(i.carga.id)}" ${rep ? '' : 'checked'}>
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
  if (!exigirEdicao('configurar a expedição')) return;
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
    <div class="info-box" style="margin-top:8px;font-size:12px;">Este é o limite <b>único</b> de volume por <b>perna</b> (ida e volta contam separado), válido para <b>todas as janelas</b> — é o que aparece na folha de OE. A expedição é sempre interna, entre estas duas unidades.</div>`;
  openModal('modal-exp');
}

function abrirModalExpOcorrencia(janelaId, dataOrig) {
  if (!exigirEdicao('cancelar ou remarcar expedições')) return;
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
    if (!exigirEdicao('cadastrar janelas de expedição')) return;
    const nome = v('ej-nome').trim();
    if (!nome) return toast('Informe o nome da janela', 'err');
    const tipo = v('ej-tipo') === 'data' ? 'data' : 'semanal';
    const diasSemana = Array.from(document.querySelectorAll('.ej-dia:checked')).map(el => Number(el.value));
    if (tipo === 'semanal' && !diasSemana.length) return toast('Marque ao menos um dia da semana', 'err');
    const data = v('ej-data');
    if (tipo === 'data' && !data) return toast('Informe a data da expedição', 'err');
    const horaIda = v('ej-hora-ida'), horaVolta = v('ej-hora-volta');
    if (!horaIda || !horaVolta) return toast('Informe os horários de ida e de volta', 'err');
    // Limite de volume não é mais por janela: vem de Unidades e carga. Zera
    // qualquer override antigo pra não voltar a divergir da folha de OE.
    const reg = {
      nome, tipo, diasSemana: tipo === 'semanal' ? diasSemana : [],
      data: tipo === 'data' ? data : '',
      horaIda, horaVolta, volMin: '', volMax: '',
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
    if (!exigirEdicao('alocar OS na expedição')) return;
    const [janelaId, data, perna] = v('ec-ocorrencia').split('|');
    if (!janelaId || !data || !perna) return toast('Selecione a expedição', 'err');
    const osId = v('ec-os');
    if (!osId) return toast('Selecione a OS', 'err');
    // Composição por pacote (lote parcial), quando a OS tem grade. O seletor
    // manda no número de volumes; sem seletor (OS sem grade), vale o campo.
    const seletor = document.getElementById('ec-pacotes');
    const temSeletor = seletor && document.getElementById('ec-pacotes-wrap')
      && document.getElementById('ec-pacotes-wrap').style.display !== 'none';
    let pacotes = null, reposicao = false;
    if (temSeletor) {
      pacotes = [];
      seletor.querySelectorAll('.ec-pk-in').forEach(inp => {
        let q = parseInt(inp.value) || 0;
        const max = parseInt(inp.max) || 0;
        if (q > max) q = max;
        const tam = inp.dataset.tam;
        const tomRaw = inp.dataset.tom;
        const tom = tomRaw === '' ? null : (parseInt(tomRaw) || null);
        for (let i = 0; i < q; i++) pacotes.push({ tam, tom });
      });
      const repChk = document.getElementById('ecpk-rep');
      reposicao = !!(repChk && repChk.checked && !repChk.disabled);
    }
    const volumes = temSeletor
      ? (pacotes.length + (reposicao ? 1 : 0))
      : (parseInt(v('ec-volumes')) || 0);
    if (!(volumes > 0)) return toast(temSeletor ? 'Marque ao menos um pacote para esta carga' : 'Informe quantos volumes esta OS ocupa', 'err');
    if (!Array.isArray(STATE.expedicaoCargas)) STATE.expedicaoCargas = [];
    // Ao mover, a propria carga nao conta como duplicata dela mesma.
    const jaTem = STATE.expedicaoCargas.some(c => c.id !== ctx.editId
      && c.janelaId === janelaId && c.data === data && c.perna === perna && c.osId === osId);
    if (jaTem) return toast('Esta OS já está nesta carga', 'err');
    const campos = { janelaId, data, perna, osId, volumes, obs: v('ec-obs').trim() };
    // Guarda (ou limpa) a composição por pacote. Sem seletor, remove qualquer
    // composição antiga pra não ficar incoerente com o número manual.
    if (temSeletor) { campos.pacotes = pacotes; campos.reposicao = reposicao; }
    else { campos.pacotes = null; campos.reposicao = false; }
    if (ctx.editId) {
      const i = STATE.expedicaoCargas.findIndex(c => c.id === ctx.editId);
      if (i >= 0) STATE.expedicaoCargas[i] = { ...STATE.expedicaoCargas[i], ...campos };
    } else {
      STATE.expedicaoCargas.push({ id: uid(), ...campos });
    }
    await saveState('expedicaoCargas');
    toast(ctx.editId ? 'Expedição da OS alterada' : 'OS alocada na expedição', 'ok');

  } else if (ctx.tipo === 'volta') {
    if (!exigirEdicao('alocar OS na expedição')) return;
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
      // A volta espelha a ida: leva a MESMA composição por pacote da carga de
      // origem, quando ela tem uma (o que veio, volta).
      const origem = (STATE.expedicaoCargas || []).find(c => c.id === el.dataset.carga);
      const nova = {
        id: uid(), janelaId: ctx.janelaId, data: ctx.dataOrig, perna: 'volta',
        osId, volumes: parseInt(el.dataset.vol, 10) || 0, obs: ''
      };
      if (origem && Array.isArray(origem.pacotes)) {
        nova.pacotes = origem.pacotes.map(p => ({ tam: p.tam, tom: p.tom }));
        nova.reposicao = !!origem.reposicao;
      }
      STATE.expedicaoCargas.push(nova);
      n++;
    });
    if (!n) return toast('Essas OSs já estão na volta', 'err');
    await saveState('expedicaoCargas');
    toast(`${n} OS trazida(s) para a volta`, 'ok');

  } else if (ctx.tipo === 'config') {
    if (!exigirEdicao('configurar a expedição')) return;
    const unidadeA = v('ex-uni-a').trim(), unidadeB = v('ex-uni-b').trim();
    if (!unidadeA || !unidadeB) return toast('Informe o nome das duas unidades', 'err');
    const volMin = parseInt(v('ex-vol-min')) || 0;
    const volMax = parseInt(v('ex-vol-max')) || 0;
    if (volMax > 0 && volMin > volMax) return toast('O volume mínimo não pode ser maior que o máximo', 'err');
    if (!STATE.meta || typeof STATE.meta !== 'object') STATE.meta = {};
    STATE.meta.expedicao = { ...(STATE.meta.expedicao || {}), unidadeA, unidadeB, volMin, volMax };
    await saveState('meta');
    // Estes valores aparecem na folha (limite por perna, nomes das unidades),
    // mas moram em `meta` — que não está em _CHAVES_OE, a lista que dispara a
    // regravação da OE. Sem esta chamada, mudar o limite mínimo/máximo atualizava
    // a TELA (que lê expCfg ao vivo) e deixava o PDF da pasta com o valor antigo.
    // A chamada é explícita em vez de pôr 'meta' na lista: `meta` é gravada por
    // muitas outras razões (migrações, flags) que não mudam nada na folha.
    if (typeof agendarAutoSaveOE === 'function') agendarAutoSaveOE();
    toast('Configuração salva', 'ok');

  } else if (ctx.tipo === 'ocorrencia') {
    if (!exigirEdicao('cancelar ou remarcar expedições')) return;
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

// Marca/desmarca a OS da carga como FEITA — o quadrinho da folha de OE. É só o
// avanço do dia (separada e embarcada), não muda volume, pacote nem estoque:
// quem move peça é a composição da carga, não este visto.
async function alternarCargaFeita(id) {
  if (!exigirEdicao('marcar a carga como feita')) return;
  const c = (STATE.expedicaoCargas || []).find(x => x.id === id);
  if (!c) return;
  c.feita = !c.feita;
  await saveState('expedicaoCargas');
  renderExpedicaoPlano();
  const folha = document.getElementById('print-sheet-exp');
  if (folha && folha.innerHTML) renderPrintPlanoExpedicao();
}

async function excluirCargaExp(id) {
  if (!exigirAdmin('remover OS da expedição')) return;
  if (!confirm('Tirar esta OS da carga?')) return;
  STATE.expedicaoCargas = (STATE.expedicaoCargas || []).filter(c => c.id !== id);
  // Exclusão INTENCIONAL pode zerar as cargas: libera a trava anti-apagamento da
  // expedição só para este flush (senão remover a ÚLTIMA carga seria bloqueado).
  _permitirFlushVazio = true;
  try {
    await saveState('expedicaoCargas');
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; await cloudFlush(); }
  } finally { _permitirFlushVazio = false; }
  toast('OS removida da carga', 'ok');
  renderExpedicaoPlano();
}

// Redefine os volumes das cargas FUTURAS pela regra da grade (tamanhos + 1).
// Corrige valores gravados por regras antigas (ex.: OS que ficou com 600).
// Não toca em expedição já realizada — aquilo é histórico do que saiu.
async function recalcularVolumesExpedicao() {
  if (!exigirEdicao('recalcular volumes')) return;
  const hoje = _expHoje();
  let n = 0;
  (STATE.expedicaoCargas || []).forEach(c => {
    if (_expDataEfetivaCarga(c) < hoje) return; // já realizada: não reescreve
    if (Array.isArray(c.pacotes)) return;       // lote parcial: volume vem dos pacotes
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
  // Exclusão INTENCIONAL da janela (e das cargas dela) pode zerar as cargas:
  // libera a trava anti-apagamento da expedição só para este flush.
  _permitirFlushVazio = true;
  try {
    await saveState('expedicaoJanelas');
    await saveState('expedicaoCargas');
    await saveState('expedicaoExcecoes');
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; await cloudFlush(); }
  } finally { _permitirFlushVazio = false; }
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

// Paleta de cores distintas para as barras da linha de tempo. A cor identifica a
// FUNÇÃO (posto): todas as operações do mesmo posto saem na mesma cor. Cores
// escuras o bastante para o texto branco por cima; vermelho puro fica de fora
// para não se confundir com o contorno de conflito.
const _OP_PALETA = [
  '#2563eb', '#16a34a', '#9333ea', '#db2777', '#0891b2', '#d97706',
  '#4338ca', '#0d9488', '#be123c', '#65a30d', '#c026d3', '#0369a1'
];
// PRETO é reservado: só operação de horário FIXO (a que o usuário travou) sai
// nesta cor, em qualquer função. Assim, olhando a linha do tempo, o que não se
// move é reconhecido de longe — e nenhuma função disputa a mesma leitura, porque
// a paleta acima não tem tom escuro parecido.
const _OP_COR_FIXO = '#111827';
// Mapa op.id → cor, uma cor POR FUNÇÃO. A ordem vem de STATE.funcoes (estável),
// então a cor de um posto é a mesma em todos os dias, nas duas vistas (por posto
// e por pessoa) e na folha impressa. Funções fora do cadastro entram no fim.
function _opMapaCores(doDia) {
  const m = new Map();
  const ordemFuncao = new Map();
  (STATE.funcoes || []).forEach((f, i) => ordemFuncao.set(_normFuncaoNome(f && f.nome), i));
  let extra = (STATE.funcoes || []).length;
  const corDe = nome => {
    const key = _normFuncaoNome(nome);
    let idx = ordemFuncao.get(key);
    if (idx == null) { idx = extra++; ordemFuncao.set(key, idx); }
    return _OP_PALETA[idx % _OP_PALETA.length];
  };
  doDia.forEach(op => m.set(op.id, op.inicioFixo ? _OP_COR_FIXO : corDe(_opFuncaoNome(op))));
  return m;
}

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

// Vista da agenda: 'posto' (faixa por função) ou 'pessoa' (faixa por pessoa —
// mostra num quadro só tudo que cada pessoa faz e onde ela se sobrepõe).
let opVista = 'posto';
// Dia cujo diagnóstico está aberto na agenda (vazio = nenhum). É sob demanda: a
// análise é uma pergunta que o usuário faz, não um aviso permanente na tela.
let opAnaliseDia = '';

function analisarDiaOperacoes(data) {
  opAnaliseDia = (opAnaliseDia === data) ? '' : data;
  renderOperacoes();
  if (opAnaliseDia) {
    const el = document.getElementById('op-analise-' + opAnaliseDia);
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' });
  }
}

// Pares de operações da MESMA FUNÇÃO que se cruzam no tempo. É o conflito que o
// organizador do dia NÃO resolve sozinho — mexer na fila de um posto é decisão
// de quem planeja —, então é justamente o que a análise precisa mostrar.
function _opSobreposicoesMesmaFuncao(doDia) {
  const pares = [];
  _opGruposSobreposicao(doDia).forEach((arr, chave) => {
    if (chave.indexOf('|F|') < 0) return;
    arr.slice().sort((a, b) => _opInicioMin(a) - _opInicioMin(b)).forEach((op, i, lista) => {
      if (!i) return;
      const ant = lista[i - 1];
      const ini = Math.max(_opInicioMin(op), _opInicioMin(ant));
      const fim = Math.min(_opFimMin(op), _opFimMin(ant));
      if (fim > ini) pares.push({ funcao: _opFuncaoNome(op), a: ant, b: op, ini, fim, min: fim - ini });
    });
  });
  return pares;
}
try { opVista = sessionStorage.getItem('gos:op:vista') || opVista; } catch (e) {}
function opSetVista(v) {
  opVista = (v === 'pessoa') ? 'pessoa' : 'posto';
  try { sessionStorage.setItem('gos:op:vista', opVista); } catch (e) {}
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
// Etapas que a função executa (cadastradas nela), na ordem das etapas. É a fonte
// do dropdown de etapas no planejamento de operações.
// O que a função executa, para o modo "etapas específicas" do planejamento. A
// fonte passou a ser a lista de OPERAÇÕES da própria função (as que ela executa
// no plano, com o tempo de cada uma) no lugar das etapas de produção marcadas —
// planejar o posto por partes é escolher entre as operações dele, não entre as
// fases do fluxo da OS. Função antiga, sem operações cadastradas, ainda cai nas
// etapas marcadas, para não ficar sem nada enquanto não for aberta.
function _opEtapasDaFuncao(funcaoId) {
  const f = (STATE.funcoes || []).find(x => x.id === funcaoId);
  if (!f) return [];
  const ops = _opsDaFuncao(f).map(o => o.nome).filter(Boolean);
  if (ops.length) return ops;
  const ids = Array.isArray(f.etapasIds) ? f.etapasIds : [];
  if (!ids.length) return [];
  return etapasOrdenadas().filter(e => ids.includes(e.id)).map(e => e.nome).filter(Boolean);
}

function _opSugestoesOperacao(funcaoId) {
  return [...new Set(_opEtapasDaFuncao(funcaoId))];
}

// Nomes das etapas marcadas no checklist do modal de operação.
function _opEtapasMarcadas() {
  return Array.from(document.querySelectorAll('.op-etapa-chk:checked')).map(el => el.value);
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

// Sincroniza os horários DENTRO de um posto: cada operação começa quando a
// anterior termina (fim de uma = início da seguinte), na ordem da sequência do
// posto. A 1ª operação com horário é a ÂNCORA (mantém o início); as demais são
// reencaixadas — mesmo que isso mude o horário que estava lá. Operação sem
// duração não avança o relógio nem é reposicionada. Devolve se algo mudou.
//
// Exceção que faz o encadeamento ser ajustável: operação com `inicioFixo` é uma
// âncora nova. O horário dela foi escolhido à mão, então nunca é reescrito — e o
// relógio do posto passa a contar dali em diante, o que permite deixar INTERVALO
// VAZIO (posto parado, almoço, espera de tecido) entre uma operação e a seguinte.
function _opSincronizarPosto(itens) {
  let running = null, mudou = false;
  (itens || []).forEach(op => {
    const dur = _opDuracao(op);
    const ini = _opInicioMin(op);
    // Duas âncoras, com donos diferentes: `inicioFixo` é do USUÁRIO (a caixa
    // "horário fixo" ou uma hora digitada fora do encaixe) e aparece como 📌;
    // `inicioAuto` é do ORGANIZADOR do dia, que precisa que o horário calculado
    // por ele não seja desfeito pelo encadeamento — mas não é escolha de
    // ninguém, então não se apresenta como fixa nem pede que alguém a solte.
    if ((op.inicioFixo || op.inicioAuto) && ini != null) {
      running = ini + dur;
      return;
    }
    if (running == null) {                 // ainda sem âncora
      if (ini != null) running = ini + (dur || 0);   // 1ª com horário fixa a âncora
      return;
    }
    if (dur > 0) {
      const novo = _opHHMM(running);
      if (op.inicio !== novo) { op.inicio = novo; mudou = true; }
      running += dur;
    }
  });
  return mudou;
}

// Aplica a sincronização a TODOS os postos de um dia. O posto "Operador de
// esteira de corte" fica de fora: a máquina é automática e as operações rodam em
// paralelo (mesma exceção da detecção de sobreposição). Só muta o STATE — quem
// chama é que grava. Devolve se algo mudou.
function _opSincronizarHorariosDia(data) {
  let mudou = false;
  _opBlocosDoDia(data).forEach(b => {
    if (ehFuncaoOperadorEsteira(b.nome)) return;
    if (_opSincronizarPosto(b.itens)) mudou = true;
  });
  return mudou;
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
  _opSincronizarHorariosDia(op.data);   // reencaixa os horários na nova ordem
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
// Os grupos em que a sobreposição É conflito, com a mesma regra usada tanto para
// ACUSAR quanto para CORRIGIR — se as duas leituras divergissem, o botão de
// corrigir deixaria selo aceso ou moveria o que não devia.
function _opGruposSobreposicao(lista) {
  const grupos = new Map();
  (lista || []).forEach(op => {
    if (_opInicioMin(op) == null || !_opDuracao(op)) return;
    // Operador de esteira de corte: máquina automática, a pessoa toca duas
    // operações ao mesmo tempo — sobreposição aqui é normal, não conflito.
    if (ehFuncaoOperadorEsteira(_opFuncaoNome(op))) return;
    // Cruzar jornadas é erro tanto no MESMO POSTO (função) quanto na MESMA
    // PESSOA (responsável) — a pessoa não pode estar em dois lugares ao mesmo
    // tempo, mesmo que os postos sejam diferentes. A chave inclui a DATA: o
    // horário é dentro do dia, então comparar dias diferentes acusaria
    // sobreposição onde não há (jornada longa na segunda e outra na terça).
    const chaves = [op.data + '|F|' + (op.funcaoId || _opFuncaoNome(op))];
    // A pessoa é comparada pelo NOME cadastrado na Equipe (normalizado), não pelo
    // id: é o nome que o usuário enxerga, e a mesma pessoa às vezes aparece com
    // id diferente (cadastro repetido) ou só com o nome gravado na operação.
    const pessoa = _normNome(op.responsavelNome || '') || op.responsavelId || '';
    if (pessoa) chaves.push(op.data + '|P|' + pessoa);
    chaves.forEach(k => { if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(op); });
  });
  return grupos;
}

function _opConflitos(lista) {
  const ids = new Set();
  const grupos = _opGruposSobreposicao(lista);
  grupos.forEach(arr => {
    arr.sort((a, b) => _opInicioMin(a) - _opInicioMin(b));
    for (let i = 1; i < arr.length; i++) {
      if (_opInicioMin(arr[i]) < _opFimMin(arr[i - 1])) { ids.add(arr[i].id); ids.add(arr[i - 1].id); }
    }
  });
  return ids;
}

/* ---------------- jornada da fábrica ---------------- */

// Nenhuma operação começa antes das 07:15 nem termina depois das 17:30 — é a
// jornada do setor. Vale como regra de gravação (o modal recusa fora disso), como
// teto do ajuste automático (não empurra operação para depois do expediente) e
// como selo na agenda, para o que já está gravado fora da janela aparecer.
const _OP_JORNADA = { ini: 7 * 60 + 15, fim: 17 * 60 + 30 };

function _opJornadaTexto() {
  return `${_opHHMM(_OP_JORNADA.ini)} às ${_opHHMM(_OP_JORNADA.fim)}`;
}

// Horário calculado pelo programa cai sempre em marca de 5 minutos (07:15, 09:20,
// 13:45…) — é como o chão de fábrica lê o relógio, e um plano cheio de 08:17 e
// 09:53 não se combina em voz alta. Arredonda para CIMA: para baixo, a operação
// começaria ANTES de a anterior terminar, que é o conflito que o ajuste veio
// corrigir. Vale só para o que o programa muda; horário digitado à mão é do
// usuário e fica como ele escreveu.
const _OP_PASSO_MIN = 5;
function _opArredondar(min) {
  const v = Number(min) || 0;
  return Math.ceil(v / _OP_PASSO_MIN) * _OP_PASSO_MIN;
}

// Próximo dia ÚTIL: o que não cabe hoje vai para amanhã, e o que não cabe na
// sexta vai para a segunda — sábado e domingo não recebem operação.
function _opProximoDiaUtil(iso) {
  let d = _expAddDias(iso, 1);
  for (let i = 0; i < 7; i++) {
    const dow = _expData(d).getDay();
    if (dow !== 0 && dow !== 6) return d;
    d = _expAddDias(d, 1);
  }
  return d;
}

// Por que esta operação está fora da jornada (ou null quando está dentro).
function _opForaDaJornada(op) {
  const i = _opInicioMin(op);
  if (i == null) return null;
  const f = i + _opDuracao(op);
  if (i < _OP_JORNADA.ini) return `Começa ${_opHHMM(i)}, antes da abertura (${_opHHMM(_OP_JORNADA.ini)})`;
  if (f > _OP_JORNADA.fim) return `Termina ${_opHHMM(f)}, depois do fim do expediente (${_opHHMM(_OP_JORNADA.fim)})`;
  return null;
}

/* ---------------- sequência obrigatória do lote ---------------- */

// A ORDEM em que um mesmo lote (a OS citada na referência) atravessa os postos.
// É uma corrente física: não se corta um enfesto que ainda está sendo estendido,
// não se separa peça que não foi cortada, não se empacota o que não foi separado.
// O padrão é casado pelo NOME da operação já normalizado (sem acento, minúsculo),
// e a lista é percorrida na ordem escrita: os padrões específicos vêm antes do
// genérico "enfesto", senão "Corte de enfesto" cairia no passo do enfesto.
// São TRÊS correntes que correm em paralelo — o lote atravessa a principal, e as
// outras duas acontecem ao lado, sem depender dela. Passo de correntes
// diferentes nunca é conflito de ordem entre si.
//   principal — do preparo da máquina ao retalho estocado;
//   carga     — o que a expedição faz com os pacotes prontos;
//   materia   — o preparo da matéria prima que alimenta o enfesto.
// A lista é percorrida na ORDEM ESCRITA e o primeiro teste que casa vence: os
// específicos vêm antes dos genéricos ("Corte de enfesto" antes de "Enfesto",
// "Desmontagem de carga" antes de "Montagem de carga", "matéria prima" antes de
// "materiais"). O teste é função, não regex solta, porque quase todo passo se
// identifica por DUAS palavras (o verbo e o objeto) em qualquer ordem.
//
// `porFase` marca os passos que se repetem A CADA FASE DO ENFESTO. Uma OS de
// blusa moletom tricolor tem 5 fases (Corpo 1, Corpo 2, Corpo 3, Forro de capuz,
// Barra/Punhos) e cada uma é um enfesto próprio, com tecido, cor e dimensões
// próprios: ela é estendida, movida, cortada, e as unidades dela são movidas,
// separadas, empacotadas e estocadas — e os retalhos dela, empacotados e
// estocados. São NOVE operações por fase, 45 no tricolor, não nove no lote
// inteiro. Camiseta tricolor e qualquer outra grade multifase seguem a mesma
// regra: quem diz quantas voltas a corrente dá é o cadastro de fases da grade.
// Os passos SEM `porFase` (preparação das máquinas, reposição de materiais) são
// rotina de abertura do dia: acontecem uma vez e servem a todas as fases.
const _OP_SEQUENCIA = [
  { cadeia: 'materia',   ordem: 1,  nome: 'Etapas de preparação de matéria prima', teste: n => /prepar/.test(n) && /materia prima|materia-prima/.test(n) },
  { cadeia: 'principal', ordem: 1,  nome: 'Preparação das máquinas',               teste: n => /prepar/.test(n) && /maquin/.test(n) },
  { cadeia: 'principal', ordem: 2,  nome: 'Reposição de materiais',                teste: n => /reposi/.test(n) && /material|materiais/.test(n) },
  { cadeia: 'carga',     ordem: 1,  nome: 'Seleção de pacotes',                    teste: n => /selec|seleca/.test(n) && /pacote/.test(n) },
  { cadeia: 'carga',     ordem: 2,  nome: 'Desmontagem de carga',                  teste: n => /desmonta/.test(n) },
  { cadeia: 'carga',     ordem: 3,  nome: 'Montagem de carga',                     teste: n => /monta/.test(n) && /carga/.test(n) },
  { cadeia: 'principal', ordem: 4,  nome: 'Movimentação de enfesto',           porFase: true, teste: n => /mover|moviment/.test(n) && /enfesto/.test(n) },
  { cadeia: 'principal', ordem: 5,  nome: 'Corte de enfesto',                  porFase: true, teste: n => /cort/.test(n) && /enfesto/.test(n) },
  { cadeia: 'principal', ordem: 6,  nome: 'Movimentação de unidades cortadas', porFase: true, teste: n => /mover|moviment/.test(n) && /cortad/.test(n) },
  { cadeia: 'principal', ordem: 7,  nome: 'Separação de unidades cortadas',    porFase: true, teste: n => /separa/.test(n) && /cortad/.test(n) },
  { cadeia: 'principal', ordem: 8,  nome: 'Empacotamento de unidades cortadas',porFase: true, teste: n => /empacot/.test(n) && /cortad/.test(n) },
  { cadeia: 'principal', ordem: 9,  nome: 'Estocagem de pacotes',              porFase: true, teste: n => /estoc/.test(n) && /pacote/.test(n) },
  { cadeia: 'principal', ordem: 10, nome: 'Empacotamento de retalhos',         porFase: true, teste: n => /empacot/.test(n) && /retalh/.test(n) },
  { cadeia: 'principal', ordem: 11, nome: 'Estocagem de retalhos',             porFase: true, teste: n => /estoc/.test(n) && /retalh/.test(n) },
  { cadeia: 'principal', ordem: 3,  nome: 'Enfesto',                           porFase: true, teste: n => /enfesto/.test(n) }
];

// Passos da corrente PRINCIPAL na ordem, para dizer o que falta num lote.
const _OP_SEQ_PRINCIPAL = _OP_SEQUENCIA
  .filter(p => p.cadeia === 'principal')
  .slice().sort((a, b) => a.ordem - b.ordem);

// Os que se repetem por fase do enfesto, e as rotinas de abertura que não.
const _OP_SEQ_FASE   = _OP_SEQ_PRINCIPAL.filter(p => p.porFase);
const _OP_SEQ_ROTINA = _OP_SEQ_PRINCIPAL.filter(p => !p.porFase);

// É a operação de ENFESTO propriamente dita (não o corte nem a movimentação dele).
function _opEhEnfesto(op) {
  const p = _opPassoSequencia(op);
  return !!p && p.cadeia === 'principal' && p.nome === 'Enfesto';
}

// Em que passo da corrente esta operação está (ou null quando não é uma delas).
function _opPassoSequencia(op) {
  const n = _normNome(op && op.operacao);
  if (!n) return null;
  return _OP_SEQUENCIA.find(p => p.teste(n)) || null;
}

/* ---------------- fases do enfesto dentro da corrente ---------------- */

// As FASES DO ENFESTO de uma OS, na ordem. Cada fase é um enfesto inteiro e
// separado — tecido, cor e dimensões próprios —, e por isso puxa a sua própria
// volta na corrente principal. A fonte é a OS (que guarda a cópia das fases no
// momento em que foi emitida) e, se ela não tiver, o cadastro da grade. Grade
// sem fases cadastradas devolve UMA fase sem nome: é o caso da peça unicolor, em
// que a corrente dá uma volta só — exatamente o que o programa fazia antes.
function _opFasesDaOS(os) {
  if (!os) return [{ ordem: 1, nome: '' }];
  let fases = Array.isArray(os.fases) && os.fases.length ? os.fases : null;
  if (!fases) {
    const g = (STATE.grades || []).find(x => x.id === _gradeIdDaOS(os));
    if (g && Array.isArray(g.fases) && g.fases.length) fases = g.fases;
  }
  if (!fases) return [{ ordem: 1, nome: '' }];
  return fases
    .map((f, i) => ({ ordem: Number(f.ordem) || i + 1, nome: String(f.nome || '').trim() }))
    .sort((a, b) => a.ordem - b.ordem);
}

// A OS de um número de lote ("436", "0436").
function _opOsDoLote(lote) {
  const semZero = n => String(n).replace(/^0+/, '') || '0';
  const alvo = semZero(lote);
  return (STATE.ordens || []).find(o => semZero(String(o.os || '').trim()) === alvo) || null;
}

// Referência que a operação carrega: OS, modelo e — quando a grade tem mais de
// uma fase — a fase do enfesto a que aquela volta da corrente pertence.
// O "F3/5" é o que faz a agenda, a folha e o conferente do chão de fábrica
// saberem QUAL enfesto está sendo cortado: sem ele, cinco cortes idênticos na
// mesma OS viram cinco linhas indistinguíveis.
function _opRefDaFase(os, fase, totalFases) {
  const base = `${os.os}${os.modeloNome ? ' · ' + os.modeloNome : ''}`;
  if (!fase || totalFases <= 1) return base;
  return `${base} · F${fase.ordem}/${totalFases}${fase.nome ? ' ' + fase.nome : ''}`;
}

// A que fase do enfesto uma operação pertence. Vale o campo gravado; sem ele,
// lê o "F3/5" da referência (operação digitada à mão pode trazer só o texto).
// Sem nenhum dos dois é a PRIMEIRA fase: é o que uma operação planejada antes
// desta regra sempre foi — a corrente dava uma volta só, e essa volta é a fase 1.
// Assim o plano antigo continua fechando em vez de aparecer duplicado.
function _opFaseDaOperacao(op) {
  const n = Number(op && op.faseOrdem);
  if (n > 0) return n;
  const m = String((op && op.referencia) || '').match(/(?:^|[·\s])F(\d+)(?:\s*\/\s*\d+)?\b/i);
  return m ? Number(m[1]) : 1;
}

// Chave da corrente a que a operação pertence dentro de um lote: passos por fase
// se ordenam DENTRO da fase (o corte da fase 2 não vem depois do empacotamento
// da fase 1 — são enfestos diferentes, correm um ao lado do outro); as rotinas
// de abertura ficam fora das fases.
function _opChaveFase(op, passo) {
  return passo && passo.porFase ? String(_opFaseDaOperacao(op)) : 'r';
}

// Tempo já MEDIDO do enfesto de uma fase, na grade da OS. É a média apurada dos
// horários lançados na folha — a única fonte que sabe que "Corpo Parte 3" leva
// mais do que "Barra/Punhos".
function _opTempoMedidoDaFase(os, faseNome) {
  const gradeId = os ? _gradeIdDaOS(os) : '';
  if (!gradeId || !faseNome) return 0;
  const alvo = _normFaseNome(faseNome);
  const l = temposFasesDaGrade(gradeId).find(x => x.n > 0 && _normFaseNome(x.nome) === alvo);
  return (l && l.mediaMin) || 0;
}

/* ------- tempo do enfesto: apurado, nunca cadastrado ------- */

// Cada horário de enfesto já lançado numa folha de OS, com o que explica o
// tamanho dele: de que MODELO e de que GRADE é, que FASE é, quantas CAMADAS
// tinha e quanto media o pano.
function _medicoesEnfesto() {
  const out = [];
  (STATE.ordens || []).forEach(o => {
    const tempos = (o.progresso || {}).enfestosTempos || {};
    const blocos = ((o.enfesto || {}).blocos) || [];
    const camGlobal = parseInt((o.enfesto || {}).camadas, 10) || 0;
    (o.fases || []).forEach(f => {
      const t = tempos[f.ordem] || {};
      const ini = _opMin(t.enfIni), fim = _opMin(t.enfFim);
      if (ini == null || fim == null || fim <= ini) return;
      const bloco = blocos.find(b => String(b.ordem) === String(f.ordem)) || {};
      const camadas = parseInt(bloco.camadas, 10) || camGlobal;
      if (!(camadas > 0)) return;
      out.push({
        modeloId: o.modeloId || '', gradeId: o.gradeId || '',
        faseNome: f.nome || '', minutos: fim - ini, camadas,
        comp: parseFloat(f.comp) || parseFloat(bloco.comp) || 0,
        os: o.os
      });
    });
  });
  return out;
}

// Quanto o enfesto de UMA FASE vai levar, apurado do histórico. Este número
// NUNCA vem do cadastro de Funções: quanto leva um enfesto é do trabalho, não do
// posto — depende do modelo, da grade, de qual fase é e de quantas camadas ela
// tem. Cadastrar "Enfesto = 40 min" no operador de enfestadeira daria o mesmo
// tempo para estender 2,51 m de Corpo Parte 1 e 5,73 m de Corpo Parte 3.
//
// A medição é convertida em TAXA antes de ser aplicada, senão o histórico de uma
// grade não serviria para outra: uma fase medida com 23 camadas não diz nada
// direto sobre a mesma fase com 36. A taxa é por camada e por metro de pano
// quando as duas pontas têm o comprimento; sem ele, por camada.
//
// A amostra é a mais ESPECÍFICA que existir, nesta ordem:
//   1. a mesma fase da MESMA GRADE;
//   2. a mesma fase do MESMO MODELO (outra grade do mesmo tipo de roupa);
//   3. a mesma fase em qualquer OS.
// Devolve { min, n, escopo } — `min` já arredondado na marca de 5.
function _opTempoEnfestoPrevisto(os, fase) {
  const vazio = { min: 0, n: 0, escopo: '' };
  if (!os || !fase) return vazio;
  const alvoNome = _normFaseNome(fase.nome);
  if (!alvoNome) return vazio;
  const blocos = ((os.enfesto || {}).blocos) || [];
  const bloco = blocos.find(b => String(b.ordem) === String(fase.ordem)) || {};
  const camadas = parseInt(bloco.camadas, 10) || parseInt((os.enfesto || {}).camadas, 10) || 0;
  if (!(camadas > 0)) return vazio;
  const comp = parseFloat(fase.comp) || parseFloat(bloco.comp) || 0;

  const daFase = _medicoesEnfesto().filter(m => _normFaseNome(m.faseNome) === alvoNome && m.os !== os.os);
  const escopos = [
    { nome: 'desta grade', lista: daFase.filter(m => m.gradeId && m.gradeId === os.gradeId) },
    { nome: 'deste modelo', lista: daFase.filter(m => m.modeloId && m.modeloId === os.modeloId) },
    { nome: 'de todas as OS', lista: daFase }
  ];
  for (const e of escopos) {
    if (!e.lista.length) continue;
    // Por camada E por metro só quando os dois lados têm o comprimento; senão a
    // conta viraria divisão por zero e o tempo sairia zerado justo nas fases sem
    // medida cadastrada (o forro e a barra/punhos do tricolor).
    const podeMetro = comp > 0 && e.lista.every(m => m.comp > 0);
    const taxas = e.lista.map(m => podeMetro
      ? m.minutos / (m.camadas * m.comp)
      : m.minutos / m.camadas);
    const taxa = taxas.reduce((s, x) => s + x, 0) / taxas.length;
    const bruto = taxa * camadas * (podeMetro ? comp : 1);
    if (!(bruto > 0)) continue;
    return { min: Math.max(5, _opArredondar(bruto)), n: e.lista.length, escopo: e.nome };
  }
  return vazio;
}

// Lotes citados na referência da operação: os números de OS que aparecem no
// texto ("0440", "OS 0430 · 3000 pç", "1042/1051"). Vale o número que É de uma OS
// cadastrada — senão "3000 pç" viraria um lote fantasma e juntaria operações que
// não têm nada a ver. Quando nenhum número bate com OS (quem usa código próprio
// de lote), valem todos. O zero à esquerda cai, para "0440" e "440" serem o mesmo
// lote.
function _opLotesDaOperacao(op) {
  const nums = String((op && op.referencia) || '').match(/\d{3,}/g) || [];
  if (!nums.length) return [];
  const semZero = n => String(n).replace(/^0+/, '') || '0';
  const conhecidos = new Set((STATE.ordens || []).map(o => semZero(String(o.os || '').trim())).filter(Boolean));
  const casados = nums.filter(n => conhecidos.has(semZero(n)));
  return (casados.length ? casados : nums).map(semZero);
}

// Conflitos de ORDEM: dentro do mesmo lote, um passo da corrente que começa
// antes do passo anterior. Dois casos, ambos acusados:
//   • invertida  — começa antes de a anterior COMEÇAR (a ordem está trocada);
//   • encavalada — começa antes de a anterior TERMINAR (o posto seguinte pega o
//     lote que ainda está na mão do anterior).
// Devolve Map id → { nivel, msgs }. Os DOIS níveis são ERRO — a regra da casa é
// que o passo seguinte não começa antes de o anterior terminar naquele lote
// (cortar antes de o enfesto acabar é sempre erro, confirmado com o Junior). O
// nível só distingue o que aconteceu, para o selo dizer qual é o caso:
//   • 'invertida'  — começa antes de a anterior COMEÇAR (corrente ao contrário);
//   • 'encavalada' — começa antes de a anterior TERMINAR.
function _opConflitosOrdem(lista) {
  const porLote = new Map();
  (lista || []).forEach(op => {
    const passo = _opPassoSequencia(op);
    if (!passo || _opInicioMin(op) == null || !_opDuracao(op)) return;
    // A chave inclui a CORRENTE: seleção de pacotes não vem depois do enfesto,
    // são coisas que correm em paralelo. Só passos da mesma corrente se ordenam.
    // A chave inclui a FASE DO ENFESTO: cada fase do tricolor é um enfesto
    // próprio e atravessa a corrente sozinha. Sem isto, o corte da fase 1 (que
    // acontece cedo) contra o enfesto da fase 4 (que acontece tarde) seria
    // acusado de ordem invertida — e o plano correto de uma OS de 5 fases
    // nasceria coberto de selos vermelhos.
    _opLotesDaOperacao(op).forEach(lote => {
      const k = op.data + '|' + passo.cadeia + '|' + lote + '|' + _opChaveFase(op, passo);
      if (!porLote.has(k)) porLote.set(k, []);
      porLote.get(k).push({ op, passo, lote, fase: _opChaveFase(op, passo) });
    });
  });
  const out = new Map();
  const anota = (op, nivel, msg) => {
    let e = out.get(op.id);
    if (!e) { e = { nivel, msgs: [] }; out.set(op.id, e); }
    if (nivel === 'invertida') e.nivel = 'invertida';   // erro manda sobre aviso
    if (!e.msgs.includes(msg)) e.msgs.push(msg);
  };
  porLote.forEach(itens => {
    itens.sort((a, b) => a.passo.ordem - b.passo.ordem || _opInicioMin(a.op) - _opInicioMin(b.op));
    for (let i = 1; i < itens.length; i++) {
      const ant = itens[i - 1], cur = itens[i];
      if (ant.passo.ordem === cur.passo.ordem) continue;
      const iAnt = _opInicioMin(ant.op), fAnt = _opFimMin(ant.op), iCur = _opInicioMin(cur.op);
      const ondeFase = cur.fase !== 'r' ? ` (fase ${cur.fase} do enfesto)` : '';
      const quem = `${cur.passo.nome} (${_opHHMM(iCur)}) × ${ant.passo.nome} (${_opHHMM(iAnt)} → ${_opHHMM(fAnt)}) no lote ${cur.lote}${ondeFase}`;
      if (iCur < iAnt) {
        const msg = `ORDEM INVERTIDA: ${quem}`;
        anota(cur.op, 'invertida', msg); anota(ant.op, 'invertida', msg);
      } else if (iCur < fAnt) {
        const msg = `COMEÇA ANTES DE TERMINAR A ANTERIOR: ${quem}`;
        anota(cur.op, 'encavalada', msg); anota(ant.op, 'encavalada', msg);
      }
    }
  });
  return out;
}

// Selo do conflito de ordem, no formato de cada folha. `tipo` = 'tela' | 'papel'.
// Os dois casos saem em VERMELHO: os dois quebram a corrente do lote.
function _opSeloOrdem(e, tipo) {
  if (!e) return '';
  const rot = e.nivel === 'invertida' ? 'fora de ordem' : 'antes da anterior terminar';
  const t = esc(e.msgs.join(' · '));
  return tipo === 'papel'
    ? ` <span class="tag alto" title="${t}">${rot}</span>`
    : ` <span class="exp-badge alto" title="${t}">${rot}</span>`;
}

/* ---------------- pausas sincronizadas ---------------- */

// Café da manhã, almoço e café da tarde não são trabalho: são paradas da fábrica
// inteira e acontecem no MESMO horário em todas as funções. Reconhecidas pelo
// nome da operação, já normalizado.
function _opTipoPausa(op) {
  const n = _normNome(op && op.operacao);
  if (!n) return '';
  if (/almoc/.test(n)) return 'almoço';
  if (/cafe/.test(n) && /manha/.test(n)) return 'café da manhã';
  if (/cafe/.test(n) && /tarde/.test(n)) return 'café da tarde';
  if (/cafe|lanche|intervalo/.test(n)) return n;   // outra parada: sincroniza pelo próprio nome
  return '';
}
function _opEhPausa(op) { return !!_opTipoPausa(op); }

// Índice das horas marcadas do cadastro de Funções (o campo "todo dia às").
// É cache porque `_opHorarioDeRotina` é perguntado aos milhares por clique — o
// organizador do dia varre o plano inteiro a cada uma das suas 20 passadas —, e
// reler o cadastro a cada pergunta travava a tela. Refeito quando o cadastro
// muda: `_opHoraFixaVersao` sobe no `saveState('funcoes')`, e comparar a
// identidade do array cobre as trocas em bloco (carregar, importar, receber de
// outro aparelho).
let _opHoraFixaVersao = 0;
let _opHoraFixaCache = null;
function _opIndiceHorasFixas() {
  const fs = STATE.funcoes || [];
  if (_opHoraFixaCache && _opHoraFixaCache.fonte === fs && _opHoraFixaCache.versao === _opHoraFixaVersao) {
    return _opHoraFixaCache;
  }
  const porFuncao = new Map();   // "funcaoId|nome normalizado" → hora ('' se não tem)
  const porNome = new Map();     // "nome normalizado" → hora, para quem não diz a função
  fs.forEach(f => {
    _opsDaFuncao(f).forEach(o => {
      const nome = _normNome(o.nome);
      if (!nome) return;
      const hora = String(o.horaFixa || '').trim();
      porFuncao.set(f.id + '|' + nome, hora);
      if (hora && !porNome.has(nome)) porNome.set(nome, hora);
    });
  });
  _opHoraFixaCache = { fonte: fs, versao: _opHoraFixaVersao, porFuncao, porNome };
  return _opHoraFixaCache;
}

// O HORÁRIO FIXO que o cadastro dá a esta operação ('' quando não há).
// Quem decide é a função DA OPERAÇÃO: uma função marcar "Enfesto" às 07:15 não
// pode fixar o enfesto de todo mundo. Só quando a operação não diz de que função
// é (o programa às vezes pergunta apenas pelo nome) a busca se abre para as
// demais — aí o cadastro é a única pista que existe.
function _opHoraFixaCadastrada(op) {
  const alvo = _normNome(op && op.operacao);
  if (!alvo) return '';
  const ix = _opIndiceHorasFixas();
  const id = op && op.funcaoId;
  if (id && (STATE.funcoes || []).some(f => f.id === id)) {
    return ix.porFuncao.get(id + '|' + alvo) || '';
  }
  return ix.porNome.get(alvo) || '';
}

// Operações de HORÁRIO FIXO: acontecem em hora marcada, não entram na fila do
// posto e o organizador não as empurra para resolver conflito de ninguém. São de
// dois tipos:
//   • as que o USUÁRIO marcou com "todo dia às" no cadastro da função — é a
//     forma de dizer isso de qualquer operação, sem depender do nome;
//   • as que o programa reconhece pelo nome desde sempre: as pausas e as rotinas
//     de abertura e fechamento (preparação das máquinas, reposição de materiais
//     e limpeza de ambiente). Ficam aqui para que quem nunca abriu o cadastro
//     continue com o dia funcionando como antes.
// (Diferente do 📌: aqui é o TIPO da operação que manda, não uma escolha caso a
// caso; por isso elas são ancoradas sem aparecer como fixadas pelo usuário.)
function _opHorarioDeRotina(op) {
  if (_opEhPausa(op)) return true;
  if (_opHoraFixaCadastrada(op)) return true;
  const n = _normNome(op && op.operacao);
  if (!n) return false;
  if (/prepar/.test(n) && /maquin/.test(n)) return true;
  if (/reposi/.test(n) && /material|materiais/.test(n)) return true;
  if (/limpeza/.test(n) && /ambiente/.test(n)) return true;
  return false;
}

// Alguma pausa do dia está em horários diferentes entre as funções?
function _opPausasDessincronizadas(lista) {
  const porTipo = new Map();
  (lista || []).forEach(op => {
    const t = _opTipoPausa(op);
    if (!t || _opInicioMin(op) == null) return;
    if (!porTipo.has(t)) porTipo.set(t, new Set());
    porTipo.get(t).add(op.inicio + '|' + _opDuracao(op));
  });
  return Array.from(porTipo.values()).some(s => s.size > 1);
}

// Põe cada tipo de pausa no MESMO horário em todas as funções do dia. A hora de
// referência é a que MAIS funções já usam (o combinado de fato); empatou, vence a
// MAIS TARDE — adiantar a pausa faria a operação que estava terminando às 09:30
// invadir um café que passou a começar 09:25, criando conflito onde havia
// acerto; atrasar só gera espera, que a agenda já mostra como tempo parado. A
// duração segue o mesmo critério, com a MAIOR no empate: pausa sincronizada é a
// mesma janela para todo mundo, senão o refeitório recebe em ondas. A pausa vira
// âncora (inicioFixo) para o encadeamento do posto não a arrastar.
function _opSincronizarPausasDoDia(data) {
  const doDia = (STATE.operacoes || []).filter(o => o.data === data && _opEhPausa(o) && _opInicioMin(o) != null);
  const porTipo = new Map();
  doDia.forEach(op => {
    const t = _opTipoPausa(op);
    if (!porTipo.has(t)) porTipo.set(t, []);
    porTipo.get(t).push(op);
  });
  const ajustadas = [];
  porTipo.forEach(itens => {
    const moda = (valores) => {
      const conta = new Map();
      valores.forEach(v => conta.set(v, (conta.get(v) || 0) + 1));
      return Array.from(conta.entries()).sort((a, b) => b[1] - a[1] || b[0] - a[0])[0][0];
    };
    const ini = _opArredondar(moda(itens.map(_opInicioMin)));
    const dur = _opArredondar(moda(itens.map(_opDuracao).filter(d => d > 0)));
    itens.forEach(op => {
      const mudouIni = _opInicioMin(op) !== ini;
      const mudouDur = dur > 0 && _opDuracao(op) !== dur;
      if (mudouIni || mudouDur) {
        ajustadas.push({ op, de: `${op.inicio} +${_opDuracao(op)}min`, para: `${_opHHMM(ini)} +${dur}min` });
        op.inicio = _opHHMM(ini);
        if (dur > 0) op.duracaoMin = dur;
      }
      op.inicioAuto = true;   // âncora do organizador: a fábrica inteira para nessa hora
    });
  });
  return ajustadas;
}

// Realinha a ORDEM MANUAL de cada posto do dia ao relógio: quem começa antes
// aparece antes. A ordem gravada é o que o encadeamento percorre e o que as setas
// de mover manipulam — deixá-la apontando para uma sequência que os horários já
// desmentiram faz a agenda mostrar 09:30 acima de 08:15 e o encadeamento
// reconstruir a fila errada.
function _opReordenarPostosPorHorario(data) {
  const blocos = _opBlocosDoDia(data);
  blocos.forEach(b => b.itens.sort((x, y) => {
    const ix = _opInicioMin(x), iy = _opInicioMin(y);
    return (ix == null ? 1e9 : ix) - (iy == null ? 1e9 : iy);
  }));
  _opGravarOrdem(blocos);
}

// Tempo MÍNIMO que uma operação precisa, pelo que já foi medido na grade da OS
// citada na referência. Vale para as fases de ENFESTO, que é o que o histórico
// mede (_opTempoMedidoParaOS só marca `aplicavel` nesses casos). É o piso que
// impede o ajuste de conflitos de encurtar um enfesto para ele caber na agenda.
function _opDuracaoNecessaria(op) {
  const os = _opOsDaReferencia(op && op.referencia);
  if (!os) return 0;
  // ENFESTO: o tempo é apurado do histórico e é a fonte única — não existe
  // número cadastrado para comparar. Vale a fase da operação; sem ela, o enfesto
  // da OS não é de fase nenhuma e não há o que estimar.
  if (_opEhEnfesto(op)) {
    const fase = _opFasesDaOS(os).find(f => f.ordem === _opFaseDaOperacao(op));
    return fase ? _opTempoEnfestoPrevisto(os, fase).min : 0;
  }
  const etapas = Array.isArray(op.etapas) ? op.etapas.slice() : (op.etapa ? [op.etapa] : []);
  const r = _opTempoMedidoParaOS(os, op.operacao, etapas, _opFuncaoNome(op));
  return (r && r.aplicavel && r.min > 0) ? r.min : 0;
}

// Corrige os horários que quebram a corrente do lote (cada passo começa quando o
// passo anterior daquele lote TERMINA) e as sobreposições da MESMA PESSOA em
// funções diferentes (ninguém está em dois lugares ao mesmo tempo). Só empurra
// para FRENTE — puxar para trás resolveria o selo e criaria outro problema
// (trabalho antes da hora).
// A operação movida vira âncora (inicioFixo), senão a sincronização do posto a
// traria de volta para logo depois da vizinha de cima; o 📌 na agenda desfaz.
// Cada passada reencaixa os postos, o que pode empurrar outras operações e criar
// novas quebras — daí o laço, que para quando ninguém mais se move.
// Devolve { movidas: [{op, de, para}], travadas: [op] } — travadas são as que só
// caberiam depois da meia-noite, que o campo de horário do dia não representa.
function _opCorrigirOrdemDoDia(data, profundidade = 0) {
  const movidas = new Map(), travadas = new Set(), adiadas = new Map(), ampliadas = new Map();
  const hhmm = min => String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0');
  const destino = _opProximoDiaUtil(data);
  // Primeiro as pausas: elas são o esqueleto do dia (a fábrica inteira para
  // junto). Sincronizadas antes, viram âncoras e o resto do ajuste se encaixa em
  // volta delas, em vez de empurrá-las.
  const pausas = _opSincronizarPausasDoDia(data);
  // As demais rotinas de hora marcada (preparação das máquinas, reposição de
  // materiais, limpeza de ambiente) não são sincronizadas entre funções — cada
  // posto tem a sua —, mas ficam ancoradas: o encadeamento não as arrasta.
  (STATE.operacoes || []).forEach(op => {
    if (op.data === data && _opHorarioDeRotina(op) && _opInicioMin(op) != null) op.inicioAuto = true;
  });
  if (pausas.length) _opSincronizarHorariosDia(data);
  // O que não cabe na jornada vai para o PRÓXIMO DIA ÚTIL, levando junto os
  // passos seguintes do mesmo lote que estavam neste dia — eles dependem dele, e
  // deixá-los para trás poria a costura antes do corte. No dia de destino a
  // operação entra no FIM da fila do posto (sem `ordem` e sem horário fixo), e o
  // encadeamento de lá dá o horário.
  const adiar = op => {
    if (adiadas.has(op.id)) return;
    op.data = destino;
    op.inicio = hhmm(_OP_JORNADA.ini);
    op.inicioFixo = false;
    delete op.inicioAuto;
    delete op.ordem;
    adiadas.set(op.id, { op, de: data, para: destino });
  };
  for (let passada = 0; passada < 20; passada++) {
    const doDia = (STATE.operacoes || []).filter(o => o.data === data);
    let mudou = false;
    // ANTES de mexer em horário: nenhum enfesto encurtado para caber. Se a grade
    // da OS já mostrou que aquela fase leva mais tempo do que o planejado, a
    // duração sobe para o tempo medido. Sem isso o ajuste "resolveria" o conflito
    // mandando o corte começar no fim de um enfesto que, na prática, ainda estaria
    // rodando — o conflito voltaria no chão de fábrica, não na tela.
    doDia.forEach(op => {
      // Duração DIGITADA pelo usuário não é elevada. O tempo medido serve para
      // quem não decidiu; quem decidiu tem motivo — e sem esta linha o ajuste
      // devolvia a medição a cada "Organizar o dia", desfazendo a correção
      // manual tantas vezes quantas o botão fosse clicado.
      if (op.duracaoManual) return;
      // A média medida cai em qualquer minuto (1h27). Sobe para a marca de 5
      // seguinte: nunca menos que o necessário, e o fim da operação também cai
      // numa hora redonda para a próxima encaixar.
      const nec = _opArredondar(_opDuracaoNecessaria(op));
      if (nec <= _opDuracao(op)) return;
      const de = _opDuracao(op);
      op.duracaoMin = nec;
      const reg = ampliadas.get(op.id);
      if (reg) reg.para = nec; else ampliadas.set(op.id, { op, de, para: nec });
      mudou = true;
    });
    // Operação que começa antes da abertura entra NA jornada do próprio dia: o
    // trabalho existe, só estava marcado antes de a fábrica abrir.
    doDia.forEach(op => {
      const ini = _opInicioMin(op);
      if (ini == null || ini >= _OP_JORNADA.ini) return;
      const de = op.inicio;
      op.inicio = hhmm(_OP_JORNADA.ini);
      op.inicioAuto = true;
      const reg = movidas.get(op.id);
      if (reg) reg.para = op.inicio; else movidas.set(op.id, { op, de, para: op.inicio });
      mudou = true;
    });
    const porLote = new Map();
    doDia.forEach(op => {
      const passo = _opPassoSequencia(op);
      if (!passo || _opInicioMin(op) == null || !_opDuracao(op)) return;
      // Uma corrente por lote E POR FASE do enfesto: encadear as cinco fases do
      // tricolor numa fila só empurraria o enfesto da fase 5 para depois da
      // estocagem dos retalhos da fase 1, quando na verdade elas só disputam o
      // posto — e disso cuida a regra de sobreposição, logo abaixo.
      _opLotesDaOperacao(op).forEach(lote => {
        const k = lote + '|' + _opChaveFase(op, passo);
        if (!porLote.has(k)) porLote.set(k, []);
        porLote.get(k).push({ op, passo });
      });
    });
    // Empurra `cur` para começar quando `alvo` manda, registrando o movimento.
    // PAUSA não se move: café e almoço são hora marcada da fábrica inteira, e
    // empurrar um deles para resolver conflito de uma função dessincronizaria
    // todas as outras. O conflito com pausa fica acusado, para o usuário decidir.
    const empurrar = (cur, alvoBruto) => {
      if (_opHorarioDeRotina(cur)) return false;
      const alvo = _opArredondar(alvoBruto);   // 09:22 vira 09:25
      if (_opInicioMin(cur) >= alvo) return false;
      // Não empurra para fora do expediente: o dia acaba às 17:30 e uma operação
      // marcada depois disso é plano que ninguém executa. Fica onde está e sai na
      // lista de "não coube na jornada", para o usuário decidir o que corta.
      if (alvo + _opDuracao(cur) > _OP_JORNADA.fim) { travadas.add(cur); return false; }
      const de = cur.inicio;
      cur.inicio = hhmm(alvo);
      cur.inicioAuto = true;
      const reg = movidas.get(cur.id);
      if (reg) reg.para = cur.inicio; else movidas.set(cur.id, { op: cur, de, para: cur.inicio });
      return true;
    };
    porLote.forEach(itens => {
      itens.sort((a, b) => a.passo.ordem - b.passo.ordem || _opInicioMin(a.op) - _opInicioMin(b.op));
      for (let i = 1; i < itens.length; i++) {
        if (itens[i - 1].passo.ordem === itens[i].passo.ordem) continue;
        const cur = itens[i].op, alvo = _opFimMin(itens[i - 1].op);
        if (_opInicioMin(cur) >= alvo) continue;
        if (alvo + _opDuracao(cur) > _OP_JORNADA.fim) {
          // Não cabe mais hoje: este passo e TODOS os seguintes do lote vão para
          // o próximo dia útil.
          for (let j = i; j < itens.length; j++) adiar(itens[j].op);
          mudou = true;
          break;
        }
        if (empurrar(cur, alvo)) mudou = true;
      }
    });
    // Sobreposição dentro da MESMA FUNÇÃO: o posto é uma pessoa e uma máquina —
    // não faz duas coisas ao mesmo tempo. A que começa DEPOIS espera a outra
    // terminar. Quando a de baixo é rotina de hora marcada (café, almoço,
    // limpeza), quem sai da frente é a de cima: ela recomeça depois da rotina,
    // que não se mexe. Era a última incoerência que ficava só apontada; agora o
    // programa a resolve, como o resto.
    _opGruposSobreposicao(doDia).forEach((arr, chave) => {
      if (chave.indexOf('|F|') < 0) return;
      const lista = arr.slice().sort((a, b) => _opInicioMin(a) - _opInicioMin(b));
      for (let i = 1; i < lista.length; i++) {
        const ant = lista[i - 1], cur = lista[i];
        if (_opInicioMin(cur) >= _opFimMin(ant)) continue;
        if (empurrar(cur, _opFimMin(ant))) { mudou = true; continue; }
        if (empurrar(ant, _opFimMin(cur))) mudou = true;
      }
    });
    // Sobreposição da MESMA PESSOA em funções diferentes: ninguém está em dois
    // lugares ao mesmo tempo, então a segunda operação espera a primeira acabar.
    // Só os grupos de PESSOA entram aqui — cruzamento dentro do mesmo posto é o
    // encadeamento do próprio posto que resolve, e mexer nele aqui atropelaria a
    // ordem que o usuário montou para aquela função.
    _opGruposSobreposicao(doDia).forEach((arr, chave) => {
      if (chave.indexOf('|P|') < 0) return;
      arr.sort((a, b) => _opInicioMin(a) - _opInicioMin(b));
      for (let i = 1; i < arr.length; i++) {
        if (_opInicioMin(arr[i]) < _opFimMin(arr[i - 1]) && empurrar(arr[i], _opFimMin(arr[i - 1]))) mudou = true;
      }
    });
    // Sobrou operação que TERMINA depois do expediente (marcada assim de origem,
    // ou empurrada até aqui): vai inteira para a jornada do próximo dia útil,
    // levando junto os passos seguintes do mesmo lote. É o que fecha a regra —
    // nada fica planejado fora da jornada.
    doDia.forEach(op => {
      const ini = _opInicioMin(op);
      if (ini == null || adiadas.has(op.id)) return;
      if (ini + _opDuracao(op) <= _OP_JORNADA.fim) return;
      if (_opEhPausa(op)) return;                 // pausa fora da janela é caso de cadastro, não de fila
      adiar(op);
      // Os passos seguintes do lote no mesmo dia não podem ficar para trás.
      const passo = _opPassoSequencia(op);
      if (passo) {
        const lotes = _opLotesDaOperacao(op);
        const fase = _opChaveFase(op, passo);
        doDia.forEach(outra => {
          if (outra === op || adiadas.has(outra.id)) return;
          const p2 = _opPassoSequencia(outra);
          if (!p2 || p2.ordem <= passo.ordem) return;
          if (!_opLotesDaOperacao(outra).some(l => lotes.includes(l))) return;
          // Só os passos seguintes da MESMA FASE dependem desta operação: o
          // corte da fase 4 não depende do enfesto da fase 2, e arrastá-lo junto
          // esvaziaria o dia inteiro por causa de um enfesto que atrasou.
          if (_opChaveFase(outra, p2) !== fase) return;
          adiar(outra);
        });
      }
      mudou = true;
    });
    if (!mudou) break;
    // A ordem manual do posto é realinhada ao RELÓGIO antes de encadear. Sem
    // isto, o encadeamento — que anda pela ordem gravada — reencaixa a vizinha
    // logo depois da operação que acabou de ser empurrada e desfaz a correção:
    // as duas ficam se empurrando a cada passada até uma estourar o expediente e
    // ser jogada para o dia seguinte. Era isso que dava a impressão de que
    // "Organizar o dia" não mexia na ordem — ele mexia e o passo seguinte
    // desmanchava.
    _opReordenarPostosPorHorario(data);
    _opSincronizarHorariosDia(data);
    if (adiadas.size) {
      _opReordenarPostosPorHorario(destino);
      _opSincronizarHorariosDia(destino);   // o dia de destino também encaixa
    }
  }
  if (movidas.size || adiadas.size || ampliadas.size || pausas.length) {
    _opReordenarPostosPorHorario(data);
    if (adiadas.size) _opReordenarPostosPorHorario(destino);
  }
  const saida = {
    movidas: Array.from(movidas.values()),
    travadas: Array.from(travadas),
    adiadas: Array.from(adiadas.values()),
    ampliadas: Array.from(ampliadas.values()),
    pausas,
    destino
  };
  // O dia de destino recebeu trabalho: a corrente do lote precisa ser arrumada lá
  // também (as operações chegam todas em 07:15, cada uma na fila do seu posto).
  // Até 3 saltos — se o serviço ainda transbordar depois disso, é falta de dia
  // útil, não de ajuste, e o usuário vê pelos selos.
  if (adiadas.size && profundidade < 3) {
    const seguinte = _opCorrigirOrdemDoDia(destino, profundidade + 1);
    saida.movidas = saida.movidas.concat(seguinte.movidas);
    saida.travadas = saida.travadas.concat(seguinte.travadas);
    saida.adiadas = saida.adiadas.concat(seguinte.adiadas);
    saida.ampliadas = saida.ampliadas.concat(seguinte.ampliadas);
    saida.pausas = saida.pausas.concat(seguinte.pausas);
  }
  return saida;
}

// Botão da agenda: arruma o dia inteiro. Mexe em horário planejado, então
// pergunta antes e diz depois exatamente o que mudou.
async function corrigirOrdemOperacoes(data) {
  if (!exigirAdmin('corrigir os horários das operações')) return;
  const doDia = () => (STATE.operacoes || []).filter(o => o.data === data);
  // Quantas sobreposições são da MESMA PESSOA (as que este ajuste resolve): as do
  // mesmo posto ficam de fora, então não podem entrar na conta nem na promessa.
  const pessoaCruzada = lista => {
    const ids = new Set();
    _opGruposSobreposicao(lista).forEach((arr, chave) => {
      if (chave.indexOf('|P|') < 0) return;
      arr.sort((a, b) => _opInicioMin(a) - _opInicioMin(b));
      for (let i = 1; i < arr.length; i++) {
        if (_opInicioMin(arr[i]) < _opFimMin(arr[i - 1])) { ids.add(arr[i].id); ids.add(arr[i - 1].id); }
      }
    });
    return ids;
  };
  const ordemAntes = _opConflitosOrdem(doDia()).size;
  const pessoaAntes = pessoaCruzada(doDia()).size;
  const pausasFora = _opPausasDessincronizadas(doDia());
  const foraJornada = doDia().filter(o => _opForaDaJornada(o)).length;
  const mesmaFuncao = _opSobreposicoesMesmaFuncao(doDia()).length;
  // Operações de hora marcada do cadastro que este dia ainda não tem. Sem esta
  // conta, quem cadastrasse o "todo dia às" depois de o dia já estar planejado
  // teria de alocar a OS de novo só para a rotina entrar.
  const fixasFaltando = (STATE.funcoes || []).reduce((n, f) => n + _opsDaFuncao(f).filter(o => {
    const ini = _opMin(String(o.horaFixa || '').trim());
    if (!String(o.nome || '').trim() || ini == null) return false;
    if (ini < _OP_JORNADA.ini || ini + (Number(o.duracaoMin) || 0) > _OP_JORNADA.fim) return false;
    return !doDia().some(x => x.funcaoId === f.id && _normNome(x.operacao) === _normNome(o.nome));
  }).length, 0);
  // Operações sem dono, num posto que tem uma pessoa só: elas deixam a agenda
  // sem o vínculo entre os postos que a mesma pessoa ocupa.
  const semDono = doDia().filter(o => !o.responsavelId && !o.responsavelNome
    && _opResponsavelDoPosto(_opFuncaoNome(o))).length;
  if (!ordemAntes && !pessoaAntes && !pausasFora && !foraJornada && !mesmaFuncao && !fixasFaltando && !semDono) {
    return toast('O dia já está organizado: sem conflitos, dentro da jornada e com as pausas sincronizadas', 'ok');
  }
  const destinoPrev = _opProximoDiaUtil(data);
  if (!confirm('Organizar os horários deste dia?\n\n'
    + `· ${ordemAntes} operação(ões) fora da ordem do lote\n`
    + `· ${pessoaAntes} com a mesma pessoa em duas funções ao mesmo tempo\n`
    + `· pausas ${pausasFora ? 'em horários diferentes entre as funções' : 'já sincronizadas'}\n`
    + `· ${foraJornada} fora da jornada (${_opJornadaTexto()})\n`
    + `· ${mesmaFuncao} par(es) sobrepostos dentro da mesma função\n`
    + `· ${fixasFaltando} operação(ões) de horário fixo do cadastro que faltam neste dia\n`
    + `· ${semDono} sem responsável, em posto que tem uma pessoa só\n\n`
    + 'As de horário fixo entram na hora cadastrada e o resto se encaixa em volta delas.\n'
    + 'Operação sem dono recebe a pessoa do posto — é o que faz a linha do tempo de uma função mostrar o que a mesma pessoa faz na outra.\n'
    + 'Cada uma passa a começar quando a anterior termina — só para frente, nunca para trás.\n'
    + `O que não couber até ${_opHHMM(_OP_JORNADA.fim)} passa para ${formatDate(destinoPrev)}, junto com os passos seguintes do mesmo lote.\n`
    + 'Enfesto planejado com menos tempo do que a grade da OS já mostrou ter recebe a duração medida.\n'
    + 'Café da manhã, almoço e café da tarde ficam no mesmo horário em todas as funções.\n'
    + 'Nenhuma função fica com duas operações ao mesmo tempo.')) return;
  // A rotina de hora marcada entra ANTES do ajuste: ela é âncora, e o resto do
  // dia é que se encaixa em volta dela.
  const _fx = _opAplicarHorariosFixosNoDia(data);
  const fixasNovas = _fx.criadas.concat(_fx.marcadas);
  const donosNovos = _opPreencherResponsaveisDoDia(data);
  const { movidas, travadas, adiadas, ampliadas, pausas, destino } = _opCorrigirOrdemDoDia(data);
  if (!movidas.length && !adiadas.length && !ampliadas.length && !pausas.length
      && !fixasNovas.length && !donosNovos.length) {
    return toast('Nada a mover: os conflitos não se resolvem empurrando para frente', 'err');
  }
  await saveState('operacoes');
  const restou = _opConflitosOrdem(doDia()).size + pessoaCruzada(doDia()).size;
  toast(`${movidas.length} operação(ões) reencaixada(s)`
    + (donosNovos.length ? ` · ${donosNovos.length} com o responsável do posto` : '')
    + (fixasNovas.length ? ` · ${fixasNovas.length} de horário fixo incluída(s)` : '')
    + (pausas.length ? ` · ${pausas.length} pausa(s) sincronizada(s)` : '')
    + (ampliadas.length ? ` · ${ampliadas.length} com a duração medida da grade` : '')
    + (adiadas.length ? ` · ${adiadas.length} passada(s) para ${formatDate(destino)}` : '')
    + (travadas.length ? ` · ${travadas.length} sem encaixe` : '')
    + (restou ? ` · ainda ${restou} em conflito` : ''), restou || travadas.length ? 'err' : 'ok');
  renderOperacoes();
}

// O que FALTA para o lote atravessar a corrente principal inteira.
//
// Duas coisas que a versão anterior errava e faziam a lista cobrar o que não
// devia:
//   • olhava SÓ o dia aberto. Passo do mesmo lote planejado em OUTRA data (o
//     enfesto de ontem, o empacotamento que o organizador passou para amanhã)
//     aparecia como ausente. Agora vale o plano inteiro: o passo existe, e a
//     linha só diz em que dia ele está;
//   • cobrava passo que não CABE mais no dia. Uma corrente que começou às 16h
//     não fecha até as 17:30, e insistir nela é pedir o impossível. Só entra em
//     `faltam` — a lista com botão de incluir — o passo que ainda cabe hoje; o
//     resto vai para `naoCabem`, que a tela mostra como continuação no próximo
//     dia útil, sem botão.
// As rotinas de hora marcada (preparação das máquinas, reposição de materiais)
// contam como atendidas se acontecem no dia, com ou sem referência ao lote: são
// feitas uma vez e servem a todos os lotes.
//
// A conta é POR FASE DO ENFESTO. Uma OS de blusa moletom tricolor tem 5 fases e
// cada uma precisa das suas 9 operações; cobrar 11 passos do lote inteiro dizia
// "sequência completa" com 4/5 do serviço sem planejar. Cada fase vira uma linha
// própria (as rotinas de abertura entram na primeira, porque são do dia).
// Devolve [{ lote, fase, faseNome, nFases, rotulo, feitos, faltam, naoCabem,
// noutroDia, total }].
function _opLotesIncompletos(doDia, data) {
  const dia = data || (doDia && doDia[0] && doDia[0].data) || '';
  const lotesDoDia = new Set();
  const rotinaNoDia = new Set();
  (doDia || []).forEach(op => {
    const passo = _opPassoSequencia(op);
    if (!passo || passo.cadeia !== 'principal') return;
    if (_opHorarioDeRotina(op)) rotinaNoDia.add(passo.ordem);
    _opLotesDaOperacao(op).forEach(l => lotesDoDia.add(l));
  });
  // Onde cada passo de cada lote/fase está, em QUALQUER data do plano.
  const ondeEsta = new Map();   // "lote|fase" → Map(ordem → {data, fim})
  (STATE.operacoes || []).forEach(op => {
    const passo = _opPassoSequencia(op);
    if (!passo || passo.cadeia !== 'principal') return;
    _opLotesDaOperacao(op).forEach(l => {
      if (!lotesDoDia.has(l)) return;
      const k = l + '|' + _opChaveFase(op, passo);
      if (!ondeEsta.has(k)) ondeEsta.set(k, new Map());
      const m = ondeEsta.get(k);
      if (!m.has(passo.ordem) || String(op.data) < String(m.get(passo.ordem).data)) {
        m.set(passo.ordem, { data: op.data, fim: _opFimMin(op) });
      }
    });
  });
  const saida = [];
  lotesDoDia.forEach(lote => {
    const fases = _opFasesDaOS(_opOsDoLote(lote));
    fases.forEach((fase, idx) => {
      const onde = ondeEsta.get(lote + '|' + fase.ordem) || new Map();
      // As rotinas de abertura são do DIA, não da fase: entram na primeira.
      const passos = idx === 0 ? _OP_SEQ_ROTINA.concat(_OP_SEQ_FASE) : _OP_SEQ_FASE;
      const faltam = [], naoCabem = [], noutroDia = [];
      // O que falta é uma FILA: cada passo consome o tempo do seguinte. Medir um
      // a um contra o mesmo horário diria que todos cabem — foi o que fazia a
      // lista cobrar dez operações num dia que só comporta uma.
      let relogio = null, estourou = false;
      passos.forEach(p => {
        // Rotina do dia: quem responde por ela é a fase 1, e o mapa dela está na
        // corrente 'r' (fora das fases).
        const mapa = p.porFase ? onde : (ondeEsta.get(lote + '|r') || new Map());
        const existente = mapa.get(p.ordem);
        if (existente) {
          if (existente.data !== dia) noutroDia.push({ passo: p, data: existente.data });
          else if (existente.fim != null) relogio = Math.max(relogio == null ? 0 : relogio, existente.fim);
          return;
        }
        if (rotinaNoDia.has(p.ordem)) return;
        if (estourou) { naoCabem.push(p); return; }   // o anterior não coube: este também não
        // A simulação usa as MESMAS regras da cascata: duração real do passo e o
        // primeiro vão livre do posto. `_opSugestaoPasso` já devolve as duas
        // coisas — e o `cabe` dele diz se ainda dá para pôr no dia.
        const s = _opSugestaoPasso(dia, lote, p, fase.ordem);
        const ini = Math.max(_opMin(s.inicio) || 0, relogio == null ? 0 : relogio);
        if (!s.cabe || ini + (s.duracaoMin || 0) > _OP_JORNADA.fim) { estourou = true; naoCabem.push(p); return; }
        relogio = ini + (s.duracaoMin || 0);
        faltam.push(p);
      });
      const pendentes = faltam.length + naoCabem.length;
      saida.push({
        lote, faltam, naoCabem, noutroDia,
        fase: fase.ordem, faseNome: fase.nome, nFases: fases.length,
        rotulo: `OS ${lote}` + (fases.length > 1
          ? ` · F${fase.ordem}/${fases.length}${fase.nome ? ' ' + fase.nome : ''}` : ''),
        feitos: passos.length - pendentes,
        total: passos.length
      });
    });
  });
  return saida.sort((a, b) => (b.faltam.length + b.naoCabem.length) - (a.faltam.length + a.naoCabem.length)
    || String(a.lote).localeCompare(String(b.lote), undefined, { numeric: true })
    || a.fase - b.fase);
}

/* ---------------- alocar uma OS: a cascata de operações do lote ---------------- */

// Põe no dia as operações que o cadastro de Funções marcou com HORÁRIO FIXO
// ("todo dia às") e que ainda não estão lá.
//
// Elas são INDEPENDENTES das OS alocadas: café, almoço, preparação das máquinas,
// limpeza do fim do expediente são da JORNADA, não do lote. Por isso entram UMA
// VEZ por função e por dia — não uma por OS, nem uma por fase do enfesto —, e
// nascem SEM referência a OS nenhuma. Sem referência elas ficam fora de toda
// corrente de lote: não são cobradas no quadro "o que falta", não entram na
// checagem de ordem e alocar uma segunda OS não as duplica. O que a corrente do
// lote faz com elas é só desviar: a fila do posto se encaixa em volta da hora
// marcada.
// Devolve { criadas, marcadas } — `marcadas` são as que já estavam no dia e
// passaram a exibir o 📌 de horário fixo.
function _opAplicarHorariosFixosNoDia(data) {
  if (!Array.isArray(STATE.operacoes)) STATE.operacoes = [];
  const criadas = [], marcadas = [];
  const noDia = (STATE.operacoes || []).filter(o => o.data === data);
  (STATE.funcoes || []).forEach(f => {
    _opsDaFuncao(f).forEach(o => {
      const nome = String(o.nome || '').trim();
      const hora = String(o.horaFixa || '').trim();
      const ini = _opMin(hora);
      if (!nome || ini == null) return;
      // Já está no dia (posta pelo cadastro, pela cascata ou à mão): não repete.
      // Mas garante o 📌 — o cadastro diz que aquela operação é de hora marcada,
      // e a agenda tem que mostrá-la como fixa (barra preta) em vez de deixá-la
      // com cara de operação comum, que qualquer reencaixe empurra. A HORA não é
      // reescrita: se alguém moveu aquele dia à mão, foi de propósito.
      const jaLa = noDia.filter(x => x.funcaoId === f.id && _normNome(x.operacao) === _normNome(nome));
      if (jaLa.length) {
        jaLa.forEach(x => { if (!x.inicioFixo) { x.inicioFixo = true; marcadas.push(x); } });
        return;
      }
      const dur = Math.max(0, Math.round(Number(o.duracaoMin) || 0));
      // Hora cadastrada fora da jornada é erro de cadastro, não plano do dia:
      // criar a operação aqui só encheria a agenda de selo "fora da jornada".
      if (ini < _OP_JORNADA.ini || ini + dur > _OP_JORNADA.fim) return;
      const pessoa = _opResponsavelDoPosto(f.nome);
      const nova = {
        id: uid(), data,
        funcaoId: f.id, funcaoNome: f.nome,
        operacao: nome, escopo: 'completa', etapas: [],
        inicio: hora, duracaoMin: dur,
        // HORÁRIO FIXO de verdade, não âncora do organizador: quem decidiu a
        // hora foi o usuário, no cadastro da função ("todo dia às"). Por isso sai
        // com o 📌 e em PRETO na linha do tempo, como qualquer hora travada à
        // mão — e o encadeamento do posto não a reescreve.
        inicioFixo: true, inicioAuto: true,
        responsavelId: pessoa ? pessoa.id : '', responsavelNome: pessoa ? pessoa.nome : '',
        // Sem referência DE PROPÓSITO: esta operação não é de OS nenhuma.
        referencia: '', status: 'pendente', prioridade: 'eletiva', obs: ''
      };
      STATE.operacoes.push(nova);
      noDia.push(nova);
      criadas.push(nova);
    });
  });
  return { criadas, marcadas };
}

// Põe o responsável nas operações do dia que estão sem ele, quando o posto tem
// UMA pessoa só na Equipe. Operação sem dono deixa o programa cego para a única
// coisa que liga dois postos — a pessoa —, e é isso que faz a linha do tempo de
// uma função não mostrar o que a mesma pessoa está fazendo na outra.
// Devolve as que receberam nome.
function _opPreencherResponsaveisDoDia(data) {
  const tocadas = [];
  (STATE.operacoes || []).forEach(op => {
    if (op.data !== data || op.responsavelId || op.responsavelNome) return;
    const pessoa = _opResponsavelDoPosto(_opFuncaoNome(op));
    if (!pessoa) return;
    op.responsavelId = pessoa.id;
    op.responsavelNome = pessoa.nome;
    tocadas.push(op);
  });
  return tocadas;
}

// As operações de hora marcada existem em TODO dia planejado, não só naquele em
// que alguém alocou uma OS: café, almoço, preparação das máquinas e limpeza
// acontecem todo dia, tenha ou não OS no dia. Ao abrir a agenda, elas são postas
// nos dias ÚTEIS do período mostrado.
//
// De HOJE PARA FRENTE apenas. Dia passado é registro do que aconteceu, e
// carimbar um almoço em retrospecto seria inventar histórico — quem quiser
// completar um dia que já passou usa "Organizar o dia", que pede confirmação.
// Devolve as criadas.
function _opGarantirHorariosFixosNoPeriodo(ini, fim) {
  const hoje = _expHoje();
  // HORIZONTE: dois meses à frente. Sem ele, navegar a agenda até 2027 encheria
  // a base de café e almoço de um ano inteiro — e o plano de verdade não vai
  // tão longe. Quem planejar além disso aloca a OS, e a rotina entra junto.
  const horizonte = _expAddDias(hoje, 60);
  const ate = fim < horizonte ? fim : horizonte;
  const criadas = [], marcadas = [];
  let d = ini < hoje ? hoje : ini;
  for (let i = 0; i < 62 && d <= ate; i++, d = _expAddDias(d, 1)) {
    const dow = _expData(d).getDay();
    if (dow === 0 || dow === 6) continue;   // sábado e domingo não são planejados
    const r = _opAplicarHorariosFixosNoDia(d);
    criadas.push(...r.criadas);
    marcadas.push(...r.marcadas);
  }
  if (criadas.length) {
    Array.from(new Set(criadas.map(o => o.data))).forEach(dia => {
      _opReordenarPostosPorHorario(dia);
      _opSincronizarHorariosDia(dia);
    });
  }
  return criadas.concat(marcadas);
}

// Primeiro horário, a partir de `piso`, em que o posto fica livre por `dur`
// minutos seguidos. Existe por causa das operações de hora marcada: sem ele, o
// almoço das 11:30 empurraria a corrente inteira para as 13:00 e uma limpeza
// marcada para as 17:00 esvaziaria a tarde toda — quando na verdade o serviço
// cabe ANTES delas. Agora a fila se encaixa em volta da hora marcada.
function _opPrimeiroVagoNoPosto(data, funcaoId, piso, dur) {
  const ocupadas = (STATE.operacoes || [])
    .filter(o => o.data === data && o.funcaoId === funcaoId && _opInicioMin(o) != null && _opDuracao(o) > 0)
    .map(o => ({ ini: _opInicioMin(o), fim: _opFimMin(o) }))
    .sort((a, b) => a.ini - b.ini);
  let t = _opArredondar(Math.max(piso == null ? _OP_JORNADA.ini : piso, _OP_JORNADA.ini));
  for (let i = 0; i < ocupadas.length; i++) {
    const o = ocupadas[i];
    if (o.fim <= t) continue;              // já passou
    if (o.ini >= t + dur) break;           // cabe inteira antes desta
    t = _opArredondar(o.fim);              // colide: começa quando ela termina
  }
  return t;
}

// Qual FUNÇÃO executa cada passo da corrente, e com que nome e tempo. A fonte é o
// cadastro: cada função tem as suas operações (com o tempo de cada uma), e é o
// nome da operação que diz a que passo ela pertence. Sem cadastro para um passo,
// cai no histórico do plano — quem já fez aquele passo antes.
function _opFuncaoDoPasso(passo) {
  const doCadastro = [];
  (STATE.funcoes || []).forEach(f => {
    _opsDaFuncao(f).forEach(o => {
      const p = _opPassoSequencia({ operacao: o.nome });
      if (p && p.cadeia === passo.cadeia && p.ordem === passo.ordem) {
        doCadastro.push({ funcaoId: f.id, funcaoNome: f.nome, nome: o.nome, duracaoMin: Number(o.duracaoMin) || 0 });
      }
    });
  });
  // Mais de uma função cadastrando o mesmo passo: vence a que tem tempo definido
  // (é a que alguém realmente configurou); empatando, a primeira do cadastro.
  if (doCadastro.length) {
    return doCadastro.slice().sort((a, b) => (b.duracaoMin > 0) - (a.duracaoMin > 0))[0];
  }
  const iguais = (STATE.operacoes || []).filter(o => {
    const p = _opPassoSequencia(o);
    return p && p.cadeia === passo.cadeia && p.ordem === passo.ordem;
  });
  if (!iguais.length) return null;
  const funcaoId = _opModa(iguais.map(o => o.funcaoId).filter(Boolean));
  const f = (STATE.funcoes || []).find(x => x.id === funcaoId);
  return {
    funcaoId: funcaoId || '', funcaoNome: (f && f.nome) || _opFuncaoNome(iguais[0]),
    nome: _opModa(iguais.map(o => String(o.operacao || '').trim()).filter(Boolean)) || passo.nome,
    duracaoMin: _opModa(iguais.map(o => _opDuracao(o)).filter(d => d > 0)) || 0
  };
}

// Monta no dia as operações que faltam para a OS atravessar a corrente inteira,
// UMA VOLTA POR FASE DO ENFESTO. Numa blusa moletom tricolor são 5 fases (Corpo
// 1, Corpo 2, Corpo 3, Forro de capuz, Barra/Punhos), e cada uma é um enfesto
// próprio que precisa das suas nove operações: ser estendida, movida, cortada,
// ter as unidades movidas, separadas, empacotadas e estocadas, e os retalhos
// empacotados e estocados. Antes o programa montava NOVE operações no lote
// inteiro — o plano nascia com 1/5 do serviço do dia, e as outras quatro fases
// eram digitadas à mão ou simplesmente esqueciam.
//
// Cada fase é uma corrente independente: dentro dela, um passo começa quando o
// anterior termina; entre fases, o que separa uma da outra é o POSTO — a
// enfestadeira só estende uma de cada vez, e a fila do posto já cuida disso.
// Por isso a fase 2 pode estar sendo cortada enquanto a 1 está sendo separada,
// que é como a fábrica trabalha de verdade.
//
// Passo que a fase já tem não é recriado, e as rotinas de hora marcada
// (preparação das máquinas, reposição de materiais) entram uma vez só no dia:
// são de abertura, não de fase.
// Devolve { criadas, semFuncao, naoCoube, fases }.
function _opMontarCascataDoLote(data, os) {
  const lote = String(os.os || '').trim().replace(/^0+/, '') || String(os.os || '').trim();
  const doDia = () => (STATE.operacoes || []).filter(o => o.data === data);
  const criadas = [], semFuncao = [], naoCoube = [];
  const fases = _opFasesDaOS(os);
  // Antes da corrente, as operações de hora marcada do cadastro: elas são o
  // esqueleto do dia e a fila do lote se encaixa em volta delas.
  const fixas = _opAplicarHorariosFixosNoDia(data).criadas;
  const jaTem = new Set();     // "fase|ordem" que o lote já tem ("r|ordem" nas rotinas)
  const rotinaNoDia = new Set();
  doDia().forEach(op => {
    const p = _opPassoSequencia(op);
    if (!p || p.cadeia !== 'principal') return;
    if (_opHorarioDeRotina(op)) rotinaNoDia.add(p.ordem);
    if (_opLotesDaOperacao(op).includes(lote)) jaTem.add(_opChaveFase(op, p) + '|' + p.ordem);
  });
  // Cria um passo da corrente. `fase` é null nas rotinas de abertura.
  const criar = (passo, fase, pisoRelogio) => {
    const chave = (fase ? String(fase.ordem) : 'r') + '|' + passo.ordem;
    if (jaTem.has(chave)) return null;
    const cad = _opFuncaoDoPasso(passo);
    if (!cad || !cad.funcaoId) { semFuncao.push({ passo, fase }); return null; }
    if (_opHorarioDeRotina({ operacao: cad.nome }) && rotinaNoDia.has(passo.ordem)) return null;
    // Duração: a cadastrada na função — MENOS no enfesto, que é apurado do
    // histórico e não do cadastro. "Corpo Parte 3" (5,73 m) não leva o mesmo que
    // "Corpo Parte 2" (1,13 m), e o posto não tem como saber a diferença.
    let duracaoMin = cad.duracaoMin;
    if (fase && _opEhEnfesto({ operacao: cad.nome })) {
      duracaoMin = _opTempoEnfestoPrevisto(os, fase).min;
    }
    // A operação da corrente NUNCA herda o horário fixo do cadastro: a hora
    // marcada é do DIA, não do lote. Se ela mandasse aqui, marcar "Enfesto" às
    // 07:15 grudaria os cinco enfestos do tricolor no mesmo minuto. Quem entra
    // na hora cadastrada é a operação independente, criada em
    // `_opAplicarHorariosFixosNoDia` — esta aqui apenas se encaixa em volta dela.
    //
    // Quando o posto fica livre: depois da última operação de TRABALHO já marcada
    // nele. As de hora marcada ficam fora desta conta — elas são âncoras no meio
    // do dia, não o fim da fila —, e a colisão com elas é resolvida logo abaixo,
    // procurando o primeiro vago.
    const fimDoPosto = doDia()
      .filter(o => o.funcaoId === cad.funcaoId && _opInicioMin(o) != null && !_opHorarioDeRotina(o))
      .reduce((mx, o) => Math.max(mx, _opFimMin(o)), _OP_JORNADA.ini);
    const piso = Math.max(pisoRelogio == null ? _OP_JORNADA.ini : pisoRelogio, fimDoPosto);
    const ini = _opPrimeiroVagoNoPosto(data, cad.funcaoId, piso, duracaoMin);
    if (ini + duracaoMin > _OP_JORNADA.fim) { naoCoube.push({ passo, fase }); return null; }
    const pessoa = _opResponsavelDoPosto(cad.funcaoNome);
    const nova = {
      id: uid(), data,
      funcaoId: cad.funcaoId, funcaoNome: cad.funcaoNome,
      operacao: cad.nome, escopo: 'completa', etapas: [],
      inicio: _opHHMM(ini), duracaoMin,
      // Âncora do programa: o horário veio da corrente do LOTE, que atravessa
      // vários postos. Sem ancorar, o encadeamento de cada posto reescreveria
      // tudo pela fila dele e a corrente se quebraria no mesmo instante em que
      // foi montada — o corte caía depois da separação. Não é 📌 do usuário: ele
      // segue livre para mudar o que quiser.
      inicioAuto: true,
      responsavelId: pessoa ? pessoa.id : '', responsavelNome: pessoa ? pessoa.nome : '',
      referencia: fase ? _opRefDaFase(os, fase, fases.length) : _opRefDaFase(os, null, 1),
      status: 'pendente', prioridade: 'eletiva', obs: ''
    };
    // A fase fica GRAVADA, não só escrita na referência: é ela que diz a qual
    // corrente a operação pertence quando o usuário reescreve o texto do campo.
    if (fase && fases.length > 1) { nova.faseOrdem = fase.ordem; nova.faseNome = fase.nome; }
    STATE.operacoes.push(nova);
    criadas.push(nova);
    return ini + duracaoMin;
  };
  // Rotinas de abertura primeiro: elas abrem o dia e o resto conta a partir dali.
  let pisoDoDia = null;
  _OP_SEQ_ROTINA.forEach(passo => {
    const fim = criar(passo, null, pisoDoDia);
    if (fim != null) pisoDoDia = fim;
  });
  // Uma volta completa da corrente por fase do enfesto.
  fases.forEach(fase => {
    let relogio = pisoDoDia;   // toda fase começa depois da abertura do dia
    let travou = false;        // um passo não coube: os seguintes dependem dele
    _OP_SEQ_FASE.forEach(passo => {
      // Passo que não cabe no dia leva junto o resto da corrente DAQUELA fase.
      // Sem isto, um enfesto de 5h45 que não cabia deixava o corte, a separação
      // e o empacotamento da mesma fase marcados assim mesmo — a fase aparecia
      // sendo cortada sem nunca ter sido estendida.
      if (travou) { naoCoube.push({ passo, fase }); return; }
      const antes = naoCoube.length;
      const fim = criar(passo, fase, relogio);
      if (naoCoube.length > antes) { travou = true; return; }
      if (fim != null) relogio = fim;
    });
  });
  if (criadas.length || fixas.length) {
    _opReordenarPostosPorHorario(data);
    _opSincronizarHorariosDia(data);
  }
  return { criadas, semFuncao, naoCoube, fases, fixas };
}

// Modal: escolher a OS que abre (ou continua) o planejamento do dia.
function abrirModalAlocarOS(data) {
  if (!exigirAdmin('planejar operações')) return;
  const dia = data || opPlanoAncora || _expHoje();
  const noDia = new Set();
  (STATE.operacoes || []).filter(o => o.data === dia)
    .forEach(o => _opLotesDaOperacao(o).forEach(l => noDia.add(l)));
  const semZero = n => String(n).replace(/^0+/, '') || '0';
  // Candidatas: OS em produção (dentro do fluxo), com a que já está no dia
  // marcada — alocar de novo só acrescenta o que falta, e é bom deixar isso claro.
  const ordens = (STATE.ordens || [])
    .filter(o => (o.os || '').toString().trim())
    .sort((a, b) => String(b.os).localeCompare(String(a.os), undefined, { numeric: true }))
    .slice(0, 120);
  document.getElementById('modal-aloc-title').textContent = `Alocar OS em ${formatDate(dia)}`;
  document.getElementById('modal-aloc-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Dia *</label><input type="date" id="aloc-data" value="${esc(dia)}"></div>
      <div class="field full">
        <label>OS *</label>
        <select id="aloc-os">
          <option value="">— selecione —</option>
          ${ordens.map(o => `<option value="${esc(o.id)}">${esc(o.os)}${o.modeloNome ? ' · ' + esc(o.modeloNome) : ''}${
            noDia.has(semZero(o.os)) ? ' (já está no dia)' : ''}</option>`).join('')}
        </select>
        <div class="field-hint">O programa monta as operações que faltam para esta OS atravessar a sequência inteira, cada uma na função que a executa, com o tempo cadastrado ali. Alocar uma segunda OS acrescenta a sequência dela <b>depois</b> do que cada posto já tem — a ordem entre as OS é você que decide, movendo as operações.</div>
        <div class="field-hint" id="aloc-fases"></div>
      </div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;">
      A sequência é montada <b>uma vez por fase do enfesto</b> da grade: cada fase é um enfesto próprio e tem as suas nove operações (enfesto, mover, cortar, mover as unidades, separar, empacotar, estocar, empacotar os retalhos e estocar os retalhos). Preparação das máquinas e reposição de materiais entram uma vez só, na abertura do dia.<br>
      Os tempos vêm do cadastro de <b>Funções</b>; o enfesto de cada fase usa o tempo já medido na grade quando ele for maior. Operação sem tempo cadastrado entra com duração zero — dá para ajustar depois, uma a uma.
    </div>`;
  // Mostra quantas voltas a corrente vai dar assim que a OS é escolhida: 5 fases
  // do tricolor são 45 operações, e é melhor saber disso antes de clicar.
  const selOS = document.getElementById('aloc-os');
  const box = document.getElementById('aloc-fases');
  const mostrarFases = () => {
    if (!box) return;
    const o = (STATE.ordens || []).find(x => x.id === (selOS && selOS.value));
    if (!o) { box.innerHTML = ''; return; }
    const fs = _opFasesDaOS(o);
    box.innerHTML = `<b>${fs.length} fase(s) de enfesto</b>: ${esc(fs.map(f => `F${f.ordem}${f.nome ? ' ' + f.nome : ''}`).join(' · '))}`
      + ` → até <b>${_OP_SEQ_ROTINA.length + fs.length * _OP_SEQ_FASE.length} operações</b> na sequência completa.`;
  };
  if (selOS) selOS.onchange = mostrarFases;
  mostrarFases();
  openModal('modal-aloc-os');
}

async function confirmarAlocarOSNoDia() {
  if (!exigirAdmin('planejar operações')) return;
  const data = document.getElementById('aloc-data')?.value || '';
  const osId = document.getElementById('aloc-os')?.value || '';
  if (!data) return toast('Escolha o dia', 'err');
  const os = (STATE.ordens || []).find(o => o.id === osId);
  if (!os) return toast('Escolha a OS', 'err');
  if (!Array.isArray(STATE.operacoes)) STATE.operacoes = [];
  const { criadas, semFuncao, naoCoube, fases, fixas } = _opMontarCascataDoLote(data, os);
  if (!criadas.length && !fixas.length && !semFuncao.length && !naoCoube.length) {
    return toast(`OS ${os.os} já tem a sequência completa em ${formatDate(data)}`, 'ok');
  }
  await saveState('operacoes');
  closeModal('modal-aloc-os');
  const rotuloPasso = x => x.passo.nome + (x.fase && fases.length > 1 ? ` (F${x.fase.ordem}${x.fase.nome ? ' ' + x.fase.nome : ''})` : '');
  // Operação sem tempo entra com duração zero e some da linha do tempo. Com uma
  // volta da corrente isso passava batido; com cinco, um enfesto sem tempo
  // cadastrado esconde cinco enfestos — o aviso tem que dizer QUAIS faltam.
  const semTempo = Array.from(new Set(criadas.filter(o => !_opDuracao(o)).map(o => o.operacao)));
  toast(`${criadas.length} operação(ões) criada(s) para a OS ${os.os}`
    + (fases.length > 1 ? ` em ${fases.length} fases de enfesto` : '')
    + (fixas.length ? ` · + ${fixas.length} de horário fixo do dia (não são da OS)` : '')
    + (semTempo.length ? ` · sem tempo cadastrado: ${semTempo.join(', ')}` : '')
    + (semFuncao.length ? ` · ${semFuncao.length} sem função cadastrada` : '')
    + (naoCoube.length ? ` · ${naoCoube.length} não coube(m) no dia` : ''),
    (semFuncao.length || naoCoube.length || semTempo.length) ? 'err' : 'ok');
  if (semFuncao.length) {
    console.info('Passos sem função cadastrada:', semFuncao.map(rotuloPasso).join(', '));
  }
  if (naoCoube.length) {
    console.info('Passos que não couberam na jornada:', naoCoube.map(rotuloPasso).join(', '));
  }
  opPlanoAncora = data;
  try { sessionStorage.setItem('gos:op:ancora', opPlanoAncora); } catch (e) {}
  renderOperacoes();
}

/* ---------------- retirar uma OS do planejamento (ação em massa) ---------------- */

// As operações planejadas de um lote, opcionalmente só num dia. É o conjunto que
// a retirada apaga — e o mesmo que o modal conta antes de perguntar.
// Rotinas de hora marcada (café, almoço, preparação das máquinas) NÃO entram:
// elas são da jornada, não da OS, e nascem sem referência a OS nenhuma.
function _opOperacoesDoLote(lote, data) {
  const alvo = String(lote).replace(/^0+/, '') || String(lote);
  return (STATE.operacoes || []).filter(o =>
    (!data || o.data === data) && _opLotesDaOperacao(o).includes(alvo));
}

// Modal: escolher a OS que sai do planejamento e se sai de um dia só ou de todos.
function abrirModalDesalocarOS(data) {
  if (!exigirAdmin('planejar operações')) return;
  const dia = data || opPlanoAncora || _expHoje();
  const semZero = n => String(n).replace(/^0+/, '') || '0';
  // Só OS que REALMENTE têm operação planejada: oferecer as outras seria
  // prometer um desfazer que não desfaz nada.
  const lotes = new Map();   // lote → { noDia, total }
  (STATE.operacoes || []).forEach(o => {
    _opLotesDaOperacao(o).forEach(l => {
      const e = lotes.get(l) || { noDia: 0, total: 0 };
      e.total++;
      if (o.data === dia) e.noDia++;
      lotes.set(l, e);
    });
  });
  if (!lotes.size) {
    return toast('Nenhuma OS com operações planejadas para retirar', 'err');
  }
  const linhas = Array.from(lotes.entries())
    .map(([lote, e]) => {
      const os = _opOsDoLote(lote);
      return { lote, e, os, rot: `${os ? os.os : lote}${os && os.modeloNome ? ' · ' + os.modeloNome : ''}` };
    })
    .sort((a, b) => (b.e.noDia > 0) - (a.e.noDia > 0)
      || String(b.lote).localeCompare(String(a.lote), undefined, { numeric: true }));
  document.getElementById('modal-desaloc-title').textContent = 'Retirar OS do planejamento';
  document.getElementById('modal-desaloc-fields').innerHTML = `
    <div class="form-grid cols-2">
      <div class="field"><label>Dia *</label><input type="date" id="desaloc-data" value="${esc(dia)}" onchange="abrirModalDesalocarOS(this.value)"></div>
      <div class="field full">
        <label>OS *</label>
        <select id="desaloc-os" onchange="_opDesalocResumo()">
          <option value="">— selecione —</option>
          ${linhas.map(l => `<option value="${esc(l.lote)}" data-nodia="${l.e.noDia}" data-total="${l.e.total}">${
            esc(l.rot)} — ${l.e.noDia} em ${esc(formatDate(dia))}${l.e.total > l.e.noDia ? ` · ${l.e.total} no plano inteiro` : ''}</option>`).join('')}
        </select>
        <div class="field-hint">Aparecem só as OS que têm operação planejada. Café, almoço e as demais rotinas de hora marcada <b>não</b> são retiradas: elas são da jornada, não da OS.</div>
      </div>
      <div class="field full">
        <label>Alcance *</label>
        <select id="desaloc-escopo" onchange="_opDesalocResumo()">
          <option value="dia">Só o dia ${esc(formatDate(dia))}</option>
          <option value="tudo">Todos os dias do plano</option>
        </select>
        <div class="field-hint">A corrente de uma OS costuma transbordar para o próximo dia útil. <b>Todos os dias</b> tira a OS do planejamento inteiro de uma vez.</div>
      </div>
    </div>
    <div class="info-box" style="margin-top:8px;font-size:12px;" id="desaloc-resumo">Escolha a OS para ver quantas operações serão retiradas.</div>`;
  openModal('modal-desaloc-os');
  _opDesalocResumo();
}

// Resumo do que vai sair, atualizado a cada escolha do modal.
function _opDesalocResumo() {
  const box = document.getElementById('desaloc-resumo');
  if (!box) return;
  const lote = document.getElementById('desaloc-os')?.value || '';
  const data = document.getElementById('desaloc-data')?.value || '';
  const escopo = document.getElementById('desaloc-escopo')?.value || 'dia';
  if (!lote) { box.textContent = 'Escolha a OS para ver quantas operações serão retiradas.'; return; }
  const alvo = _opOperacoesDoLote(lote, escopo === 'dia' ? data : '');
  if (!alvo.length) { box.innerHTML = '<b>Nenhuma operação</b> desta OS neste alcance.'; return; }
  const porDia = new Map();
  const postos = new Set();
  alvo.forEach(o => {
    porDia.set(o.data, (porDia.get(o.data) || 0) + 1);
    postos.add(_opFuncaoNome(o));
  });
  const feitas = alvo.filter(o => _opStatus(o) === 'feita').length;
  box.innerHTML = `Serão retiradas <b>${alvo.length}</b> operação(ões) da OS <b>${esc(lote)}</b>, em ${postos.size} posto(s):<br>`
    + Array.from(porDia.entries()).sort().map(([d, n]) => `${esc(formatDate(d))}: <b>${n}</b>`).join(' · ')
    + (feitas ? `<br><span style="color:var(--alert);"><b>Atenção:</b> ${feitas} já ${feitas === 1 ? 'está marcada' : 'estão marcadas'} como <b>feita</b> — retirar apaga o registro de que ${feitas === 1 ? 'foi executada' : 'foram executadas'}.</span>` : '');
}

async function confirmarDesalocarOS() {
  if (!exigirAdmin('planejar operações')) return;
  const lote = document.getElementById('desaloc-os')?.value || '';
  const data = document.getElementById('desaloc-data')?.value || '';
  const escopo = document.getElementById('desaloc-escopo')?.value || 'dia';
  if (!lote) return toast('Escolha a OS', 'err');
  const alvo = _opOperacoesDoLote(lote, escopo === 'dia' ? data : '');
  if (!alvo.length) return toast('Nenhuma operação desta OS neste alcance', 'err');
  const feitas = alvo.filter(o => _opStatus(o) === 'feita').length;
  const dias = Array.from(new Set(alvo.map(o => o.data))).sort();
  if (!confirm(`Retirar ${alvo.length} operação(ões) da OS ${lote} do planejamento?\n\n`
    + `· ${escopo === 'dia' ? formatDate(data) : dias.map(formatDate).join(', ')}\n`
    + (feitas ? `· ${feitas} já marcada(s) como FEITA — o registro de execução se perde\n` : '')
    + '\nAs rotinas de hora marcada (café, almoço, preparação das máquinas) ficam: são da jornada, não da OS.\n'
    + 'Isto não se desfaz — para recolocar, use "+ Alocar OS".')) return;
  const ids = new Set(alvo.map(o => o.id));
  STATE.operacoes = (STATE.operacoes || []).filter(o => !ids.has(o.id));
  // Os dias que perderam operação são reencaixados: quem ficou sobe na fila do
  // posto em vez de deixar o buraco da OS que saiu.
  dias.forEach(d => { _opReordenarPostosPorHorario(d); _opSincronizarHorariosDia(d); });
  await saveState('operacoes');
  closeModal('modal-desaloc-os');
  toast(`${alvo.length} operação(ões) da OS ${lote} retirada(s) de ${dias.length === 1 ? formatDate(dias[0]) : dias.length + ' dias'}`, 'ok');
  renderOperacoes();
}

// Quem ocupa um posto, para a operação criada pelo programa já nascer com o
// responsável. É a pessoa da Equipe cuja função principal é esta — e só quando
// há UMA: com duas (Marina e Roseane em "Auxiliar de costura #1") quem escolhe é
// o usuário, e chutar poria o nome errado na agenda.
//
// Sem responsável, a operação fica sem dono e o programa perde a única coisa que
// liga dois postos: a PESSOA. Quem trabalha em duas funções (o João na
// enfestadeira e na esteira, o Denis na produção e na expedição) tem uma jornada
// só, e é o responsável que faz a linha do tempo de um posto mostrar o que a
// mesma pessoa está fazendo no outro.
function _opResponsavelDoPosto(funcaoNome) {
  const { dentro } = _opPessoasDaFuncao(funcaoNome);
  return dentro.length === 1 ? dentro[0] : null;
}

// Valor que mais se repete numa lista (moda); empate fica com o primeiro.
function _opModa(valores) {
  const conta = new Map();
  (valores || []).forEach(v => conta.set(v, (conta.get(v) || 0) + 1));
  const top = Array.from(conta.entries()).sort((a, b) => b[1] - a[1])[0];
  return top ? top[0] : null;
}

// Onde um passo que FALTA se encaixaria: começa quando o passo anterior daquele
// lote termina (é o buraco exato que ele preenche) e, sem passo anterior, na
// hora do passo seguinte ou na abertura. Posto e duração vêm do histórico: quem
// costuma fazer aquele passo e quanto costuma levar — assim o modal já abre com
// o dia coerente, e não com um formulário em branco no meio da corrente.
// `fase` é a fase do enfesto a que o passo pertence: o buraco a preencher é o da
// corrente DAQUELA fase — o corte da fase 3 entra depois do enfesto da fase 3,
// não depois do enfesto da fase 1.
function _opSugestaoPasso(data, lote, passo, fase) {
  const alvo = String(lote);
  const chaveFase = passo.porFase ? String(fase || 1) : 'r';
  const doLote = (STATE.operacoes || [])
    .filter(o => o.data === data && _opLotesDaOperacao(o).includes(alvo))
    .map(o => ({ op: o, p: _opPassoSequencia(o) }))
    .filter(x => x.p && x.p.cadeia === passo.cadeia && _opInicioMin(x.op) != null
      && _opChaveFase(x.op, x.p) === chaveFase);
  const antes = doLote.filter(x => x.p.ordem < passo.ordem).sort((a, b) => b.p.ordem - a.p.ordem)[0];
  const depois = doLote.filter(x => x.p.ordem > passo.ordem).sort((a, b) => a.p.ordem - b.p.ordem)[0];
  let ini = antes ? _opFimMin(antes.op) : (depois ? _opInicioMin(depois.op) : _OP_JORNADA.ini);
  if (ini == null) ini = _OP_JORNADA.ini;
  // Mesma marca de 5 minutos que o ajuste usa: a sugestão do botão não pode
  // propor 09:22 se o organizador poria 09:25.
  ini = Math.max(_OP_JORNADA.ini, Math.min(_opArredondar(ini), _OP_JORNADA.fim));
  const iguais = (STATE.operacoes || []).filter(o => {
    const p = _opPassoSequencia(o);
    return p && p.cadeia === passo.cadeia && p.ordem === passo.ordem;
  });
  // O POSTO sai do cadastro, como na cascata; o histórico é só o reserva.
  const cad = _opFuncaoDoPasso(passo);
  const funcaoId = (cad && cad.funcaoId)
    || _opModa(iguais.map(o => o.funcaoId).filter(Boolean))
    || (antes && antes.op.funcaoId) || (depois && depois.op.funcaoId) || '';
  // NOME como o planejamento escreve. O nome do passo é o rótulo interno da
  // corrente ("Preparação das máquinas"); no plano ele pode se chamar "Etapas de
  // preparação das máquinas do setor". Quem lê a agenda tem que ver a operação
  // com o nome que ela tem na casa — e a operação criada pelo botão nasce com
  // esse mesmo texto, senão o dia acumularia duas grafias para a mesma coisa.
  const nome = (cad && cad.nome)
    || _opModa(iguais.map(o => String(o.operacao || '').trim()).filter(Boolean)) || passo.nome;
  // DURAÇÃO pelas mesmas regras da cascata. O enfesto é apurado do histórico
  // DAQUELA FASE; os demais vêm do cadastro da função. Antes valia a moda das
  // durações já planejadas, e ela dizia 1h30 para um enfesto de Barra/Punhos que
  // leva 5h45 — a fila do "o que falta" então achava que tudo cabia no dia.
  const os = _opOsDoLote(lote);
  const faseObj = (os && passo.porFase) ? _opFasesDaOS(os).find(f => f.ordem === Number(fase || 1)) : null;
  let duracaoMin = 0;
  if (os && faseObj && _opEhEnfesto({ operacao: (cad && cad.nome) || passo.nome })) {
    duracaoMin = _opTempoEnfestoPrevisto(os, faseObj).min;
  } else {
    duracaoMin = (cad && cad.duracaoMin) || _opModa(iguais.map(o => _opDuracao(o)).filter(d => d > 0)) || 0;
  }
  // HORÁRIO em que o posto realmente tem esse vão livre. Sem isto a sugestão
  // propunha 07:15 para um posto ocupado o dia inteiro, e incluir criava
  // sobreposição na hora.
  if (funcaoId) ini = _opPrimeiroVagoNoPosto(data, funcaoId, ini, duracaoMin);
  return {
    inicio: _opHHMM(ini),
    funcaoId,
    nome,
    duracaoMin,
    cabe: ini + duracaoMin <= _OP_JORNADA.fim
  };
}

// Botão "+" do quadro de lotes: abre o modal já com o passo que falta, no
// horário do buraco, no posto que costuma fazê-lo e com a OS — e a FASE DO
// ENFESTO — na referência. Sem a fase no texto, a operação criada aqui entraria
// na corrente da fase 1 e o quadro seguiria cobrando o passo da fase certa.
function incluirOperacaoFaltante(data, lote, cadeia, ordem, fase) {
  if (!exigirAdmin('planejar operações')) return;
  const passo = _OP_SEQUENCIA.find(p => p.cadeia === cadeia && p.ordem === Number(ordem));
  if (!passo) return;
  const nFase = Number(fase) || 1;
  const s = _opSugestaoPasso(data, lote, passo, nFase);
  const os = _opOsDoLote(lote);
  const fases = os ? _opFasesDaOS(os) : [];
  const faseObj = passo.porFase ? fases.find(f => f.ordem === nFase) : null;
  const ref = os ? _opRefDaFase(os, faseObj, fases.length) : String(lote);
  abrirModalOperacao('', data, s.funcaoId, {
    operacao: s.nome, inicio: s.inicio, duracaoMin: s.duracaoMin, referencia: ref,
    faseOrdem: faseObj && fases.length > 1 ? faseObj.ordem : 0,
    faseNome: faseObj && fases.length > 1 ? faseObj.nome : ''
  });
}

/* ---------------- vazios (tempo sem operação) ---------------- */

// Jornada REAL de um conjunto de operações: da primeira hora de início ao último
// fim. Diferente de _opJanelaDoDia, que arredonda para horas cheias porque é
// eixo de desenho — aqui interessa o tempo de verdade.
function _opJornada(ops) {
  let ini = null, fim = null;
  (ops || []).forEach(op => {
    const i = _opInicioMin(op);
    if (i == null || !_opDuracao(op)) return;
    const f = i + _opDuracao(op);
    if (ini == null || i < ini) ini = i;
    if (fim == null || f > fim) fim = f;
  });
  return ini == null ? null : { ini, fim };
}

// Janelas SEM operação dentro de uma jornada. Usada em dois níveis:
//   • por FUNÇÃO (jornada = a do dia): quando aquele posto fica parado enquanto
//     o dia corre — é o "espaço vazio" que o planejamento precisa enxergar;
//   • do DIA inteiro (itens = todas as operações): quando NENHUMA função está
//     operando, que é onde a continuidade entre as funções se rompe.
// `tipo` diz onde o vazio cai: antes da primeira operação, entre duas, ou depois
// da última. `minimo` filtra ruído de arredondamento.
function _opVazios(itens, jornada, minimo = 1) {
  if (!jornada) return [];
  const blocos = (itens || [])
    .filter(o => _opInicioMin(o) != null && _opDuracao(o) > 0)
    .map(o => ({ i: _opInicioMin(o), f: _opFimMin(o) }))
    .sort((a, b) => a.i - b.i);
  if (!blocos.length) {
    const m = jornada.fim - jornada.ini;
    return m >= minimo ? [{ ini: jornada.ini, fim: jornada.fim, min: m, tipo: 'todo' }] : [];
  }
  const out = [];
  let cursor = jornada.ini;
  blocos.forEach(b => {
    if (b.i - cursor >= minimo) {
      out.push({ ini: cursor, fim: b.i, min: b.i - cursor, tipo: cursor === jornada.ini ? 'antes' : 'entre' });
    }
    cursor = Math.max(cursor, b.f);
  });
  if (jornada.fim - cursor >= minimo) {
    out.push({ ini: cursor, fim: jornada.fim, min: jornada.fim - cursor, tipo: 'depois' });
  }
  return out;
}

// Vazio menor que isto é ruído de arredondamento do plano, não ociosidade: não
// vira linha na agenda nem na folha.
const _OP_VAZIO_MIN = 5;

// O que a MESMA PESSOA está fazendo em OUTRA função durante um vazio. A linha do
// vazio dizia que o posto parou, mas não por quê — e o motivo mais comum é a
// pessoa ter sido puxada para outra função. Com isso, o vazio de um posto e a
// ocupação da pessoa aparecem na mesma linha.
function _opPessoaEmOutraFuncao(v, pessoa, funcaoNome, doDia, max = 2) {
  const alvo = _normNome(pessoa || '');
  if (!alvo) return [];
  const daFuncao = _normNome(funcaoNome || '');
  return (doDia || [])
    .filter(o => _normNome(_opFuncaoNome(o)) !== daFuncao
      && _normNome(_opResponsavelNome(o)) === alvo
      && _opInicioMin(o) != null && _opDuracao(o) > 0)
    // Horário RECORTADO no vazio: a operação da outra função quase nunca começa
    // e termina junto com o buraco, e mostrar a janela inteira dela faria parecer
    // que a pessoa esteve fora o tempo todo. O que interessa é a parte que cai
    // dentro do vazio — e só quando ela é relevante (>= o mesmo piso do vazio).
    .map(o => ({
      funcao: _opFuncaoNome(o), operacao: o.operacao || '—',
      ini: Math.max(_opInicioMin(o), v.ini), fim: Math.min(_opFimMin(o), v.fim)
    }))
    .filter(x => x.fim - x.ini >= _OP_VAZIO_MIN)
    .sort((a, b) => a.ini - b.ini)
    .slice(0, max);
}

// "Corte de enfesto (Operador de esteira) 09:50→10:00 · ..."
function _opOcupacaoTexto(lista) {
  return (lista || [])
    .map(x => `${x.operacao} (${x.funcao}) ${_opHHMM(x.ini)}→${_opHHMM(x.fim)}`)
    .join(' · ');
}

// Texto curto de um vazio para a linha da agenda e da folha.
function _opVazioTexto(v) {
  const quando = `${_opHHMM(v.ini)} → ${_opHHMM(v.fim)}`;
  if (v.tipo === 'todo') return `${quando} · ${_opDurTexto(v.min)} sem nenhuma operação`;
  if (v.tipo === 'antes') return `${quando} · ${_opDurTexto(v.min)} parada antes de começar`;
  if (v.tipo === 'depois') return `${quando} · ${_opDurTexto(v.min)} parada até o fim do dia`;
  return `${quando} · ${_opDurTexto(v.min)} sem operação`;
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

// Impede que a criação das operações de hora marcada, feita ao desenhar a
// agenda, entre em laço: ela grava e manda redesenhar, e o redesenho chamaria a
// criação de novo.
let _opFixasEmCurso = false;

function renderOperacoes() {
  const cont = document.getElementById('operacoes-painel');
  if (!cont) return;
  const { ini, fim } = _expRange(opPlanoModo, opPlanoAncora);
  // A rotina de hora marcada aparece já posta em todo dia útil do período — ela
  // é da jornada e não depende de haver OS alocada. Só quem edita cria dados,
  // e só uma vez: a segunda passada não acha nada para criar.
  if (!_opFixasEmCurso && currentRole === 'admin') {
    _opFixasEmCurso = true;
    try {
      const novas = _opGarantirHorariosFixosNoPeriodo(ini, fim);
      if (novas.length) {
        saveState('operacoes')
          .then(() => { _opFixasEmCurso = false; renderOperacoes(); })
          .catch(() => { _opFixasEmCurso = false; });
        return;
      }
    } catch (e) { console.warn('_opGarantirHorariosFixosNoPeriodo', e); }
    _opFixasEmCurso = false;
  }
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
      <div class="exp-seg" title="Como agrupar a linha do tempo">
        <button class="${opVista === 'posto' ? 'active' : ''}" onclick="opSetVista('posto')">Por posto</button>
        <button class="${opVista === 'pessoa' ? 'active' : ''}" onclick="opSetVista('pessoa')">Por pessoa</button>
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;">
        <button class="btn accent" onclick="goto('print-operacoes')">🖨 Folha do plano</button>
      </div>
      <div class="admin-only" style="display:flex;gap:6px;flex-wrap:wrap;">
        <button class="btn primary" title="O dia começa pela OS: o programa monta a sequência de operações dela." onclick="abrirModalAlocarOS(opPlanoAncora)">+ Alocar OS no dia</button>
        ${(STATE.operacoes || []).some(o => _opLotesDaOperacao(o).length)
          ? `<button class="btn" title="Tira do planejamento, de uma vez, todas as operações de uma OS alocada — em um dia ou no plano inteiro. É o desfazer do «Alocar OS»." onclick="abrirModalDesalocarOS(opPlanoAncora)">− Retirar OS</button>` : ''}
        <button class="btn" onclick="abrirModalOperacao()">+ Operação avulsa</button>
        ${opPlanoModo === 'dia' ? `<button class="btn" onclick="copiarOperacoesDoDiaAnterior()" title="Repete no dia mostrado a jornada do último dia planejado antes dele. Copia como pendente e não duplica o que já existe.">⧉ Repetir dia anterior</button>` : ''}
      </div>
    </div>`;

  const conflitos = _opConflitos(ops);
  const foraDeOrdem = _opConflitosOrdem(ops);
  const foraJornadaN = ops.filter(o => _opForaDaJornada(o)).length;
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
      <div class="item ${foraDeOrdem.size ? 'alerta' : ''}"><div class="num">${fmt(foraDeOrdem.size)}</div><div class="lbl">Fora de ordem</div></div>
      <div class="item ${foraJornadaN ? 'alerta' : ''}" title="Jornada do setor: ${esc(_opJornadaTexto())}"><div class="num">${fmt(foraJornadaN)}</div><div class="lbl">Fora da jornada</div></div>
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
    return `<div class="op-regua"><div class="op-regua-lbl" title="Barra preta = operação com horário fixo (📌), em qualquer função">Horário do dia</div><div class="op-regua-eixo">${ticks.join('')}</div></div>`;
  };

  // Barra da operação dentro da janela do dia. É onde "07:12 por 3h20" vira
  // uma coisa que se enxerga ao lado dos outros postos.
  const barraHtml = (op, jan, comFuncao, cor) => {
    const i = _opInicioMin(op), dur = _opDuracao(op);
    if (i == null || !dur) return '';
    const larg = jan.fim - jan.ini;
    const left = (i - jan.ini) / larg * 100;
    const width = Math.max(1.2, dur / larg * 100);
    const st = _opStatus(op);
    const pr = _opPrioridade(op);
    const conf = conflitos.has(op.id) ? ' conflito' : '';
    const rot = comFuncao ? `${_opFuncaoNome(op)}: ${op.operacao}` : op.operacao;
    // Cor própria da operação (identifica a etapa na linha de tempo). O status
    // "feita" fica esmaecido por CSS e o conflito ganha contorno vermelho — os
    // dois convivem com a cor de fundo sem escondê-la.
    const bg = cor ? `background:${cor};` : '';
    return `<div class="op-bar ${st} prio-${pr}${conf}" style="left:${left.toFixed(3)}%;width:${width.toFixed(3)}%;${bg}"
      title="${esc(rot)} · ${esc(_opJanelaTexto(op))}${conf ? ' · SOBREPOSTA (mesma pessoa/posto em dois horários)' : ''}"><span>${esc(rot)}</span></div>`;
  };

  const linhaHtml = (op, pos, qtd) => {
    const st = _opStatus(op);
    const pr = _opPrioridade(op);
    const resp = _opResponsavelNome(op);
    const conflito = conflitos.has(op.id);
    const ordemErr = foraDeOrdem.get(op.id) || null;   // corrente do lote quebrada
    const foraJornada = _opForaDaJornada(op);          // fora do expediente
    // O selo distingue as duas naturezas: processo inteiro do posto (o padrão,
    // sem selo) e etapa avulsa planejada à parte.
    const _opEtapas = Array.isArray(op.etapas) ? op.etapas : (op.etapa ? [op.etapa] : []);
    const selo = op.escopo === 'etapa'
      ? ` <span class="exp-badge info" title="Planejada por etapas da função${_opEtapas.length ? ': ' + _opEtapas.join(', ') : ''}">${_opEtapas.length > 1 ? 'etapas' : 'etapa'}</span>`
      : '';
    const selopr = pr === 'eletiva' ? '' : ` <span class="op-prio ${pr}">${esc(_OP_PRIORIDADE[pr].lbl)}</span>`;
    return `
      <div class="op-row prio-${pr} ${st === 'feita' ? 'feita' : ''}">
        <span class="admin-only op-mover">
          <button title="Subir esta operação no posto" onclick="moverOperacao('${esc(op.id)}',-1)" ${pos === 0 ? 'disabled' : ''}>▲</button>
          <button title="Descer esta operação no posto" onclick="moverOperacao('${esc(op.id)}',1)" ${pos === qtd - 1 ? 'disabled' : ''}>▼</button>
        </span>
        <span class="janela">${esc(_opJanelaTexto(op))}${op.inicioFixo
          ? ` <button type="button" class="op-fixo" onclick="soltarHorarioOperacao('${esc(op.id)}')" title="Horário fixo: definido à mão, não reencaixa após a anterior. Clique para voltar ao encaixe automático.">📌</button>`
          : ''}</span>
        <span class="oper">${esc(op.operacao) || '(sem descrição)'}${selopr}${selo}${conflito ? ' <span class="exp-badge alto" title="Este posto tem outra operação no mesmo horário">sobreposta</span>' : ''}${
          _opSeloOrdem(ordemErr, 'tela')}${
          foraJornada ? ` <span class="exp-badge alto" title="${esc(foraJornada)}. Jornada do setor: ${esc(_opJornadaTexto())}">fora da jornada</span>` : ''}${op.obs ? ` <span class="obs">· ${esc(op.obs)}</span>` : ''}</span>
        <span class="resp">${esc(resp) || '<span class="obs">a definir</span>'}</span>
        <span class="ref">${esc(op.referencia) || ''}</span>
        <button type="button" class="exp-badge ${_OP_STATUS[st].cls} op-status" onclick="alternarStatusOperacao('${esc(op.id)}')" title="Clique para mudar: pendente → em andamento → feita">${esc(_OP_STATUS[st].lbl)}</button>
        <span class="admin-only op-acoes">
          <button title="Editar esta operação" onclick="abrirModalOperacao('${esc(op.id)}')">✎</button>
          <button title="Duplicar esta operação" onclick="duplicarOperacao('${esc(op.id)}')">⧉</button>
          <button title="Excluir esta operação" onclick="excluirOperacao('${esc(op.id)}')">×</button>
        </span>
      </div>`;
  };

  // Um bloco por FUNÇÃO dentro do dia: as funções correm em paralelo, então
  // cada uma tem a sua faixa própria no mesmo eixo de horas.
  const diaHtml = (data, doDia) => {
    const jan = _opJanelaDoDia(doDia);
    const cores = _opMapaCores(doDia);
    // A mesma estrutura que as setas de mover manipulam — desenhar a partir dela
    // garante que a ordem vista é a ordem gravada.
    const grupos = _opBlocosDoDia(data);
    // Jornada real do dia: é contra ela que se mede o tempo em que cada função
    // fica parada enquanto as outras trabalham.
    const jornadaDia = _opJornada(doDia);
    // Faixa hachurada do vazio na linha do tempo da função.
    const vazioBarra = (v, j) => {
      const larg = j.fim - j.ini;
      const left = (v.ini - j.ini) / larg * 100;
      const width = Math.max(0.6, v.min / larg * 100);
      return `<div class="op-bar-vazio" style="left:${left.toFixed(3)}%;width:${width.toFixed(3)}%;"
        title="${esc(_opVazioTexto(v))}"></div>`;
    };
    // A linha do vazio diz o buraco E, quando a pessoa daquele posto está em
    // outra função nesse intervalo, o que ela foi fazer.
    // Barra SOMBRA: a operação que a MESMA PESSOA faz em OUTRA função, desenhada
    // na faixa deste posto. Quem trabalha em duas funções tem uma jornada só, e
    // vê-la partida em duas linhas do tempo escondia justamente o encontro das
    // duas — agora cada linha mostra as duas, e o cruzamento salta.
    const barraSombra = (op, jan, cor) => {
      const i = _opInicioMin(op), dur = _opDuracao(op);
      if (i == null || !dur) return '';
      const larg = jan.fim - jan.ini;
      const left = (i - jan.ini) / larg * 100;
      const width = Math.max(0.8, dur / larg * 100);
      // COR DE ORIGEM: a mesma da função de onde a operação vem, para se
      // reconhecer de qual quadro ela é sem precisar ler o rótulo. O que a
      // distingue do trabalho DESTE posto é o contorno tracejado e o tom mais
      // claro, não a cor.
      const pintura = cor ? `background:${cor};border-color:${cor};` : '';
      return `<div class="op-bar-outra" style="left:${left.toFixed(3)}%;width:${width.toFixed(3)}%;${pintura}"
        title="${esc(_opResponsavelNome(op))} está em ${esc(_opFuncaoNome(op))}: ${esc(op.operacao) || '—'} · ${esc(_opJanelaTexto(op))}"
        ><span>${esc(_opFuncaoNome(op))}</span></div>`;
    };
    const vazioLinha = (v, pessoa, funcaoNome) => {
      const ocup = _opPessoaEmOutraFuncao(v, pessoa, funcaoNome, doDia);
      return `
      <div class="op-row op-vazio">
        <span class="op-mover admin-only"></span>
        <span class="janela">${esc(_opHHMM(v.ini))} → ${esc(_opHHMM(v.fim))}</span>
        <span class="oper">— ${esc(_opDurTexto(v.min))} sem operação —${ocup.length
          ? ` <span class="obs">${esc(pessoa)} em: ${esc(_opOcupacaoTexto(ocup))}</span>`
          : ''}</span>
      </div>`;
    };

    const blocos = grupos.map((g, gi) => {
      const minutos = g.itens.reduce((s, o) => s + _opDuracao(o), 0);
      const pend = g.itens.filter(o => _opStatus(o) !== 'feita').length;
      const comHora = g.itens.filter(o => _opInicioMin(o) != null);
      const jIni = comHora.length ? Math.min(...comHora.map(_opInicioMin)) : null;
      const jFim = comHora.length ? Math.max(...comHora.map(_opFimMin)) : null;
      // Vazios DESTA função dentro da jornada do dia — inclusive antes de ela
      // começar e depois de ela terminar, que é quando o posto fica parado
      // enquanto o resto da fábrica anda.
      const vazios = _opVazios(g.itens, jornadaDia, _OP_VAZIO_MIN);
      const paradoMin = vazios.reduce((s, v) => s + v.min, 0);
      // Quem trabalha neste posto e o que essas mesmas pessoas fazem nos outros.
      const pessoasDoPosto = new Set(g.itens.map(o => _normNome(_opResponsavelNome(o))).filter(Boolean));
      const sombras = pessoasDoPosto.size
        ? doDia.filter(o => _normNome(_opFuncaoNome(o)) !== _normNome(g.nome)
            && pessoasDoPosto.has(_normNome(_opResponsavelNome(o))))
        : [];
      // As linhas da função, com o vazio ENTRE duas operações virando linha
      // própria: é onde o buraco aparece para quem lê a sequência.
      const linhas = [];
      let fimAnterior = null, pessoaAnterior = '';
      g.itens.forEach((op, i) => {
        const ini = _opInicioMin(op);
        if (fimAnterior != null && ini != null && ini - fimAnterior >= _OP_VAZIO_MIN) {
          // A pessoa do vazio é a que acabou de parar; sem ela, a que retoma.
          const pessoa = pessoaAnterior || _opResponsavelNome(op);
          linhas.push(vazioLinha({ ini: fimAnterior, fim: ini, min: ini - fimAnterior, tipo: 'entre' }, pessoa, g.nome));
        }
        linhas.push(linhaHtml(op, i, g.itens.length));
        if (ini != null && _opDuracao(op)) {
          fimAnterior = Math.max(fimAnterior == null ? 0 : fimAnterior, _opFimMin(op));
          pessoaAnterior = _opResponsavelNome(op) || pessoaAnterior;
        }
      });
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
              ${paradoMin > 0 ? ` · <span class="exp-badge baixo" title="${esc(vazios.map(_opVazioTexto).join(' · '))}">${esc(_opDurTexto(paradoMin))} parada</span>` : ''}
              ${pend ? ` · <span class="exp-badge baixo">${pend} a fazer</span>` : ' · <span class="exp-badge ok">tudo feito</span>'}
            </div>
          </div>
          ${jan ? `<div class="op-faixa"><div class="op-faixa-eixo">${
            vazios.map(v => vazioBarra(v, jan)).join('')}${
            sombras.map(op => barraSombra(op, jan, cores.get(op.id))).join('')}${
            g.itens.map(op => barraHtml(op, jan, false, cores.get(op.id))).join('')}</div></div>` : ''}
          ${sombras.length ? `<div class="op-sombra-leg">Barras tracejadas — ${esc(Array.from(new Set(sombras.map(o => _opResponsavelNome(o)))).join(', '))} em outra função no mesmo horário, na cor do quadro de origem: ${
            Array.from(new Set(sombras.map(o => _opFuncaoNome(o)))).map(nome => {
              const cor = cores.get((sombras.find(o => _opFuncaoNome(o) === nome) || {}).id);
              return `<span class="amostra" style="background:${cor || 'var(--line)'};"></span>${esc(nome)}`;
            }).join(' · ')}.</div>` : ''}
          ${linhas.join('')}
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
    // Onde a continuidade do dia se rompe: janelas em que NENHUMA função opera —
    // o dia inteiro parado, não só um posto.
    const paradaGeral = _opVazios(doDia, jornadaDia, _OP_VAZIO_MIN).filter(v => v.tipo === 'entre');
    // O que falta para cada lote fechar a corrente inteira.
    const lotes = _opLotesIncompletos(doDia, data);
    const incompletos = lotes.filter(l => l.faltam.length || l.naoCabem.length);
    // Cada passo que falta é um botão: abre o modal já no horário do buraco, no
    // posto que costuma fazer aquele passo e com a OS na referência.
    const l0 = s => String(s).replace(/'/g, '&#39;');
    const botaoFalta = (l, p) => {
      const s = _opSugestaoPasso(data, l.lote, p, l.fase);
      const func = (STATE.funcoes || []).find(f => f.id === s.funcaoId);
      const quem = func ? func.nome : 'posto a definir';
      const naFase = (p.porFase && l.nFases > 1) ? ` na fase ${l.fase}${l.faseNome ? ' (' + l.faseNome + ')' : ''}` : '';
      // Rótulo com o nome que o planejamento usa para esse passo, não com o
      // rótulo interno da corrente.
      return `<button type="button" class="op-falta-btn admin-only"
        title="Incluir «${esc(s.nome)}»${esc(naFase)} às ${esc(s.inicio)} em ${esc(quem)}${s.duracaoMin ? ' · ' + esc(_opDurTexto(s.duracaoMin)) : ''} (o modal abre preenchido)"
        onclick="incluirOperacaoFaltante('${esc(data)}','${esc(l0(l.lote))}','${esc(p.cadeia)}',${p.ordem},${l.fase})">
        + ${esc(s.nome)} <span class="q">${esc(s.inicio)} · ${esc(quem)}</span></button>`;
    };
    // DIAGNÓSTICO sob demanda. Três perguntas, nesta ordem: falta operação?
    // alguma função está com duas coisas ao mesmo tempo? a sequência cruzada
    // entre as funções fecha? No fim, um resumo do que os outros avisos já
    // apontam, para o relatório não dar a impressão de que o resto está limpo.
    const analiseHtml = (() => {
      if (opAnaliseDia !== data) return '';
      // Só entra na seção o lote que tem passo cabendo HOJE: cobrar o que não
      // cabe mais viraria um pedido impossível na cara de quem planeja.
      const faltando = lotes.filter(l => l.faltam.length);
      const cruzadas = _opSobreposicoesMesmaFuncao(doDia);
      const ordemMsgs = [];
      doDia.forEach(o => {
        const e = foraDeOrdem.get(o.id);
        if (e) e.msgs.forEach(m => { if (!ordemMsgs.includes(m)) ordemMsgs.push(m); });
      });
      // Sobreposição de PESSOA = o que sobra dos conflitos depois de tirar as da
      // mesma função, que já têm seção própria acima.
      const naFuncao = new Set(cruzadas.flatMap(p => [p.a.id, p.b.id]));
      const pessoaN = doDia.filter(o => conflitos.has(o.id) && !naFuncao.has(o.id)).length;
      const foraJ = doDia.filter(o => _opForaDaJornada(o));
      // Duração zero: a operação existe no papel e não ocupa tempo nenhum. É o
      // que rompe a continuidade sem acusar conflito nenhum — a corrente fica
      // "certa" e mesmo assim o dia não flui.
      const semDuracao = doDia.filter(o => !_opDuracao(o));
      const secao = (titulo, n, corpo) => `
        <div class="op-analise-sec">
          <div class="op-analise-tit ${n ? 'alerta' : 'ok'}">${esc(titulo)} <span class="q">${n ? n : 'nada a apontar'}</span></div>
          ${n ? corpo : ''}
        </div>`;
      const linhaOp = op => `${esc(_opFuncaoNome(op))} · ${esc(op.operacao) || '—'} <span class="q">${esc(_opJanelaTexto(op))}</span>`;
      return `
      <div class="op-analise" id="op-analise-${esc(data)}">
        <div class="op-analise-cab">
          Análise do dia ${esc(formatDate(data))}
          <span style="display:flex;gap:6px;">
            ${(faltando.length || cruzadas.length || ordemMsgs.length || pessoaN > 0 || foraJ.length)
              ? `<button type="button" class="btn small primary admin-only" title="Aplica no dia tudo o que esta análise apontou e o programa sabe resolver sozinho: sobreposição na mesma função, ordem do lote, mesma pessoa em duas funções, jornada e pausas." onclick="corrigirOrdemOperacoes('${esc(data)}')">⇄ Corrigir o que dá</button>`
              : ''}
            <button type="button" class="btn small" onclick="analisarDiaOperacoes('${esc(data)}')">fechar</button>
          </span>
        </div>
        ${secao('Operações esperadas que ainda cabem hoje',
          faltando.length && `${faltando.length} ${faltando.some(l => l.nFases > 1) ? 'fase(s) de lote' : 'lote(s)'}`,
          faltando.map(l => `<div class="op-analise-item"><b>${esc(l.rotulo)}</b> — ${l.feitos}/${l.total} ·
            ${l.faltam.map(p => botaoFalta(l, p)).join(' ')}${
            l.naoCabem.length ? `<div class="sub">Não cabe mais hoje — continua em ${esc(formatDate(_opProximoDiaUtil(data)))}: ${esc(l.naoCabem.map(p => p.nome).join(' · '))}</div>` : ''}</div>`).join(''))}
        ${secao('Sobreposição dentro da mesma função', cruzadas.length && `${cruzadas.length} par(es)`,
          cruzadas.map(p => `<div class="op-analise-item">
            <b>${esc(p.funcao)}</b> — ${esc(_opDurTexto(p.min))} sobrepostos (${esc(_opHHMM(p.ini))} → ${esc(_opHHMM(p.fim))})
            <div class="sub">${linhaOp(p.a)}</div>
            <div class="sub">${linhaOp(p.b)}</div>
            <div class="admin-only" style="margin-top:2px;display:flex;gap:4px;">
              <button class="btn small" onclick="abrirModalOperacao('${esc(p.a.id)}')">ajustar a 1ª</button>
              <button class="btn small" onclick="abrirModalOperacao('${esc(p.b.id)}')">ajustar a 2ª</button>
            </div>
          </div>`).join('')
          + '<div class="op-analise-item sub">A segunda operação de cada par passa a começar quando a primeira termina — é o que o botão abaixo faz.</div>')}
        ${secao('Coerência da sequência cruzada entre funções', ordemMsgs.length && `${ordemMsgs.length} ponto(s)`,
          ordemMsgs.map(m => `<div class="op-analise-item">${esc(m)}</div>`).join(''))}
        ${secao('Operações sem duração', semDuracao.length && `${semDuracao.length} operação(ões)`,
          `<div class="op-analise-item sub">Operação com <b>zero minuto</b> não ocupa tempo nenhum: ela não aparece na linha do tempo e o passo seguinte começa no mesmo instante. Quando é o <b>enfesto</b>, todas as fases ficam prontas de uma vez e o posto seguinte recebe o lote inteiro de enxurrada — é o que abre as esperas longas entre o corte e a separação. O tempo se cadastra em <b>Funções</b>, ou sai sozinho do histórico quando os horários de enfesto forem lançados na folha da OS.</div>`
          + semDuracao.map(o => `<div class="op-analise-item">${linhaOp(o)}
            <div class="admin-only" style="margin-top:2px;"><button class="btn small" onclick="abrirModalOperacao('${esc(o.id)}')">informar a duração</button></div>
          </div>`).join(''))}
        <div class="op-analise-sec">
          <div class="op-analise-tit ${(pessoaN > 0 || foraJ.length || _opPausasDessincronizadas(doDia)) ? 'alerta' : 'ok'}">Outros pontos <span class="q">${
            [pessoaN > 0 ? `${pessoaN} com a mesma pessoa em duas funções` : '',
             foraJ.length ? `${foraJ.length} fora da jornada` : '',
             _opPausasDessincronizadas(doDia) ? 'pausas dessincronizadas' : ''
            ].filter(Boolean).join(' · ') || 'nada a apontar'}</span></div>
          ${(pessoaN > 0 || foraJ.length || _opPausasDessincronizadas(doDia))
            ? '<div class="op-analise-item sub">O botão <b>Corrigir o que dá</b>, no alto desta análise, resolve estes três: separa a pessoa, traz o que está fora da jornada e alinha as pausas.</div>'
            : ''}
        </div>
      </div>`;
    })();
    const lotesHtml = lotes.length ? `
      <div class="op-lotes">
        <div class="op-lotes-cab">${lotes.some(l => l.nFases > 1)
          ? 'Lotes do dia, por fase do enfesto'
          : 'Lotes do dia'} · ${lotes.length - incompletos.length} de ${lotes.length} com a sequência completa</div>
        ${lotes.map(l => `
          <div class="op-lote-linha ${l.faltam.length ? 'falta' : 'ok'}">
            <span class="n">${esc(l.rotulo)}</span>
            <span class="c">${l.feitos}/${l.total}</span>
            <span class="f">${l.faltam.length
              ? l.faltam.map(p => botaoFalta(l, p)).join(' ')
              : (l.naoCabem.length ? 'nada mais cabe hoje' : 'sequência completa')}${
              // Continuação: o que não cabe mais no dia e o que já está marcado
              // noutra data. Sem botão — não é para incluir hoje, é para saber
              // que o lote não parou por esquecimento.
              l.naoCabem.length ? `<span class="cont">continua em ${esc(formatDate(_opProximoDiaUtil(data)))}: ${esc(l.naoCabem.map(p => p.nome).join(' · '))}</span>` : ''}${
              l.noutroDia.length ? `<span class="cont">já planejado: ${esc(l.noutroDia.map(x => `${x.passo.nome} (${formatDate(x.data)})`).join(' · '))}</span>` : ''}</span>
          </div>`).join('')}
      </div>` : '';
    return `
      <div class="card exp-ocor">
        <div class="exp-ocor-head">
          <div>
            <div class="exp-ocor-data">${_EXP_DIAS_CURTO[_expData(data).getDay()]} · ${esc(formatDate(data))}${data === _expHoje() ? ' <span class="exp-badge info">hoje</span>' : ''}</div>
            <div class="exp-ocor-nome">
              ${abre != null ? `Jornada <b>${esc(_opHHMM(abre))} → ${esc(_opHHMM(fecha))}</b> · ` : ''}${grupos.length} ${grupos.length === 1 ? 'posto' : 'postos'} em paralelo · ${esc(_opDurTexto(totMin))} de operação somados${prioridades}
            </div>
            ${paradaGeral.length ? `<div class="exp-ocor-nome" style="color:var(--accent-dark);">
              Sem <b>nenhuma</b> função operando: ${paradaGeral.map(v => `<b>${esc(_opHHMM(v.ini))} → ${esc(_opHHMM(v.fim))}</b> (${esc(_opDurTexto(v.min))})`).join(' · ')}
            </div>` : ''}
          </div>
          <div class="admin-only" style="display:flex;gap:6px;flex-wrap:wrap;">
            ${(doDia.some(o => foraDeOrdem.has(o.id) || conflitos.has(o.id) || _opForaDaJornada(o)) || _opPausasDessincronizadas(doDia))
              ? `<button class="btn" title="Sincroniza as pausas, garante a duração medida do enfesto e empurra para frente o que quebra a ordem do lote ou põe a mesma pessoa em duas funções ao mesmo tempo." onclick="corrigirOrdemOperacoes('${esc(data)}')">⇄ Organizar o dia</button>`
              : ''}
            <button class="btn" title="Confere o dia: operação esperada que falta, sobreposição dentro da mesma função e a coerência da sequência cruzada entre as funções." onclick="analisarDiaOperacoes('${esc(data)}')">${opAnaliseDia === data ? '✕ Fechar análise' : '🔍 Analisar o dia'}</button>
            <button class="btn primary" title="O planejamento do dia começa pela OS: escolhida a OS, o programa monta a sequência inteira de operações, cada uma na função que a executa." onclick="abrirModalAlocarOS('${esc(data)}')">+ Alocar OS</button>
            ${lotes.length ? `<button class="btn" title="Tira do planejamento, de uma vez, todas as operações de uma OS alocada — em um dia ou no plano inteiro. É o desfazer do «Alocar OS»." onclick="abrirModalDesalocarOS('${esc(data)}')">− Retirar OS</button>` : ''}
            <button class="btn" onclick="abrirModalOperacao('','${esc(data)}')">+ Operação avulsa</button>
          </div>
        </div>
        ${jan ? reguaHtml(jan) : ''}
        ${blocos}
        ${lotesHtml}
        ${analiseHtml}
      </div>`;
  };

  // VISTA POR PESSOA: um quadro único do dia com uma faixa por pessoa, no mesmo
  // eixo de horas. Vê-se de relance tudo que cada pessoa faz e onde ela se
  // sobrepõe (mesma pessoa em duas tarefas ao mesmo tempo = conflito vermelho).
  const diaHtmlPessoa = (data, doDia) => {
    const jan = _opJanelaDoDia(doDia);
    const cores = _opMapaCores(doDia);
    const porPessoa = new Map();
    doDia.forEach(op => {
      const nome = _opResponsavelNome(op) || '— a definir —';
      if (!porPessoa.has(nome)) porPessoa.set(nome, []);
      porPessoa.get(nome).push(op);
    });
    const lanes = Array.from(porPessoa.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([nome, itens]) => {
        const temConf = itens.some(o => conflitos.has(o.id));
        const minutos = itens.reduce((s, o) => s + _opDuracao(o), 0);
        const semHora = itens.filter(o => _opInicioMin(o) == null).length;
        return `
          <div class="op-pessoa-lane ${temConf ? 'conflito' : ''}">
            <div class="op-pessoa-nome">${esc(nome)}${temConf ? ' <span class="exp-badge alto">conflito</span>' : ''}
              <div class="op-pessoa-tot">${esc(_opDurTexto(minutos))} · ${itens.length} tarefa${itens.length > 1 ? 's' : ''}${semHora ? ` · ${semHora} sem horário` : ''}</div>
            </div>
            ${jan ? `<div class="op-faixa"><div class="op-faixa-eixo">${itens.map(op => barraHtml(op, jan, true, cores.get(op.id))).join('')}</div></div>`
                  : '<div class="op-pessoa-tot" style="padding:2px 0;">Sem horário definido nas tarefas.</div>'}
          </div>`;
      }).join('');
    const totMin = doDia.reduce((s, o) => s + _opDuracao(o), 0);
    const nConf = doDia.filter(o => conflitos.has(o.id)).length;
    return `
      <div class="card exp-ocor">
        <div class="exp-ocor-head">
          <div>
            <div class="exp-ocor-data">${_EXP_DIAS_CURTO[_expData(data).getDay()]} · ${esc(formatDate(data))}${data === _expHoje() ? ' <span class="exp-badge info">hoje</span>' : ''}</div>
            <div class="exp-ocor-nome">${porPessoa.size} pessoa${porPessoa.size === 1 ? '' : 's'} · ${esc(_opDurTexto(totMin))} de operação${nConf ? ` · <span class="exp-badge alto">${nConf} em conflito</span>` : ''}</div>
          </div>
          <div class="admin-only"><button class="btn" onclick="abrirModalOperacao('','${esc(data)}')">+ Operação neste dia</button></div>
        </div>
        ${jan ? reguaHtml(jan) : ''}
        <div class="op-pessoa-quadro">${lanes}</div>
      </div>`;
  };

  const porDia = new Map();
  ops.forEach(op => {
    if (!porDia.has(op.data)) porDia.set(op.data, []);
    porDia.get(op.data).push(op);
  });
  const montarDia = opVista === 'pessoa' ? diaHtmlPessoa : diaHtml;
  const cards = Array.from(porDia.keys()).sort().map(d => montarDia(d, porDia.get(d))).join('');

  const semFuncoes = !(STATE.funcoes || []).length;
  const vazio = `
    <div class="card">
      <div class="empty" style="padding:24px 0;text-align:center;">
        ${semFuncoes
          ? 'Nenhuma <b>função</b> cadastrada ainda. O planejamento é feito por posto de trabalho — cadastre as funções em <a href="#" onclick="goto(\'cad-funcoes\'); return false;">Funções</a> antes de começar.'
          : 'Nenhuma operação planejada neste período. Comece por <b>+ Alocar OS no dia</b>: escolhida a OS, o programa monta a sequência inteira de operações, cada uma na função que a executa.'}
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

// `pre` preenche a operação NOVA já com o que se sabe dela (usado pelo botão do
// passo que falta): { operacao, inicio, duracaoMin, referencia }.
function abrirModalOperacao(opId = '', dataPre = '', funcaoIdPre = '', pre = null) {
  if (!exigirAdmin('planejar operações')) return;
  if (!(STATE.funcoes || []).length) {
    return toast('Cadastre ao menos uma função antes de planejar operações', 'err');
  }
  const op = opId ? (STATE.operacoes || []).find(x => x.id === opId) : null;
  // A fase do enfesto viaja com o modal: quem edita mexe em horário e posto, não
  // na fase — e perder o vínculo tiraria a operação da corrente da fase dela.
  _opModalCtx = {
    editId: op ? opId : '',
    faseOrdem: op ? (Number(op.faseOrdem) || 0) : (Number(pre && pre.faseOrdem) || 0),
    faseNome: op ? (op.faseNome || '') : ((pre && pre.faseNome) || '')
  };

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
      <div class="field"><label>Data *</label><input type="date" id="op-data" value="${esc(data)}" onchange="_opSugerirInicio();_opAtualizarJanela()"></div>
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
          <option value="etapa" ${escopo === 'etapa' ? 'selected' : ''}>Etapas específicas</option>
        </select>
        <div class="field-hint">O padrão engloba todas as etapas da função. Escolha <b>etapas específicas</b> para planejar o posto por partes.</div>
      </div>
      <div class="field full hidden" id="op-wrap-etapa">
        <label>Etapas *</label>
        <div id="op-etapas-lista" style="border:1px solid var(--line);border-radius:6px;padding:6px 10px;max-height:180px;overflow:auto;"></div>
        <div class="field-hint">Marque uma ou mais operações — aparecem as cadastradas na função escolhida.</div>
      </div>
      <div class="field full">
        <label>Operação *</label>
        <input type="text" id="op-operacao" list="op-sugestoes" value="${esc(op ? (op.operacao || '') : '')}" placeholder="Ex.: Enfesto e corte do dia" autocomplete="off" oninput="_opTempoMedioDaReferencia()">
        <datalist id="op-sugestoes"></datalist>
        <div class="field-hint" id="op-sug-hint">Descreva a operação inteira do posto — as etapas internas ficam subentendidas.</div>
      </div>
      <div class="field">
        <label>Início *</label>
        <input type="time" id="op-inicio" value="${esc(op ? (op.inicio || '') : _opHHMM(_OP_JORNADA.ini))}" min="${esc(_opHHMM(_OP_JORNADA.ini))}" max="${esc(_opHHMM(_OP_JORNADA.fim))}" oninput="_opAtualizarJanela()">
        <label class="op-fixo-chk" style="display:flex;gap:6px;align-items:center;font-weight:400;font-size:12px;margin:6px 0 0;">
          <input type="checkbox" id="op-inicio-fixo" ${op && op.inicioFixo ? 'checked' : ''} onchange="_opAtualizarJanela()" style="width:auto;margin:0;">
          Horário fixo (não reencaixar após a operação anterior)
        </label>
        <div class="field-hint">Se a hora digitada <b>não for</b> o encaixe logo após a operação anterior, ela é mantida como está e as seguintes contam a partir dela — pode ficar <b>intervalo vazio</b> antes. Digitando exatamente o encaixe, a operação segue <b>automática</b> e anda junto quando a anterior mudar. Marque a caixa para travar a hora mesmo quando ela coincide com o encaixe; <b>desmarcar solta</b> a operação de volta para a fila do posto, e o horário passa a ser o que o encadeamento der.</div>
      </div>
      <div class="field">
        <label>Duração total *</label>
        <div style="display:flex;gap:6px;align-items:center;">
          <input type="number" min="0" step="1" id="op-dur-h" value="${esc(durH)}" placeholder="0" oninput="_opDuracaoManual()" style="width:70px;">
          <span style="font-size:12px;color:var(--ink-3);">h</span>
          <input type="number" min="0" max="59" step="5" id="op-dur-m" value="${esc(durM)}" placeholder="0" oninput="_opDuracaoManual()" style="width:70px;">
          <span style="font-size:12px;color:var(--ink-3);">min</span>
        </div>
        <div class="field-hint">Tempo total até concluir, com todas as etapas do posto incluídas.</div>
        <div class="field-hint" id="op-tempo-medio"></div>
      </div>
      <div class="field">
        <label>Responsável</label>
        <select id="op-responsavel"></select>
        <div class="field-hint">Pessoas da <b>Equipe</b> com esta função aparecem primeiro.</div>
      </div>
      <div class="field">
        <label>Referência (opcional)</label>
        <input type="text" id="op-referencia" list="op-refs" value="${esc(op ? (op.referencia || '') : '')}" placeholder="Ex.: lote inverno, OS 1042/1051" autocomplete="off" oninput="_opTempoMedioDaReferencia()">
        <datalist id="op-refs">${refs.map(r => `<option value="${esc(r)}"></option>`).join('')}</datalist>
        <div class="field-hint">Texto livre — lote, coleção, OSs do dia. Escolhendo uma OS da lista, a <b>duração</b> vem do tempo já medido na grade dela. O <b>F3/5</b> depois do modelo é a <b>fase do enfesto</b>: é ela que diz em qual das correntes da OS esta operação entra — apagar o marcador joga a operação para a fase 1.</div>
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

  const etapasIniciais = op ? (Array.isArray(op.etapas) ? op.etapas : (op.etapa ? [op.etapa] : [])) : [];
  _opTrocouFuncao(op ? (op.responsavelId || '') : '', etapasIniciais);
  const campoOp = document.getElementById('op-operacao');
  if (campoOp) campoOp.dataset.etapasAnterior = etapasIniciais.join(' + ');
  // Pré-preenchimento DEPOIS de _opTrocouFuncao: ela sugere início e duração
  // sozinha, e aqui o que vale é o buraco que o botão veio tapar.
  if (!op && pre) {
    const por = (id, v) => { const el = document.getElementById(id); if (el && v != null && v !== '') el.value = v; };
    por('op-operacao', pre.operacao);
    por('op-inicio', pre.inicio);
    por('op-referencia', pre.referencia);
    if (pre.duracaoMin > 0) {
      por('op-dur-h', Math.floor(pre.duracaoMin / 60) || '');
      por('op-dur-m', pre.duracaoMin % 60 || (Math.floor(pre.duracaoMin / 60) ? 0 : ''));
    }
    // A referência recém-posta pode ter tempo medido na grade: deixa a regra do
    // tempo médio opinar sobre a duração (ela só escreve se o campo estiver vazio
    // ou se o valor for dela).
    _opTempoMedioDaReferencia();
  }
  _opAtualizarJanela();
  openModal('modal-op');
}

// Função escolhida muda três coisas: as sugestões de operação, as etapas
// oferecidas quando o plano é de uma etapa só, e a lista de responsáveis.
// Mantém a pessoa e a etapa já selecionadas quando continuam válidas.
function _opTrocouFuncao(responsavelPre = null, etapasPre = null) {
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

  // Etapas ofertadas quando a abrangência é "etapas específicas": checklist das
  // etapas cadastradas NA FUNÇÃO (só elas). Preserva o que já estava marcado.
  const lista = document.getElementById('op-etapas-lista');
  if (lista) {
    const marcar = etapasPre != null ? etapasPre : _opEtapasMarcadas();
    const etapasFunc = _opEtapasDaFuncao(funcaoId);
    lista.innerHTML = etapasFunc.length
      ? etapasFunc.map(nome => `<label class="etapa-check" style="display:flex;align-items:center;gap:8px;font-size:13px;padding:4px 2px;border-bottom:1px dotted var(--line);">`
          + `<input type="checkbox" class="op-etapa-chk" value="${esc(nome)}" ${marcar.includes(nome) ? 'checked' : ''} onchange="_opEscolheuEtapa()"> ${esc(nome)}</label>`).join('')
      : (funcaoId
          ? '<span style="color:var(--ink-3);font-size:12px;">Esta função não tem operações cadastradas. Cadastre as operações dela em <b>Funções</b>.</span>'
          : '<span style="color:var(--ink-3);font-size:12px;">Escolha a função para ver as etapas.</span>');
  }

  const { dentro, fora } = _opPessoasDaFuncao(funcao ? funcao.nome : '');
  const opt = p => `<option value="${esc(p.id)}" ${p.id === manter ? 'selected' : ''}>${esc(p.nome)}</option>`;
  selR.innerHTML = '<option value="">— a definir —</option>'
    + (dentro.length ? `<optgroup label="${esc(funcao ? funcao.nome : 'Da função')}">${dentro.map(opt).join('')}</optgroup>` : '')
    + (fora.length ? `<optgroup label="Outras pessoas">${fora.map(opt).join('')}</optgroup>` : '');

  _opTrocouEscopo();
  _opSugerirInicio();            // posto novo → hora em que aquele posto fica livre
  _opTempoMedioDaReferencia();   // OS já citada → duração medida na grade dela
  _opAtualizarJanela();
}

// Mostra/esconde o seletor de etapa conforme a abrangência escolhida.
function _opTrocouEscopo() {
  const escopo = document.getElementById('op-escopo')?.value || 'completa';
  const wrap = document.getElementById('op-wrap-etapa');
  const hint = document.getElementById('op-sug-hint');
  if (wrap) wrap.classList.toggle('hidden', escopo !== 'etapa');
  if (hint && escopo === 'etapa') {
    hint.innerHTML = 'Planejando <b>etapas específicas</b>: marque uma ou mais etapas. Os nomes vêm para o campo Operação e podem ser detalhados.';
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

// Marcar/desmarcar etapas preenche o nome da operação com os nomes juntos. Só
// sobrescreve campo vazio ou que ainda tem os nomes anteriores — texto digitado
// à mão fica de pé.
function _opEscolheuEtapa() {
  const campo = document.getElementById('op-operacao');
  if (!campo) return;
  const juntos = _opEtapasMarcadas().join(' + ');
  const anterior = campo.dataset.etapasAnterior || '';
  if (!campo.value.trim() || campo.value.trim() === anterior) campo.value = juntos;
  campo.dataset.etapasAnterior = juntos;
  // A etapa marcada é o que casa a operação com uma FASE medida da grade.
  _opTempoMedioDaReferencia();
}

/* ---- duração vinda do tempo já medido na grade da OS ---- */

// Id da grade cadastrada de uma OS, pela mesma chave que agrupa o histórico.
function _gradeIdDaOS(o) {
  const k = _osGradeKey(o);
  return k.indexOf('g:') === 0 ? k.slice(2) : '';
}

// A OS citada no campo Referência. O campo é texto livre ("0435 · Blusa Moletom
// Tricolor", "OS 0435", "lote inverno"), então casa pelo NÚMERO da OS que
// aparecer no texto — é o que a lista de sugestões escreve.
function _opOsDaReferencia(txt) {
  const t = String(txt || '').trim();
  if (!t) return null;
  const candidatas = (STATE.ordens || []).filter(o => (o.os || '').toString().trim());
  // Casa o número mais LONGO primeiro: "0435" não pode ganhar de "10435".
  return candidatas
    .slice()
    .sort((a, b) => String(b.os).length - String(a.os).length)
    .find(o => new RegExp('(^|\\D)' + String(o.os).trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\D|$)').test(t)) || null;
}

// Tempo já MEDIDO que serve para esta operação, a partir da grade da OS citada:
//   • se a operação (nome ou etapas marcadas) casa com uma FASE medida daquela
//     grade, vale a média daquela fase;
//   • senão, quando a operação é de enfesto, vale a soma das fases medidas — o
//     enfesto inteiro daquele lote;
//   • senão devolve só o informativo, sem número para preencher: aplicar tempo
//     de enfesto numa operação de outra natureza seria inventar dado.
function _opTempoMedidoParaOS(os, nomeOperacao, etapas, funcaoNome) {
  const vazio = { min: 0, n: 0, rotulo: '', aplicavel: false, temDados: false };
  const gradeId = os ? _gradeIdDaOS(os) : '';
  if (!gradeId) return vazio;
  const linhas = temposFasesDaGrade(gradeId).filter(l => l.n > 0);
  if (!linhas.length) return vazio;
  const alvos = [];
  (etapas || []).forEach(e => alvos.push(_normFaseNome(e)));
  if (nomeOperacao) alvos.push(_normFaseNome(nomeOperacao));
  const casada = linhas.find(l => alvos.includes(_normFaseNome(l.nome)));
  if (casada) {
    return { min: casada.mediaMin, n: casada.n, rotulo: `fase "${casada.nome}"`, osNums: casada.osNums || [], aplicavel: true, temDados: true };
  }
  const total = linhas.reduce((s, l) => s + l.mediaMin, 0);
  // Quem diz se a operação é de enfesto é a FUNÇÃO, não o nome da operação:
  // "Corte de enfesto" é operação de CORTE e leva minutos, não as horas do
  // enfesto inteiro — pelo nome ela cairia na regra e receberia o tempo errado.
  const ehEnfesto = /enfest/i.test(String(funcaoNome || ''));
  return {
    min: total, n: Math.max(...linhas.map(l => l.n)),
    rotulo: `enfesto inteiro (${linhas.length} fase${linhas.length === 1 ? '' : 's'} medida${linhas.length === 1 ? '' : 's'})`,
    osNums: Array.from(new Set(linhas.flatMap(l => l.osNums || []))),
    aplicavel: ehEnfesto, temDados: true
  };
}

// Marca a duração como DIGITADA: a partir daí, escolher outra OS não sobrescreve
// o que o usuário pôs à mão.
function _opDuracaoManual() {
  ['op-dur-h', 'op-dur-m'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.dataset.auto = '';
  });
  _opAtualizarJanela();
}

// Ao escolher a OS no campo Referência, traz o tempo já medido na grade dela para
// a duração da operação. Só preenche quando o campo está vazio ou quando o valor
// que está lá foi posto por esta mesma automação — o que o usuário digitou manda.
function _opTempoMedioDaReferencia() {
  const box = document.getElementById('op-tempo-medio');
  const h = document.getElementById('op-dur-h'), m = document.getElementById('op-dur-m');
  if (!box || !h || !m) return;
  // TEMPO CADASTRADO na função para esta operação: é o que a casa definiu, e vale
  // para qualquer operação — não só as de enfesto. Entra primeiro; o tempo medido
  // na grade (abaixo) fala mais alto quando existe, por ser daquele lote.
  const cad = _tempoOperacaoCadastrada(
    document.getElementById('op-funcao')?.value,
    document.getElementById('op-operacao')?.value);
  const podeEscreverCad = _opDuracaoDoForm() === 0 || h.dataset.auto === '1' || m.dataset.auto === '1';
  if (cad > 0 && podeEscreverCad) {
    h.value = Math.floor(cad / 60) || '';
    m.value = cad % 60 || (Math.floor(cad / 60) ? 0 : '');
    h.dataset.auto = '1'; m.dataset.auto = '1';
  }
  const os = _opOsDaReferencia(document.getElementById('op-referencia')?.value);
  if (!os) {
    box.innerHTML = cad > 0
      ? `Tempo cadastrado nesta função para <b>${esc(document.getElementById('op-operacao')?.value || '')}</b>: <b>${esc(_opDurTexto(cad))}</b>.`
      : '';
    return;
  }
  const funcao = (STATE.funcoes || []).find(f => f.id === (document.getElementById('op-funcao')?.value || ''));
  const r = _opTempoMedidoParaOS(os, document.getElementById('op-operacao')?.value, _opEtapasMarcadas(), funcao && funcao.nome);
  if (!r.temDados) {
    box.innerHTML = `OS <b>${esc(os.os || '—')}</b> · ${esc(os.modeloNome || 'sem modelo')}: a grade dela ainda não tem tempo medido.`;
    return;
  }
  const podeEscrever = _opDuracaoDoForm() === 0 || h.dataset.auto === '1' || m.dataset.auto === '1';
  if (r.aplicavel && podeEscrever) {
    h.value = Math.floor(r.min / 60) || '';
    m.value = r.min % 60 || (Math.floor(r.min / 60) ? 0 : '');
    h.dataset.auto = '1'; m.dataset.auto = '1';
  }
  // De ONDE veio o número: com uma medição só, a média é aquela OS — e é o
  // usuário quem sabe se ela vale para este lote. Sem dizer a origem e as
  // camadas, um enfesto de 3h30 medido num lote de 50 camadas volta como regra
  // sem ninguém poder discordar.
  const camadasDe = num => {
    const o2 = (STATE.ordens || []).find(x => String(x.os || '').trim() === String(num));
    const c = o2 && o2.enfesto && o2.enfesto.camadas;
    return c ? `${num} (${c} camadas)` : String(num);
  };
  const origem = (r.osNums && r.osNums.length)
    ? ` — medido em ${r.n === 1 ? 'uma OS' : r.n + ' OS'}: ${r.osNums.slice(0, 3).map(camadasDe).join(', ')}`
    : '';
  const camadasAqui = os.enfesto && os.enfesto.camadas;
  box.innerHTML = `OS <b>${esc(os.os || '—')}</b>${camadasAqui ? ` (${esc(camadasAqui)} camadas)` : ''} · grade <b>${esc(_gradeNomeDaOS(os))}</b>: `
    + `<b>${esc(_opDurTexto(r.min))}</b>${esc(origem)}`
    + (r.aplicavel
      ? (podeEscrever ? ' — <b>preenchido na duração</b>.' : ' — a duração digitada foi mantida.')
      : ' — informativo: esta operação não é de enfesto, então a duração não foi preenchida.');
  _opAtualizarJanela();
}

// Nome da grade cadastrada de uma OS (para o texto do modal).
function _gradeNomeDaOS(o) {
  const id = _gradeIdDaOS(o);
  const g = id ? (STATE.grades || []).find(x => x.id === id) : null;
  return (g && (g.nome || g.descricao)) || (o.grade && o.grade.descricao) || '—';
}

// Para operação NOVA, sugere a hora em que o posto fica livre (fim da última
// operação daquele posto no dia) — é o mesmo encaixe que a sincronização faria ao
// salvar, então o modal mostra de véspera o horário que vai valer. Não toca no
// campo se o usuário já fixou a hora. O posto da esteira fica de fora: as
// operações dele rodam em paralelo, não em fila.
function _opSugerirInicio() {
  if (!_opModalCtx || _opModalCtx.editId) return;
  const campo = document.getElementById('op-inicio');
  const fixo = document.getElementById('op-inicio-fixo');
  if (!campo || (fixo && fixo.checked)) return;
  const data = document.getElementById('op-data')?.value || '';
  const funcaoId = document.getElementById('op-funcao')?.value || '';
  if (!data || !funcaoId) return;
  const funcao = (STATE.funcoes || []).find(f => f.id === funcaoId);
  if (funcao && ehFuncaoOperadorEsteira(funcao.nome)) return;
  const fins = (STATE.operacoes || [])
    .filter(o => o.data === data && (o.funcaoId || '') === funcaoId)
    .map(_opFimMin).filter(m => m != null);
  // A sugestão nunca sai da jornada: posto que já encheu o dia não vira sugestão
  // de operação começando depois do expediente.
  if (fins.length) campo.value = _opHHMM(Math.min(Math.max(...fins), _OP_JORNADA.fim));
}

// Mostra ao vivo o término calculado — é o número que o planejador confere.
function _opAtualizarJanela() {
  const info = document.getElementById('op-info');
  if (!info) return;
  const ini = _opMin(document.getElementById('op-inicio')?.value);
  const dur = _opDuracaoDoForm();
  if (ini == null) { info.textContent = 'Informe a hora de início para ver o término.'; return; }
  if (!dur) { info.innerHTML = `Começa às <b>${esc(_opHHMM(ini))}</b>. Informe a duração total para calcular o término.`; return; }
  const fixo = !!document.getElementById('op-inicio-fixo')?.checked;
  const fora = _opForaDaJornada({ inicio: _opHHMM(ini), duracaoMin: dur });
  info.innerHTML = `Começa às <b>${esc(_opHHMM(ini))}</b>, leva <b>${esc(_opDurTexto(dur))}</b> e conclui às <b>${esc(_opHHMM(ini + dur))}</b>.`
    + (fora ? ` <span class="exp-badge alto" title="Jornada do setor: ${esc(_opJornadaTexto())}">fora da jornada</span>` : '')
    + (fixo
      ? ' <b>Horário fixo</b> — esta hora é mantida como está e as operações seguintes do posto contam a partir dela.'
      : ' Se esta hora for o encaixe logo após a operação anterior do posto, a operação segue <b>automática</b>; se for outra, ela é mantida e vira o novo ponto de partida do posto.');
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
  const etapas = escopo === 'etapa' ? _opEtapasMarcadas() : [];
  if (escopo === 'etapa' && !etapas.length) return toast('Marque ao menos uma etapa que será executada', 'err');
  // Etapas marcadas e nome livre em branco: os nomes das etapas já descrevem a
  // operação — não faz sentido exigir que o usuário redigite o mesmo texto.
  const operacao = v('op-operacao').trim() || etapas.join(' + ');
  if (!operacao) return toast('Descreva a operação', 'err');
  const inicio = v('op-inicio');
  if (_opMin(inicio) == null) return toast('Informe a hora de início', 'err');
  // Hora fixa: a sincronização do posto não reescreve esta operação (e pode
  // sobrar intervalo vazio antes dela).
  const inicioFixo = !!document.getElementById('op-inicio-fixo')?.checked;
  const duracaoMin = _opDuracaoDoForm();
  if (!duracaoMin) return toast('Informe a duração total da operação', 'err');
  // Marca de quem é a duração: DIGITADA (o campo perdeu a marca `auto` ao ser
  // editado à mão) ou vinda do tempo medido da grade. O ajuste do dia não eleva
  // duração digitada — é a diferença entre ajudar quem não decidiu e insistir
  // com quem decidiu.
  const duracaoManual = document.getElementById('op-dur-h')?.dataset.auto !== '1'
    && document.getElementById('op-dur-m')?.dataset.auto !== '1';
  // Jornada do setor: a operação tem que caber entre a abertura e o fim do
  // expediente. Sem esta trava, o encaixe automático e o ajuste de ordem
  // acabavam empurrando operação para depois das 17:30.
  const foraJornada = _opForaDaJornada({ inicio, duracaoMin });
  if (foraJornada) return toast(`${foraJornada}. A jornada vai das ${_opJornadaTexto()} — ajuste o início ou a duração.`, 'err');

  const responsavelId = v('op-responsavel');
  const pessoa = (STATE.equipe || []).find(p => p.id === responsavelId);
  const status = _OP_STATUS[v('op-status')] ? v('op-status') : 'pendente';
  const prioridade = _OP_PRIORIDADE[v('op-prioridade')] ? v('op-prioridade') : 'eletiva';

  // FASE DO ENFESTO. Manda o que a referência diz: o campo é texto livre e o
  // usuário pode escrever "…· F3/5 Corpo Parte 3" à mão — se o campo gravado
  // discordasse do texto, a agenda mostraria uma fase e a corrente usaria outra.
  // Sem marcador no texto, vale a fase que a operação já tinha.
  const referencia = v('op-referencia').trim();
  const mFase = referencia.match(/(?:^|[·\s])F(\d+)(?:\s*\/\s*\d+)?\b/i);
  const osRef = _opOsDaReferencia(referencia);
  const fasesRef = osRef ? _opFasesDaOS(osRef) : [];
  const nFase = mFase ? Number(mFase[1]) : (Number(_opModalCtx.faseOrdem) || 0);
  const faseValida = fasesRef.length > 1 && fasesRef.some(f => f.ordem === nFase);
  const faseObj = faseValida ? fasesRef.find(f => f.ordem === nFase) : null;

  const campos = {
    data,
    funcaoId, funcaoNome: funcao.nome,
    operacao, escopo, etapas,
    inicio, duracaoMin, inicioFixo, duracaoManual,
    responsavelId: pessoa ? pessoa.id : '',
    responsavelNome: pessoa ? pessoa.nome : '',
    referencia,
    faseOrdem: faseObj ? faseObj.ordem : 0,
    faseNome: faseObj ? faseObj.nome : '',
    status, prioridade,
    obs: v('op-obs').trim()
  };

  if (!Array.isArray(STATE.operacoes)) STATE.operacoes = [];
  let alvo = null;
  // Estado ANTES de gravar: é o que diferencia "quero esta hora" de "quero
  // soltar esta operação" quando a caixa é desmarcada.
  const eraFixo = !!(_opModalCtx.editId
    && (STATE.operacoes.find(x => x.id === _opModalCtx.editId) || {}).inicioFixo);
  if (_opModalCtx.editId) {
    const i = STATE.operacoes.findIndex(x => x.id === _opModalCtx.editId);
    if (i >= 0) { STATE.operacoes[i] = { ...STATE.operacoes[i], ...campos }; alvo = STATE.operacoes[i]; }
  } else {
    alvo = { id: uid(), ...campos };
    STATE.operacoes.push(alvo);
  }
  // A hora digitada é uma ESCOLHA ou é o encaixe automático? Quem responde é a
  // própria sincronização: grava sem travar nada e deixa ela opinar. Se ela
  // reescreve a hora, o que o usuário digitou não era o encaixe — então é
  // deliberado e a operação vira âncora (mantém a hora, com o intervalo vazio que
  // sobrar antes). Se a hora sobrevive, a operação segue AUTOMÁTICA e continua
  // andando junto quando a anterior mudar.
  //
  // Antes, mexer no campo já marcava a operação como fixa — inclusive quando a
  // hora era redigitada igual. Bastavam algumas edições para o posto inteiro
  // ficar travado e nada mais se reencaixar: era isso que fazia os horários
  // "não mudarem mais" depois de um tempo de uso.
  // Salvar pelo modal é o usuário assumindo esta operação: a âncora que o
  // organizador do dia tinha posto (inicioAuto) sai de cena, senão a hora
  // calculada por ele continuaria mandando sobre a que acabou de ser digitada.
  if (alvo) delete alvo.inicioAuto;
  if (alvo && !inicioFixo) {
    alvo.inicioFixo = false;
    _opSincronizarHorariosDia(data);
    // DESMARCAR a caixa é ordem de SOLTAR: a operação volta para a fila do posto
    // e passa a valer o horário que o encadeamento der, mesmo que o campo tenha
    // outra hora. Sem esta condição a regra do desvio re-fixava a operação no
    // mesmo salvamento — e uma operação já fixa não conseguia mais ser
    // desmarcada: a caixa voltava sozinha ao estado anterior.
    if (!eraFixo && alvo.inicio !== inicio) { alvo.inicio = inicio; alvo.inicioFixo = true; }
  }
  _opSincronizarHorariosDia(data);   // encadeia os horários do posto (fim de uma = início da seguinte)
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

// Devolve a operação ao encaixe automático: tira o horário fixo e reencaixa o
// posto, de modo que ela volte a começar quando a anterior termina. É a saída do
// 📌 — sem isso, uma operação travada por engano só se soltava reabrindo o modal
// e desmarcando a caixa, que quase ninguém encontra.
async function soltarHorarioOperacao(id) {
  if (!exigirAdmin('mudar o horário das operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op || !op.inicioFixo) return;
  op.inicioFixo = false;
  delete op.inicioAuto;          // solta de vez: volta para a fila do posto
  _opSincronizarHorariosDia(op.data);
  await saveState('operacoes');
  toast('Horário voltou ao encaixe automático', 'ok');
  renderOperacoes();
}

async function alternarStatusOperacao(id) {
  if (!exigirAdmin('mudar o status das operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  op.status = _OP_CICLO[_opStatus(op)];
  await saveState('operacoes');
  renderOperacoes();
  // O mesmo status é marcável pelo quadrinho da folha impressa: se ela está
  // montada, redesenha para o quadrinho refletir o clique na hora.
  const folha = document.getElementById('print-sheet-op');
  if (folha && folha.innerHTML) renderPrintPlanoOperacoes();
}

async function excluirOperacao(id) {
  if (!exigirAdmin('excluir operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  if (!confirm(`Excluir a operação "${op.operacao || 'sem descrição'}" de ${formatDate(op.data)}?`)) return;
  STATE.operacoes = (STATE.operacoes || []).filter(x => x.id !== id);
  _opSincronizarHorariosDia(op.data);   // fecha o buraco: as seguintes reencaixam
  await saveState('operacoes');
  toast('Operação excluída', 'ok');
  renderOperacoes();
}

// Duplica uma operação: cria uma cópia no MESMO posto e dia, LOGO ABAIXO da
// original. Começa quando a original termina, então não nasce sobreposta, e as
// operações que estavam depois dela reencaixam na sequência (as de horário fixo
// ficam onde estão). Copia como PENDENTE.
async function duplicarOperacao(id) {
  if (!exigirAdmin('duplicar operações')) return;
  const op = (STATE.operacoes || []).find(x => x.id === id);
  if (!op) return;
  if (!Array.isArray(STATE.operacoes)) STATE.operacoes = [];
  // Fim da ORIGINAL → início da cópia: é ela a vizinha de cima agora.
  const fimOriginal = _opFimMin(op);
  const nova = {
    ...op,
    id: uid(),
    status: 'pendente',
    inicio: fimOriginal != null ? _opHHMM(fimOriginal) : (op.inicio || ''),
    inicioFixo: false          // a cópia entra na fila do posto; a hora é calculada
  };
  STATE.operacoes.push(nova);
  // Põe a cópia IMEDIATAMENTE depois da original dentro do posto e renumera a
  // ordem de todos (a mesma mecânica das setas). Sem a original no bloco (caso
  // impossível na prática), a cópia vai para o fim.
  const blocos = _opBlocosDoDia(op.data);
  const bloco = blocos.find(b => b.itens.some(x => x.id === nova.id));
  if (bloco) {
    const k = bloco.itens.findIndex(x => x.id === nova.id);
    if (k >= 0) bloco.itens.splice(k, 1);
    const iOrig = bloco.itens.findIndex(x => x.id === op.id);
    bloco.itens.splice(iOrig < 0 ? bloco.itens.length : iOrig + 1, 0, nova);
    _opGravarOrdem(blocos);
  }
  _opSincronizarHorariosDia(op.data);   // encadeia: a cópia começa quando a anterior termina
  await saveState('operacoes');
  toast('Operação duplicada logo abaixo da original', 'ok');
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
  _opSincronizarHorariosDia(destino);   // encadeia os horários no dia de destino
  await saveState('operacoes');
  toast(`${novas.length} operação(ões) copiada(s) de ${formatDate(origem)}`, 'ok');
  renderOperacoes();
}

/* ---------------- folha impressa do plano de operações ---------------- */

// A folha que vai para o chão de fábrica. Na tela o plano é interativo (setas,
// status clicável, faixas de tempo desenhadas); no papel isso não serve — o que
// serve é a jornada de cada posto em texto, com um quadrinho para dar baixa à
// caneta. Só entram dias COM operação: dia vazio no papel é linha em branco que
// ninguém sabe se é folga ou esquecimento.
function renderPrintPlanoOperacoes() {
  const sheet = document.getElementById('print-sheet-op');
  if (!sheet) return;
  const { ini, fim } = _expRange(opPlanoModo, opPlanoAncora);
  const ops = operacoesNoPeriodo(ini, fim);
  const fmt = n => (Number(n) || 0).toLocaleString('pt-BR');

  const porDia = new Map();
  ops.forEach(op => {
    if (!porDia.has(op.data)) porDia.set(op.data, []);
    porDia.get(op.data).push(op);
  });

  const conflitos = _opConflitos(ops);
  const foraDeOrdem = _opConflitosOrdem(ops);
  const minutos = ops.reduce((s, o) => s + _opDuracao(o), 0);
  const funcoes = new Set(ops.map(_opFuncaoNome));
  const prioritarias = ops.filter(o => _opPrioridade(o) !== 'eletiva').length;

  const linha = op => {
    const pr = _opPrioridade(op);
    const resp = _opResponsavelNome(op);
    // O quadrinho do checklist é preenchível NA TELA: clicar avança o status
    // (pendente → em andamento → feita), que é a evolução do dia. No papel ele
    // continua sendo o quadrado de marcar à caneta, já com o que foi registrado.
    const st = _opStatus(op);
    const marca = st === 'feita' ? ' ok' : (st === 'andamento' ? ' meio' : '');
    return `
      <tr>
        <td class="bx"><button type="button" class="op-print-chk"
          title="${esc(_OP_STATUS[st].lbl)} — clique para avançar: pendente → em andamento → feita"
          onclick="alternarStatusOperacao('${esc(op.id)}')"><span class="exp-print-box${marca}"></span></button></td>
        <td class="jan">${esc(_opJanelaTexto(op))}</td>
        <td class="ope">${esc(op.operacao) || '—'}${
          op.escopo === 'etapa' ? ` <span class="tag">${(Array.isArray(op.etapas) ? op.etapas.length : (op.etapa ? 1 : 0)) > 1 ? 'etapas' : 'etapa'}</span>` : ''}${
          pr !== 'eletiva' ? ` <span class="tag ${pr}">${esc(_OP_PRIORIDADE[pr].lbl)}</span>` : ''}${
          conflitos.has(op.id) ? ' <span class="tag alto">sobreposta</span>' : ''}${
          _opSeloOrdem(foraDeOrdem.get(op.id), 'papel')}${
          _opForaDaJornada(op) ? ` <span class="tag alto" title="${esc(_opForaDaJornada(op))}">fora da jornada</span>` : ''}</td>
        <td class="res">${esc(resp) || '—'}</td>
        <td class="ref">${esc(op.referencia) || ''}</td>
        <td class="obs">${esc(op.obs) || ''}</td>
      </tr>`;
  };

  // Linha do tempo CONSOLIDADA do dia: todas as operações num só eixo de horas,
  // uma por linha, para se comparar todas em paralelo — é a duplicata em papel
  // das faixas que a tela mostra (lá separadas por posto/pessoa; aqui juntas).
  // As cores são as mesmas da tela (_opMapaCores), então cada operação se
  // reconhece pela cor. Operação sem horário aparece na lista, marcada, mas sem
  // barra — não há onde encaixá-la no eixo.
  const linhaTempoHtml = doDia => {
    const jan = _opJanelaDoDia(doDia);
    if (!jan) return '';   // nenhum horário no dia → não há eixo a desenhar
    const larg = jan.fim - jan.ini;
    const passo = larg <= 480 ? 60 : (larg <= 960 ? 120 : 180);
    const ticks = [];
    for (let m = jan.ini; m <= jan.fim; m += passo) {
      const pos = (m - jan.ini) / larg * 100;
      const anc = pos < 1 ? 'translateX(0)' : (pos > 99 ? 'translateX(-100%)' : 'translateX(-50%)');
      ticks.push(`<span class="tk" style="left:${pos.toFixed(3)}%;transform:${anc}">${esc(_opHHMM(m))}</span>`);
    }
    const cores = _opMapaCores(doDia);
    // Ordem CRONOLÓGICA (por horário de início): mostra a continuidade do fluxo
    // entre as funções — Corte 07h → Costura 09h → Estamparia 10h30… — em vez de
    // agrupar por posto. Operação sem horário vai para o fim.
    const ordemCron = (a, b) => {
      const ia = _opInicioMin(a), ib = _opInicioMin(b);
      if (ia == null && ib == null) return 0;
      if (ia == null) return 1;
      if (ib == null) return -1;
      return ia - ib || _opFuncaoNome(a).localeCompare(_opFuncaoNome(b));
    };
    const linhas = doDia.slice().sort(ordemCron).map(op => {
      const i = _opInicioMin(op), dur = _opDuracao(op);
      const cor = cores.get(op.id);
      // Cor da FUNÇÃO na linha inteira: faixa sólida à esquerda + fundo suave,
      // pintados como BACKGROUND (gradiente) pra não deslocar o conteúdo e manter
      // as barras alinhadas com a régua. Diferencia as operações de relance.
      const rowStyle = `background:linear-gradient(to right, ${cor} 0 4pt, ${cor}1f 4pt);`;
      const cap = `${_opFuncaoNome(op)}: ${op.operacao || '—'}`;
      if (i == null || !dur) {
        return `<div class="row" style="${rowStyle}"><div class="cap">${esc(cap)}</div><div class="track"><span class="semh">sem horário</span></div></div>`;
      }
      const left = (i - jan.ini) / larg * 100;
      const width = Math.max(1.2, dur / larg * 100);
      const conf = conflitos.has(op.id) ? ' conf' : '';
      return `<div class="row" style="${rowStyle}"><div class="cap">${esc(cap)}</div><div class="track"><div class="bar${conf}" style="left:${left.toFixed(3)}%;width:${width.toFixed(3)}%;background:${cor}" title="${esc(cap)} · ${esc(_opJanelaTexto(op))}">${esc(_opJanelaTexto(op))}</div></div></div>`;
    }).join('');
    return `
      <div class="op-print-tl no-print">
        <div class="op-print-tl-cab">Linha do tempo · operações em ordem de horário · cor por função · <span style="color:${_OP_COR_FIXO};font-weight:700;">preto = horário fixo</span></div>
        <div class="op-print-tl-regua"><div class="cap">Horário</div><div class="track">${ticks.join('')}</div></div>
        ${linhas}
      </div>`;
  };

  const blocos = Array.from(porDia.keys()).sort().map(data => {
    const doDia = porDia.get(data);
    const grupos = _opBlocosDoDia(data);
    const comHora = doDia.filter(o => _opInicioMin(o) != null);
    const abre = comHora.length ? Math.min(...comHora.map(_opInicioMin)) : null;
    const fecha = comHora.length ? Math.max(...comHora.map(_opFimMin)) : null;
    const totMin = doDia.reduce((s, o) => s + _opDuracao(o), 0);
    // Jornada real do dia e as janelas em que NENHUMA função opera — é onde a
    // continuidade entre as funções se rompe, e o papel precisa dizer isso.
    const jornadaDia = _opJornada(doDia);
    const paradaGeral = _opVazios(doDia, jornadaDia, _OP_VAZIO_MIN).filter(v => v.tipo === 'entre');
    // O que FALTA em cada lote, no papel. A folha mostrava vazios e sobreposição
    // mas nada sobre operação ausente — quem confere no chão de fábrica não tinha
    // como saber que a corrente do lote estava incompleta.
    const lotesDia = _opLotesIncompletos(doDia, data);
    const lotesFalta = lotesDia.filter(l => l.faltam.length || l.naoCabem.length);
    const lotesHtmlPapel = lotesDia.length ? `
      <div class="op-print-lotes">
        <div class="cab">${lotesDia.some(l => l.nFases > 1) ? 'Lotes do dia, por fase do enfesto' : 'Lotes do dia'} — ${lotesDia.length - lotesFalta.length} de ${lotesDia.length} com a sequência completa</div>
        ${lotesDia.map(l => `<div class="l${l.faltam.length || l.naoCabem.length ? ' falta' : ''}">
          <b>${esc(l.rotulo)}</b> ${l.feitos}/${l.total}${
            l.faltam.length ? ` · falta hoje: ${esc(l.faltam.map(p => p.nome).join(' · '))}` : ''}${
            l.naoCabem.length ? ` · continua em ${esc(formatDate(_opProximoDiaUtil(data)))}: ${esc(l.naoCabem.map(p => p.nome).join(' · '))}` : ''}${
            (!l.faltam.length && !l.naoCabem.length) ? ' · sequência completa' : ''}</div>`).join('')}
      </div>` : '';
    // Ações do dia na própria folha (não saem na impressão): daqui não dava para
    // corrigir nem analisar sem voltar à agenda, e é olhando a folha que os
    // conflitos aparecem.
    const acoesDia = `
      <div class="op-print-acoes no-print admin-only">
        <button class="btn small" onclick="corrigirOrdemOperacoes('${esc(data)}')">⇄ Organizar ${esc(formatDate(data))}</button>
        <button class="btn small" onclick="goto('operacoes'); analisarDiaOperacoes('${esc(data)}');">🔍 Analisar na agenda</button>
      </div>`;
    const funcoesHtml = grupos.map(g => {
      const gMin = g.itens.reduce((s, o) => s + _opDuracao(o), 0);
      const gCom = g.itens.filter(o => _opInicioMin(o) != null);
      const gIni = gCom.length ? Math.min(...gCom.map(_opInicioMin)) : null;
      const gFim = gCom.length ? Math.max(...gCom.map(_opFimMin)) : null;
      // Tempo em que esta função fica parada enquanto o dia corre, e o vazio
      // ENTRE duas operações virando linha própria na tabela.
      const vazios = _opVazios(g.itens, jornadaDia, _OP_VAZIO_MIN);
      const paradoMin = vazios.reduce((s, v) => s + v.min, 0);
      const linhasG = [];
      let fimAnterior = null, pessoaAnterior = '';
      g.itens.forEach(op => {
        const ini = _opInicioMin(op);
        if (fimAnterior != null && ini != null && ini - fimAnterior >= _OP_VAZIO_MIN) {
          const v = { ini: fimAnterior, fim: ini, min: ini - fimAnterior, tipo: 'entre' };
          // Na MESMA linha do vazio: o que a pessoa daquele posto está fazendo em
          // outra função nesse intervalo. É o que explica o buraco para quem lê a
          // folha no chão de fábrica.
          const pessoa = pessoaAnterior || _opResponsavelNome(op);
          const ocup = _opPessoaEmOutraFuncao(v, pessoa, g.nome, doDia);
          linhasG.push(`<tr class="vz"><td class="bx"></td><td class="jan">${esc(_opHHMM(v.ini))} ${esc(_opHHMM(v.fim))}</td>`
            + `<td class="ope" colspan="4">— ${esc(_opDurTexto(v.min))} sem operação —${
              ocup.length ? ` &nbsp;${esc(pessoa)} em: ${esc(_opOcupacaoTexto(ocup))}` : ''}</td></tr>`);
        }
        linhasG.push(linha(op));
        if (ini != null && _opDuracao(op)) {
          fimAnterior = Math.max(fimAnterior == null ? 0 : fimAnterior, _opFimMin(op));
          pessoaAnterior = _opResponsavelNome(op) || pessoaAnterior;
        }
      });
      // Somatório da função, fechando o quadro dela: quantas operações, quanto
      // tempo somam e quanto o posto fica parado. O cabeçalho traz os mesmos
      // números em letra miúda, mas quem confere lê a coluna de baixo para cima —
      // o total tem que estar onde a lista termina.
      const totalLinha = `<tr class="tot">
        <td class="bx"></td>
        <td class="jan">${gIni != null ? esc(_opHHMM(gIni)) + ' ' + esc(_opHHMM(gFim)) : ''}</td>
        <td class="ope" colspan="4">Total da função: <b>${esc(_opDurTexto(gMin))}</b> em ${g.itens.length} ${g.itens.length === 1 ? 'operação' : 'operações'}${
          paradoMin > 0 ? ` · <b>${esc(_opDurTexto(paradoMin))}</b> parada` : ''}${
          gIni != null ? ` · presença ${esc(_opDurTexto(gFim - gIni))}` : ''}</td>
      </tr>`;
      return `
        <div class="op-print-posto">
          <div class="ph">
            <span class="t">${esc(g.nome)}</span>
            <span class="h">${gIni != null ? esc(_opHHMM(gIni)) + ' → ' + esc(_opHHMM(gFim)) : 'sem horário'} · ${esc(_opDurTexto(gMin))}${
              paradoMin > 0 ? ' · ' + esc(_opDurTexto(paradoMin)) + ' parada' : ''}</span>
          </div>
          <table>${linhasG.join('')}${totalLinha}</table>
        </div>`;
    }).join('');
    return `
      <div class="exp-print-bloco">
        <div class="cab">
          <span class="d">${_EXP_DIAS_CURTO[_expData(data).getDay()]} ${esc(formatDate(data))}</span>
          <span class="j">${grupos.length} ${grupos.length === 1 ? 'função' : 'funções'} · ${esc(_opDurTexto(totMin))}${
            abre != null ? ` · jornada ${esc(_opHHMM(abre))} → ${esc(_opHHMM(fecha))}` : ''}</span>
        </div>
        ${paradaGeral.length ? `<div class="op-print-parada">Sem nenhuma função operando: ${
          paradaGeral.map(v => `${esc(_opHHMM(v.ini))} → ${esc(_opHHMM(v.fim))} (${esc(_opDurTexto(v.min))})`).join(' · ')}</div>` : ''}
        ${acoesDia}
        ${linhaTempoHtml(doDia)}
        ${funcoesHtml}
        ${lotesHtmlPapel}
      </div>`;
  }).join('');

  const emissao = new Date();
  const emissaoTxt = formatDate(_expIso(emissao)) + ' '
    + String(emissao.getHours()).padStart(2, '0') + ':' + String(emissao.getMinutes()).padStart(2, '0');

  // Mesma tabela de uma coluna da folha de OE: o <thead> é o que faz o navegador
  // repetir o título no alto de cada folha impressa.
  sheet.innerHTML = `
    <table class="exp-print-folha">
      <thead>
        <tr><td>
          <div class="exp-print-head">
            <div>
              <div class="tit">PLANEJAMENTO DE OPERAÇÕES</div>
              <div class="sub">Plano ${esc(_expNomeModo(opPlanoModo))} · ${esc(formatDate(ini))} a ${esc(formatDate(fim))} · jornada por função</div>
            </div>
            <div class="meta">
              <div>Emitido em ${esc(emissaoTxt)}</div>
            </div>
          </div>
        </td></tr>
      </thead>
      <tbody><tr><td>
    <div class="exp-print-resumo">
      <div class="item"><div class="n">${fmt(ops.length)}</div><div class="l">Operações</div></div>
      <div class="item"><div class="n">${fmt(funcoes.size)}</div><div class="l">Funções</div></div>
      <div class="item"><div class="n">${esc(_opDurTexto(minutos))}</div><div class="l">Tempo planejado</div></div>
      <div class="item"><div class="n">${fmt(prioritarias)}</div><div class="l">Urgentes / emerg.</div></div>
      <div class="item"><div class="n">${fmt(porDia.size)}</div><div class="l">Dias com operação</div></div>
    </div>
    <div style="font-size:7pt;color:#555;margin:3pt 0 5pt;">
      Quadrinho de cada operação: <span class="exp-print-box"></span> pendente ·
      <span class="exp-print-box meio"></span> em andamento ·
      <span class="exp-print-box ok"></span> feita.
      Na tela, clique no quadrinho para avançar o status; no papel, marque à caneta.
    </div>
    ${blocos || `<div style="padding:20px 0;text-align:center;font-size:9pt;font-style:italic;">Nenhuma operação planejada ${esc(_EXP_VAZIO_PERIODO[opPlanoModo] || 'neste período')}.</div>`}
        <div class="exp-print-rodape">
          <div class="ass"><div class="linha"></div><div class="lbl">Encarregado de produção</div></div>
          <div class="ass"><div class="linha"></div><div class="lbl">Conferido por</div></div>
        </div>
      </td></tr></tbody>
      <!-- Mesmo rodapé vazio da folha de OE: o espaço da margem inferior em
           cada folha impressa (ver o comentário lá). -->
      <tfoot><tr><td></td></tr></tfoot>
    </table>`;
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
    if (op.escopo == null) { op.escopo = 'completa'; mudou = true; }
    // Migração: etapa única (op.etapa) → múltiplas (op.etapas).
    if (!Array.isArray(op.etapas)) { op.etapas = op.etapa ? [op.etapa] : []; mudou = true; }
    if ('etapa' in op) { delete op.etapa; mudou = true; }
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
  // MIGRAÇÃO ÚNICA: até aqui o organizador do dia marcava `inicioFixo` em tudo
  // que ele movia (pausas, empurrões, correções de ordem). Isso pintava o 📌 de
  // "horário fixo pelo usuário" em operação que ninguém fixou, e a agenda ficava
  // toda travada. Essas âncoras passam a ser `inicioAuto`, que segura o horário
  // do mesmo jeito mas não é apresentada como escolha de ninguém. Roda uma vez
  // só (marcador em meta): depois disso, `inicioFixo` só existe quando o usuário
  // marca a caixa ou digita uma hora fora do encaixe.
  if (!STATE.meta || typeof STATE.meta !== 'object') STATE.meta = {};
  if (!STATE.meta.opFixoMigrado) {
    let convertidas = 0;
    lista.forEach(op => {
      if (!op.inicioFixo) return;
      op.inicioFixo = false;
      op.inicioAuto = true;
      convertidas++;
    });
    STATE.meta.opFixoMigrado = _expIso(new Date());
    _opFixoMigradoAgora = convertidas;
    mudou = true;
  }
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
  const TH = 'padding:0 2px;font-weight:700;border-bottom:.5pt solid #999;';
  const TD = 'padding:0 2px;text-align:center;font-family:\'IBM Plex Mono\',monospace;';

  // Tabela de quantidades de uma carga PARCIAL: os números são só o que vai
  // NESTA viagem, mas a tabela mantém a FORMA DA GRADE — todos os tamanhos (e
  // todas as tonalidades) da OS aparecem, e o pacote que não vai aparece com
  // ZERO. Assim quem confere na doca vê que aquele tamanho não foi, em vez de não
  // encontrar a coluna e ter que adivinhar. As peças de cada pacote saem de
  // _expPecasPacoteOS, que lê o mesmo "Total por tamanho" da folha de OS.
  const tabelaDaCarga = (o, carga) => {
    const cont = _expContarPacotes(carga.pacotes);
    if (!cont.size) return '';    // carga sem pacote de tamanho: quem chama escreve o aviso
    const pp = _expPecasPacoteOS(o);
    const ordem = ['P', 'M', 'G', 'GG', 'G1', 'G2', 'G3'];
    const vals = Array.from(cont.values());
    // Colunas: os tamanhos da grade da OS + qualquer tamanho que esteja na carga
    // e não esteja mais na grade (grade alterada depois de alocar).
    const daGrade = [];
    _tamanhosDaGradeExpandido(o).forEach(t => { if (!daGrade.includes(t)) daGrade.push(t); });
    const tams = ordem.filter(t => daGrade.includes(t) || vals.some(e => e.tam === t));
    vals.forEach(e => { if (!tams.includes(e.tam)) tams.push(e.tam); });   // tamanho fora da ordem conhecida
    if (!tams.length) return '';
    // Linhas: as tonalidades da OS (mesma regra — tom que não vai fica zerado).
    const tons = (totaisPorTamanhoTomOS(o).tons || []).slice();
    vals.forEach(e => { const t = e.tom == null ? null : e.tom; if (!tons.includes(t)) tons.push(t); });
    if (!tons.length) tons.push(null);
    tons.sort((a, b) => (a == null ? -1 : (b == null ? 1 : a - b)));
    const cel = (tam, tom) => cont.get(_expChavePacote({ tam, tom })) || null;
    const pecasDe = (tam, tom) => {
      const e = cel(tam, tom);
      return e ? Math.round(e.qtd * pp.de({ tam, tom })) : 0;
    };
    // Zero em cinza: o número que importa é o que vai: o zero fica legível sem
    // disputar a atenção com as quantidades reais.
    const celHtml = (tam, tom) => {
      const e = cel(tam, tom);
      const p = pecasDe(tam, tom);
      if (!e || !(p > 0)) return `<td style="${TD}color:#999;">0</td>`;
      return `<td style="${TD}">${fmt(p)}${e.qtd > 1 ? `<span style="font-size:.8em"> (${e.qtd}×)</span>` : ''}</td>`;
    };
    const totalTam = tam => tons.reduce((s, tom) => s + pecasDe(tam, tom), 0);
    const totalGeral = tams.reduce((s, tam) => s + totalTam(tam), 0);
    const totCel = (v, fundo) => `<td style="${TD}font-weight:700;${fundo ? 'background:#eef3ee;' : ''}${v > 0 ? '' : 'color:#999;'}">${fmt(v)}</td>`;
    const cabec = tams.map(t => `<th style="${TH}text-align:center;">${esc(t)}</th>`).join('');
    // Com um tom só a linha do tom é a própria linha do total: repetir o mesmo
    // número duas vezes só faz procurar diferença que não existe.
    const umTomSo = tons.length <= 1;
    const rotTotal = umTomSo
      ? (tons[0] == null ? 'Nesta carga' : `Tom ${tons[0]}`)
      : 'Total da carga';
    const linhasTom = umTomSo ? '' : tons.map(tom => `
      <tr>
        <td style="${TD}text-align:left;white-space:nowrap;">${tom == null ? 'Sem tom' : 'Tom ' + tom}</td>
        ${tams.map(tam => celHtml(tam, tom)).join('')}
        ${totCel(tams.reduce((s, tam) => s + pecasDe(tam, tom), 0), true)}
      </tr>`).join('');
    return `
      <table>
        <thead>
          <tr>
            <th style="${TH}text-align:left;width:34pt;">Pacotes</th>
            ${cabec}
            <th style="${TH}text-align:center;background:#eef3ee;">Total</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="${TD}text-align:left;white-space:nowrap;font-weight:700;">${esc(rotTotal)}</td>
            ${tams.map(tam => totCel(totalTam(tam), false)).join('')}
            ${totCel(totalGeral, true)}
          </tr>
          ${linhasTom}
        </tbody>
      </table>`;
  };

  const osPrint = (i) => {
    const o = i.os;
    // Cor do DESENHO TÉCNICO — a MESMA do banner da folha de OS (coresDaPecaOS),
    // com TODAS as cores da peça: um tricolor sai "Preto / Mostarda / Off-White",
    // não só a Cor 1. Antes esta linha lia apenas cor1Nome da 1ª variante e a OS
    // tricolor chegava na doca anunciada como se fosse preta. Sem variante (OS
    // antiga), cai na cor da 1ª fase do enfesto. Vai na 1ª LINHA, junto do modelo.
    const corPred = o
      ? (coresDaPecaOS(o).join(' / ')
         || corNomeCurto((o.fases || [])[0]?.corNome || (o.tecidos || [])[0]?.corNome || ''))
      : '';
    const volTxt = i.volumes > 0 ? fmt(i.volumes) + ' vol' : '— vol';
    // O quadrinho é o avanço da carga: marcado = esta OS já foi feita (separada e
    // embarcada). Clicável na tela, para a folha aberta ir mostrando o que já
    // andou; no papel continua sendo o quadrado de marcar à caneta.
    const feita = !!(i.carga && i.carga.feita);
    const box = i.carga
      ? `<button type="button" class="op-print-chk" title="${feita ? 'Feita — clique para desmarcar' : 'Marcar como feita (separada e embarcada)'}"
          onclick="alternarCargaFeita('${esc(i.carga.id)}')"><span class="exp-print-box${feita ? ' ok' : ''}"></span></button>`
      : '<span class="exp-print-box"></span>';
    // 1ª linha: nº da OS + modelo + cor. A cor é span próprio (fora do .m, que tem
    // ellipsis) pra não ser cortada quando o modelo é longo.
    const cab = `
      <div class="cab">
        ${box}
        <span class="n">${esc(i.osNumero)}</span>
        <span class="m">${esc(i.modelo)}</span>
        ${corPred ? `<span class="cor">${esc(corPred)}</span>` : ''}
      </div>`;
    const TT = o ? totaisPorTamanhoTomOS(o) : null;
    // Sem grade: ao menos o volume abaixo da 1ª linha.
    if (!TT || !TT.tamanhos.length) return `<div class="exp-print-os">${cab}<div class="sub">${fmt(i.pecas)} pç · ${volTxt}</div></div>`;

    // A conta do volume, escrita por extenso: é a mesma regra do planejamento
    // (nº de tamanhos × tonalidades + 1 de reposição). Divergência contra o que
    // está alocado na carga fica à vista em vez de virar surpresa na doca.
    const nTam = _expTotalTamanhosGrade(o);
    const nTons = Math.max(1, TT.tons.length);
    const volCalc = nTam > 0 ? nTam * nTons + 1 : 0;
    // Lote parcial (carga com composição por pacote) tem uma nota própria abaixo;
    // não mostra a divergência genérica, que aqui é esperada e já explicada.
    const ehParcialCarga = Array.isArray(i.carga && i.carga.pacotes);
    const diverge = volCalc > 0 && i.volumes > 0 && volCalc !== i.volumes && !ehParcialCarga;
    const canonTotal = ehParcialCarga ? _expPacotesCanonicos(o).total : 0;
    const nestaCarga = ehParcialCarga ? (i.carga.pacotes.length + (i.carga.reposicao ? 1 : 0)) : 0;
    const cargaParcial = ehParcialCarga && canonTotal > 0 && nestaCarga < canonTotal;

    // Carga parcial: o quadro descreve A CARGA, não o lote inteiro. A tabela traz
    // só os pacotes embarcados; os ausentes não são relacionados aqui (ficam no
    // Estoque de corte, disponíveis para a próxima carga). Sem aviso de "parcial":
    // com a tabela mostrando exatamente o que vai, a folha não precisa alertar
    // sobre o que não vai — o que está no papel é o que está no caminhão.
    if (cargaParcial) {
      const tab = tabelaDaCarga(o, i.carga);
      const soRep = !i.carga.pacotes.length && i.carga.reposicao;
      return `
        <div class="exp-print-os">
          ${cab}
          <div class="sub">
            ${fmt(i.pecas)} pç · <b>${fmt(nestaCarga)} volume${nestaCarga === 1 ? '' : 's'}</b> nesta carga${i.carga.reposicao ? ' (com o de reposição e ribana)' : ''}
          </div>
          ${tab || `<div class="pe">${soRep ? 'Só o pacote de reposição e ribana nesta carga.' : 'Nenhum pacote de tamanho nesta carga.'}</div>`}
        </div>`;
    }
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
          ${fmt(i.pecas)} pç · ${contaVol}${diverge ? ` · <b>carga alocada com ${fmt(i.volumes)} vol</b>` : ''}
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
        </table>${indef ? `
        <div class="pe">A divisão entre as tonalidades ainda não foi repartida na OS.</div>` : ''}
      </div>`;
  };

  const pernaPrint = (oc, perna) => {
    const r = resumoPernaExpedicao(oc, perna);
    const hora = perna === 'ida' ? oc.horaIda : oc.horaVolta;
    const linhas = r.itens.length
      ? r.itens.map(osPrint).join('')
      : '<div class="vazia">Sem OS alocada.</div>';
    // Cada perna ocupa metade da folha, cheia ou vazia — a largura fixa é o que
    // mantém a ida sempre na mesma metade, folha após folha.
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

  // A folha inteira vive dentro de uma TABELA de uma coluna só, com o cabeçalho
  // no <thead>. É o único mecanismo especificado que faz o navegador repetir um
  // trecho no alto de CADA folha impressa — tentar com position:fixed já foi
  // feito e o título acabou no rodapé. Visualmente nada muda: a tabela não tem
  // bordas nem espaçamento próprio.
  sheet.innerHTML = `
    <table class="exp-print-folha">
      <thead>
        <tr><td>
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
        </td></tr>
      </thead>
      <tbody><tr><td>
    <div class="exp-print-resumo">
      <div class="item"><div class="n">${fmt(ativas)}</div><div class="l">Expedições</div></div>
      <div class="item"><div class="n">${fmt(volIda)}</div><div class="l">Volumes ida</div></div>
      <div class="item"><div class="n">${fmt(volVolta)}</div><div class="l">Volumes volta</div></div>
      <div class="item"><div class="n">${fmt(volIda + volVolta)}</div><div class="l">Volumes total</div></div>
      <div class="item"><div class="n">${fmt(pecasTot)}</div><div class="l">Peças</div></div>
      <div class="item"><div class="n">${fmt(osTot.size)}</div><div class="l">OS alocadas</div></div>
    </div>
    <div style="font-size:7pt;color:#555;margin:3pt 0 5pt;">
      Quadrinho na frente de cada OS: <span class="exp-print-box"></span> a fazer ·
      <span class="exp-print-box ok"></span> feita (separada e embarcada).
      Na tela, clique no quadrinho para registrar o avanço; no papel, marque à caneta.
    </div>
    ${ocs.length ? blocos : `<div style="padding:20px 0;text-align:center;font-size:9pt;font-style:italic;">Nenhuma Ordem de Expedição produzida ${esc(_EXP_VAZIO_PERIODO[expPlanoModo] || 'neste período')}.</div>`}
        <div class="exp-print-rodape">
          ${_EXP_ASSINATURAS.map(f => `<div class="ass"><div class="linha"></div><div class="lbl">${esc(f)}</div></div>`).join('')}
        </div>
      </td></tr></tbody>
      <!-- Rodapé VAZIO de propósito: é só o espaço da margem inferior. O <tfoot>
           é o par do <thead> — o navegador o repete no pé de CADA folha impressa,
           que é o único jeito de a última linha não encostar na borda do papel
           quando a folha ocupa várias páginas. A altura vem do @media print. -->
      <tfoot><tr><td></td></tr></tfoot>
    </table>`;
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
  if (!exigirEdicao('renomear pastas de grade')) return;
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
  if (!exigirEdicao('renomear pastas de grade')) return;
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
  if (!exigirEdicao('reordenar pastas de grade')) return;
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
  if (!exigirEdicao('reordenar pastas de grade')) return;
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

// As opções de PASTA (tipo de peça) e SUBPASTA (variação) de uma grade, num
// lugar só. Toda tela que cria ou edita grade usa esta lista — o cadastro
// manual, o "Importar risco (PDF)" e o assistente de pasta.
//
// Isso não é organização: era um erro de destino. As telas de importação tinham
// a lista chumbada em três tipos (camiseta, blusa moletom, outro) e três
// variações, enquanto a casa já usa "Camiseta Polo", "Camiseta Oversized",
// "Bermuda Tactel", "Bermuda Moletom", "Jaguar", "Prime", "Rugão", "Espartana".
// Uma grade de PM.LISA criada pelo risco caía em `camiseta` — a pasta das CM —
// porque não havia como dizer outra coisa. A pasta vem de tipoPeca/variação, não
// do SKU: escrever PM.LISA no SKU não muda de pasta.
//
// A lista sai de três fontes, na ordem: as fixas (com o nome que a casa deu a
// elas, se renomearam), as que já existem em alguma grade, e "+ Nova pasta…",
// que pergunta o nome na hora.
function opcoesPastaGrade(kind, atual) {
  const pasta = kind === 'pasta';
  const fixas = pasta
    ? [{ v: '', lbl: labelTp('') === 'Sem categoria' ? '— sem categoria —' : labelTp('') },
       { v: 'camiseta', lbl: labelTp('camiseta') },
       { v: 'blusa_moletom', lbl: labelTp('blusa_moletom') },
       { v: 'outro', lbl: labelTp('outro') }]
    : [{ v: '', lbl: labelVr('') === 'Sem variação' ? '— sem variação —' : labelVr('') },
       { v: 'basica', lbl: labelVr('basica') },
       { v: 'bicolor', lbl: labelVr('bicolor') },
       { v: 'tricolor', lbl: labelVr('tricolor') }];
  const fixos = new Set(fixas.map(t => t.v));
  const usados = [...new Set((STATE.grades || [])
    .map(g => (pasta ? g.tipoPeca : g.variacao) || '')
    .filter(x => x && !fixos.has(x)))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  // O valor atual aparece mesmo que ainda não exista em nenhuma grade salva.
  if (atual && !fixos.has(atual) && !usados.includes(atual)) usados.push(atual);
  const rot = pasta ? ['Pastas adicionais', '+ Nova pasta…'] : ['Subpastas adicionais', '+ Nova subpasta…'];
  return fixas.map(t => `<option value="${esc(t.v)}" ${atual === t.v ? 'selected' : ''}>${esc(t.lbl)}</option>`).join('')
    + (usados.length ? `<optgroup label="${rot[0]}">${usados.map(v =>
        `<option value="${esc(v)}" ${atual === v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</optgroup>` : '')
    + `<option value="__nova__">${rot[1]}</option>`;
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

  // admin-only: renomear e reordenar pasta é escrita. Abrir e fechar a pasta
  // (o clique na linha) continua de todo mundo — é como se lê a lista.
  const folderActions = (clickAttrs) => `<span class="folder-actions admin-only" onclick="event.stopPropagation()">${clickAttrs}</span>`;

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
  // Mesma regra do acoesCell: "ver" para quem consulta, o resto admin-only.
  const acoesEtapa = (id) => `
    <span class="row-actions" style="display:inline-flex;gap:4px;">
      <button class="edit leitura-only" onclick="openCadastroModal('etapa','${esc(id)}')">ver</button>
      <button class="edit admin-only" onclick="openCadastroModal('etapa','${esc(id)}')">editar</button>
      <button class="edit admin-only" onclick="duplicarCadastro('etapa','${esc(id)}')">duplicar</button>
      <button class="del admin-only" onclick="excluirCadastro('etapa','${esc(id)}')">excluir</button>
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
    // Cada operação com o seu tempo ao lado; sem tempo, só o nome.
    const ops = _opsDaFuncao(f);
    const opsHtml = ops.length
      ? ops.map(o => `<span class="badge" style="margin-right:4px" title="${o.duracaoMin ? 'Tempo cadastrado: ' + esc(_opDurTexto(o.duracaoMin)) : 'Sem tempo cadastrado'}">${esc(o.nome)}${
          o.duracaoMin ? ` · <b>${esc(_opDurTexto(o.duracaoMin))}</b>` : ''}</span>`).join('')
      : '—';
    return `<tr><td><strong>${esc(f.nome)}</strong></td><td>${esc(f.desc)||'—'}</td><td>${opsHtml}</td>${acoesCell('funcao', f.id)}</tr>`;
  }).join('');
}


function renderHome() {
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  set('stat-os', STATE.ordens.length);
  // OEs cadastradas = expedições distintas (janela + data) com pelo menos uma OS
  // alocada. Uma OE é uma expedição montada, não um registro à parte — por isso
  // é contada pelas cargas, agrupadas por janela+data.
  const _oes = new Set((STATE.expedicaoCargas || []).map(c => (c.janelaId || '') + '|' + (c.data || '')));
  set('stat-oes', _oes.size);
  set('stat-operacoes', (STATE.operacoes || []).length);
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

// Na NOVA OS, as camadas nascem no MÁXIMO do tecido (moletom 36 / malha 80) e as
// peças-alvo saem daí (camadas × grade × mult) — não mais o 160 fixo, que dava
// 160 camadas no moletom, muito acima do limite de 36. Só age quando as camadas
// ainda não foram definidas (campo vazio) e não é edição de OS salva, pra não
// atropelar valor do usuário nem o carregado. É chamado ao aplicar a grade,
// quando o tecido e a grade já estão conhecidos.
function _aplicarCamadasMaximasDefault() {
  if (osEditId) return;                                     // edição: respeita o salvo
  const campoCam = document.getElementById('f-enf-camadas');
  if (!campoCam || (campoCam.value || '').trim() !== '') return;  // já definido: não mexe
  const { limite } = calcularLimiteCamadas();
  if (!(limite > 0) || limite === Infinity) return;         // tecido sem limite conhecido
  const temGrade = ['p','m','g','gg','g1','g2','g3'].some(k => (parseInt(document.getElementById('f-gr-'+k)?.value) || 0) > 0);
  if (!temGrade) return;                                     // sem grade não dá pra derivar as peças
  campoCam.value = limite;                                   // camadas = máx do tecido
  calcularAlvoDeCamadas();                                   // deriva peças-alvo = camadas × grade × mult
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

  _aplicarCamadasMaximasDefault();   // camadas no máx do tecido + peças-alvo derivadas (OS nova)
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
// Unidades da grade que uma camada de FORRO rende quando a grade não diz outra
// coisa. O 2 é a regra que sempre valeu — "o forro enfesta com metade das
// camadas do moletom" — e continua sendo o padrão: grade cadastrada antes deste
// campo existir calcula exatamente como calculava.
const UNIDADES_PADRAO_FORRO = 2;
const LABEL_CATEGORIA = { malha: 'Malha algodão', moletom: 'Moletom', ribana: 'Ribana', outro: 'Outro' };

/* --------- REGRA DAS BOBINAS: malha algodão, pelo comprimento --------- */
// Quantas bobinas um enfesto de malha algodão consome é coisa que a casa mede na
// prática, não que se calcule da área: a tabela abaixo é a experiência de quem
// enfesta. Ela não é proporcional — de 5 para 6 metros pula duas bobinas, de 6
// para 7 pula uma — e por isso não dá para trocar por uma multiplicação.
//
// SÓ VALE PARA MALHA ALGODÃO. Moletom e ribana consomem de outro jeito e ficam
// de fora: nesses o campo continua em branco, para quem sabe preencher.
const BOBINAS_MALHA = { 2: 5, 3: 6, 4: 7, 5: 8, 6: 10, 7: 11, 8: 13, 9: 15 };
const BOBINAS_MALHA_MIN_M = 2;    // abaixo disso, o piso da tabela
const BOBINAS_MALHA_MAX_M = 9;    // acima disso, segue o ritmo do fim dela
const BOBINAS_MALHA_PASSO = 2;    // bobinas por metro depois dos 9 m (13→15)

// O comprimento arredonda PARA CIMA ao metro inteiro. Enfesto real quase nunca é
// metro redondo — 4,70 m usa o que 5 m usa. Prever a menos é o erro caro: falta
// pano no meio do enfesto.
function bobinasPrevistasMalha(comp) {
  const c = parseFloat(String(comp == null ? '' : comp).replace(',', '.'));
  if (!isFinite(c) || c <= 0) return null;
  // A folga de 1e-6 evita que 4,00 vindo de uma conta com resto (4.0000000001)
  // seja tratado como se passasse dos 4 metros.
  const m = Math.max(BOBINAS_MALHA_MIN_M, Math.ceil(c - 1e-6));
  if (m <= BOBINAS_MALHA_MAX_M) return BOBINAS_MALHA[m];
  return BOBINAS_MALHA[BOBINAS_MALHA_MAX_M] + (m - BOBINAS_MALHA_MAX_M) * BOBINAS_MALHA_PASSO;
}

// A sugestão para uma fase: só existe quando o tecido dela é malha algodão.
function sugestaoBobinasFase(tecidoId, comp) {
  const t = (STATE.tecidos || []).find(x => x.id === tecidoId);
  if (!t || categoriaEfetivaTecido(t) !== 'malha') return null;
  return bobinasPrevistasMalha(comp);
}

// O texto que explica de onde saiu o número, para ninguém achar que é chute.
function textoRegraBobinas(comp) {
  const c = parseFloat(String(comp == null ? '' : comp).replace(',', '.'));
  const n = bobinasPrevistasMalha(c);
  if (n == null) return '';
  const m = Math.max(BOBINAS_MALHA_MIN_M, Math.ceil(c - 1e-6));
  return `${n} bobinas — regra da malha algodão para ${m} m`
       + (Math.abs(m - c) > 1e-6 ? ` (${String(c.toFixed(2)).replace('.', ',')} m arredondado para cima)` : '');
}

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
    if (!tec) return;
    // categoriaEfetivaTecido, não tec.categoria cru: as ribanas do cadastro têm
    // categoria VAZIA e são reconhecidas pelo nome. Lendo o campo cru, elas eram
    // puladas aqui — o limite saía calculado só com os demais tecidos, enquanto o
    // cálculo das camadas e o multiplicador já as tratavam como ribana. Duas
    // fontes de verdade para a mesma pergunta.
    const cat = categoriaEfetivaTecido(tec);
    if (!cat) return;
    const lim = LIMITE_CAMADAS[cat];
    if (!(lim > 0)) return;
    if (lim < limite) { limite = lim; categoriaRestritiva = cat; }
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

  // O multiplicador da peça principal (malha: 1 camada = 2 peças) sai dos
  // TECIDOS ESCOLHIDOS NA OS, não das fases da grade cadastrada. A grade é um
  // molde de tamanhos e pode perfeitamente não ter fases, ou tê-las sem tecido;
  // quando isso acontecia o multiplicador caía para 1 em silêncio e as camadas
  // DOBRAVAM — 160 no lugar de 80, produzindo o dobro do alvo pedido.
  // Os tecidos da OS já são a fonte do limite de camadas (calcularLimiteCamadas)
  // e da contagem de peças (multiplicadorPecaOS): agora as três concordam.
  // Só cai nas fases da grade se a OS ainda não tem tecido escolhido.
  const catsForm = Array.from(document.querySelectorAll('#tecidos-rows .tecido-row'))
    .map(r => r.querySelector('.tec-sel')?.value)
    .filter(Boolean)
    .map(id => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === id)))
    .filter(Boolean);
  const cats = catsForm.length
    ? catsForm
    : fases.map(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId))).filter(Boolean);
  const temMoletom = cats.includes('moletom');
  const temMalha = cats.includes('malha');

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

  // Papéis das fases. As fases da grade definem o PAPEL de cada bloco de enfesto
  // (moletom, forro de capuz, ribana…) e, por tabela, as camadas de cada um.
  // Quando a grade não tem fases, papeis[i] vinha vazio e o bloco caía no ramo
  // genérico com multiplicador 1: o campo de camadas recebia `target / 1`, ou
  // seja, o PRÓPRIO NÚMERO DE PEÇAS-ALVO no lugar das camadas (160 em vez de 80).
  // Sem fases na grade, monta-se as fases a partir da própria OS: o tecido vem
  // das linhas de tecido (que têm o id real) e o nome vem do bloco de enfesto
  // (que é o que identifica o Viés).
  // A complementação é POR FASE, não tudo-ou-nada: a grade pode ter as fases e
  // ainda assim deixar o tecido em branco em alguma delas, e aí só aquela caía
  // no ramo genérico. Cada fase usa o tecido da grade quando ele existe de fato,
  // e o da linha de tecido da OS quando não.
  const blocosParaNome = document.querySelectorAll('#f-enfestos-blocos .enfesto-bloco');
  const rowsTec = Array.from(document.querySelectorAll('#tecidos-rows .tecido-row'))
    .map(r => r.querySelector('.tec-sel')?.value || '');
  const nFases = Math.max(fases.length, rowsTec.length, blocosParaNome.length);
  const fasesEfetivas = Array.from({ length: nFases }, (_, i) => {
    const f = fases[i] || {};
    const idValido = f.tecidoId && STATE.tecidos.some(t => t.id === f.tecidoId);
    return {
      ...f,
      ordem: f.ordem || (i + 1),
      nome: (f.nome && f.nome.trim()) ? f.nome : (blocosParaNome[i]?.dataset?.nomeTecido || ''),
      tecidoId: idValido ? f.tecidoId : (rowsTec[i] || '')
    };
  });
  const papeis = calcularPapeisFases(fasesEfetivas);

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
    const fase = fasesEfetivas[i] || {};
    const ehVies = /vi[eé]s/i.test(fase.nome || '') || /vi[eé]s/i.test(papel.label || '') || /vi[eé]s/i.test(bloco.dataset.nomeTecido || '');
    let val;
    if (ehVies) {
      val = 1;
    } else if (papel.papel === 'moletom') {
      // Enfesto moletom: todos componentes moletom na mesma camada → 1 camada = 1 blusa
      val = camadasPrincipal;
    } else if (papel.papel === 'forro_capuz') {
      // Enfesto forro: uma camada rende N unidades da grade, então são N vezes
      // menos camadas que o moletom. N vem do cadastro da fase; sem ele, 2 —
      // que é a metade de sempre.
      const unidForro = parseInt(fase.unidades, 10) || UNIDADES_PADRAO_FORRO;
      val = Math.max(1, Math.ceil(camadasPrincipal / unidForro));
    } else if ((papel.papel || '').startsWith('ribana_')) {
      // Enfesto ribana com unidades cadastradas:
      // - Ribana moletom: o tecido escala com a unidade da grade (2 barras +
      //   4 punhos por tamanho cobre 2 blusas/tam quando a grade tem 2/tam).
      //   Formula: camadasPrincipal × multPrincipal × (gradeTotal/n_tamanhos) / unidades.
      // - Outras ribanas (malha algodao, gola polo): o multiplicador da fase
      //   ja cobre toda a grade — "10x" significa "10 unidades de cada slot
      //   da grade por camada", entao o gradeTotal nao precisa de ajuste.
      //   Formula: camadasPrincipal × multPrincipal / unidades.
      const fase = fasesEfetivas[i] || {};
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

// Sentido INVERSO: o usuário digita o Nº de camadas e as Peças-alvo são
// calculadas sozinhas. É o espelho de calcularCamadasParaProducao: lá
// camadas = ceil(alvo / (minQtd × multPrincipal)); aqui alvo = camadas ×
// minQtd × multPrincipal — o round-trip é exato (digitar as camadas
// resultantes de um alvo devolve o mesmo número de camadas). minQtd e
// multPrincipal são lidos exatamente como no cálculo direto, para as duas
// contas concordarem.
function calcularAlvoDeCamadas() {
  const camadas = parseInt(document.getElementById('f-enf-camadas')?.value) || 0;
  // Sem camadas: nada a inferir — só atualiza o display e preserva o alvo
  // atual (o usuário pode estar apagando para redigitar).
  if (camadas <= 0) { atualizarCalculosEnfesto(); return; }
  const qtdsPorTamanho = ['p','m','g','gg','g1','g2','g3']
    .map(k => parseInt(document.getElementById('f-gr-'+k)?.value) || 0)
    .filter(q => q > 0);
  // Sem grade não há minQtd: não dá para inferir o alvo. Silencioso (roda a
  // cada tecla) — só atualiza o display.
  if (!qtdsPorTamanho.length) { atualizarCalculosEnfesto(); return; }
  const minQtd = Math.min(...qtdsPorTamanho);
  // multPrincipal: mesma regra do cálculo direto — malha SEM moletom corta em
  // camada dupla (2 peças/camada por slot); caso contrário 1. Fonte: tecidos da
  // OS; se ainda não houver, cai nas fases da grade.
  const gradeId = document.getElementById('f-grade-preset')?.value;
  const grade = gradeId ? STATE.grades.find(g => g.id === gradeId) : null;
  const fases = grade?.fases || [];
  const catsForm = Array.from(document.querySelectorAll('#tecidos-rows .tecido-row'))
    .map(r => r.querySelector('.tec-sel')?.value)
    .filter(Boolean)
    .map(id => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === id)))
    .filter(Boolean);
  const cats = catsForm.length
    ? catsForm
    : fases.map(f => categoriaEfetivaTecido(STATE.tecidos.find(t => t.id === f.tecidoId))).filter(Boolean);
  const temMoletom = cats.includes('moletom');
  const temMalha = cats.includes('malha');
  let multPrincipal = 1;
  if (!temMoletom && temMalha) multPrincipal = MULTIPLICADOR_PECAS.malha || 2;

  const alvo = camadas * minQtd * multPrincipal;
  const inputTarget = document.getElementById('f-enf-target');
  if (inputTarget) inputTarget.value = alvo;
  // Com o alvo preenchido, reaproveita o cálculo direto para distribuir as
  // camadas nos blocos (moletom/forro/ribana) e atualizar os totais — o mesmo
  // resultado de quem tivesse digitado o alvo. Como o round-trip é exato, o
  // Nº de camadas que o usuário digitou permanece igual.
  calcularCamadasParaProducao();
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
  // Re-deriva a COR de cada bloco pela linha de Tecidos (canônica, do desenho),
  // pra o dado GRAVADO não ficar com um nomeCor velho — ex.: desenho copiado e a
  // ribana trocada depois: a linha de tecido vira "Vermelho Ribana Moletom" mas
  // o bloco guardava "Mostarda". Só quando as contagens batem (alinhamento por
  // índice seguro) e a linha tem cor real; o nomeTecido do bloco (nome da fase)
  // é preservado.
  if (tecidos.length === blocosEnfesto.length) {
    blocosEnfesto.forEach((bl, i) => {
      const c = tecidos[i] && tecidos[i].corNome;
      if (c && c !== '—') bl.nomeCor = c;
    });
  }
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

// A OS montada pelo formulário só contém os campos DO FORMULÁRIO. Tudo o que é
// preenchido depois, na folha da OS pronta, mora em `progresso` e não tem campo
// no form: checklist de etapas e tarefas, Início/Fim de corte e de cada fase de
// enfesto, os tons por fase, as tonalidades do "Total por tamanho" e o carimbo
// etapasSeq. Trocar a OS inteira pelo objeto do formulário apagava tudo isso a
// cada edição — o checklist voltava do zero, a OS sumia dos painéis de estoque
// (sem etapasSeq não há fase atual) e o volume da OE e a contagem de etiquetas
// caíam junto com as tonalidades.
// Mesclando, o que o formulário controla é sobrescrito e o resto sobrevive.
function _mesclarComOSExistente(data) {
  const ant = (STATE.ordens || []).find(o => o.id === data.id);
  return ant ? { ...ant, ...data } : data;
}

// Trava contra o duplo clique: gravar demora (Supabase + PDF), e o segundo
// clique entrava com um `id` novo — nascia uma SEGUNDA OS com o mesmo número.
// Foi assim que a 0398 virou dois registros, criados com 14 segundos de
// diferença.
let _salvandoOS = false;

// Porta única de entrada de toda gravação de OS. Devolve os dados prontos, ou
// null quando não se deve salvar. Os três botões que gravam (Salvar OS, Salvar e
// Gerar PDF, Imprimir Etiquetas) passam por aqui, senão a regra valeria em um e
// não nos outros.
function _prepararOSParaSalvar() {
  if (!exigirEdicao('criar ou editar OS')) return null;
  const data = _mesclarComOSExistente(coletaOS());
  if (!data.os && !data.codigo) {
    toast('Preencha ao menos número da OS ou código do desenho', 'err');
    return null;
  }
  // NÚMERO CANÔNICO: sempre quatro dígitos. Sem isto, "340" e "0340" eram duas
  // OS diferentes para o programa e dois arquivos diferentes na pasta, sendo o
  // mesmo número para quem trabalha.
  if (data.os) data.os = _numeroOSCanonico(data.os);
  // Número já usado por OUTRA OS: recusa. O número é a identidade da OS no chão
  // de fábrica — é por ele que a operação, a expedição, o estoque e o PDF na
  // pasta se referem ao lote. Dois registros com o mesmo número deixam tudo isso
  // ambíguo, e o programa não tem como adivinhar de qual se trata.
  if (data.os && data.os !== 'sem-numero') {
    const conflito = (STATE.ordens || []).find(o => o.id !== data.id
      && _numeroOSCanonico(o.os) === data.os);
    if (conflito) {
      toast(`A OS ${data.os} já existe (${conflito.modeloNome || 'sem modelo'} · ${formatDate(conflito.data)}). `
        + `Use outro número — o próximo livre é ${proximoNumeroOS()}.`, 'err');
      return null;
    }
  }
  if (!validarAntesDeSalvar(data)) return null;
  return data;
}

// Roda uma gravação inteira sob a trava do duplo clique.
async function _comTravaDeSalvar(fn) {
  if (_salvandoOS) { toast('Salvando… aguarde', ''); return; }
  _salvandoOS = true;
  try { await fn(); }
  finally { _salvandoOS = false; }
}

async function salvarOS() {
  const data = _prepararOSParaSalvar();
  if (!data) return;
  await _comTravaDeSalvar(() => _salvarOSConfirmada(data));
}

async function _salvarOSConfirmada(data) {
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
  if (!exigirEdicao('conectar a pasta de PDFs')) return;
  const handle = await pickPdfFolder();
  if (handle) {
    toast(`Pasta conectada: ${handle.name}`, 'ok');
    atualizarPdfFolderStatus();
  }
}

async function desconectarPastaPdf() {
  if (!exigirEdicao('desconectar a pasta de PDFs')) return;
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

/* ----- Pasta das EXPORTAÇÕES (JSON) ----- */
// Separada da pasta de backup automático de propósito. São duas coisas com
// vidas diferentes:
//   • a de BACKUP guarda um `gerador-os-dados.json` que é REESCRITO a cada
//     alteração — é o espelho do agora, e não tem história;
//   • esta guarda as exportações MANUAIS, uma por arquivo datado. É o histórico:
//     serve para voltar a um dia específico, e por isso nada aqui é sobrescrito.
// Misturar as duas na mesma pasta funciona, mas quem for procurar "o backup de
// segunda" acaba achando o espelho de hoje.
const EXPORT_DB_KEY = 'export-folder';
let exportFolderHandle = null;

async function saveExportFolderHandle(handle) {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).put(handle, EXPORT_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function loadExportFolderHandle() {
  try {
    const db = await _openPdfDb();
    const handle = await new Promise((res, rej) => {
      const tx = db.transaction(PDF_DB_STORE, 'readonly');
      const req = tx.objectStore(PDF_DB_STORE).get(EXPORT_DB_KEY);
      req.onsuccess = () => res(req.result || null);
      req.onerror = () => rej(req.error);
    });
    db.close();
    return handle;
  } catch (e) {
    console.warn('loadExportFolderHandle', e);
    return null;
  }
}

async function clearExportFolderHandle() {
  const db = await _openPdfDb();
  await new Promise((res, rej) => {
    const tx = db.transaction(PDF_DB_STORE, 'readwrite');
    tx.objectStore(PDF_DB_STORE).delete(EXPORT_DB_KEY);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function pickExportFolder() {
  if (!('showDirectoryPicker' in window)) {
    toast('Navegador não suporta seleção de pasta. Use Chrome ou Edge no desktop.', 'err');
    return null;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await saveExportFolderHandle(handle);
    exportFolderHandle = handle;
    return handle;
  } catch (e) {
    if (e.name === 'AbortError') return null;
    console.error('pickExportFolder', e);
    toast('Falha ao selecionar pasta: ' + (e.message || e), 'err');
    return null;
  }
}

async function conectarPastaExport() {
  if (!exigirEdicao('conectar a pasta das exportações')) return;
  const handle = await pickExportFolder();
  if (handle) {
    toast(`Pasta das exportações conectada: ${handle.name}`, 'ok');
    atualizarExportFolderStatus();
  }
}

async function desconectarPastaExport() {
  if (!exigirEdicao('desconectar a pasta das exportações')) return;
  await clearExportFolderHandle();
  exportFolderHandle = null;
  toast('Pasta das exportações desconectada', '');
  atualizarExportFolderStatus();
}

async function atualizarExportFolderStatus() {
  const el = document.getElementById('exportFolderStatus');
  if (!el) return;
  if (!('showDirectoryPicker' in window)) {
    el.innerHTML = '<span style="color: var(--alert);">Este navegador não suporta a API de pasta. Use Chrome ou Edge no desktop.</span>';
    return;
  }
  const handle = exportFolderHandle || (await loadExportFolderHandle());
  if (!handle) {
    el.innerHTML = '<span style="color: var(--ink-3);">Nenhuma pasta conectada — a exportação vai para os <b>Downloads</b> do navegador.</span>';
    return;
  }
  exportFolderHandle = handle;
  let permLabel = 'pronta';
  try {
    const perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') permLabel = 'precisa renovar permissão (clique em "Conectar pasta")';
  } catch (e) { permLabel = 'permissão desconhecida'; }
  // Quantas exportações já estão lá: é o que diz se o histórico está de pé.
  let quantos = 0, ultimo = '';
  try {
    if (typeof handle.entries === 'function') {
      for await (const [n, h] of handle.entries()) {
        if (h.kind === 'file' && /^BACKUP-COMPLETO-.*\.json$/i.test(n)) {
          quantos++;
          if (n > ultimo) ultimo = n;
        }
      }
    }
  } catch (e) { /* sem iterador: mostra só o nome da pasta */ }
  el.innerHTML = `Pasta conectada: <b>${esc(handle.name)}</b> — ${esc(permLabel)}.`
    + (quantos ? ` <span class="exp-badge ok">${quantos} exportação(ões) lá</span>`
        + (ultimo ? `<div style="font-size:12px;color:var(--ink-3);margin-top:2px;">a mais recente: ${esc(ultimo)}</div>` : '')
      : ' <span class="exp-badge baixo">ainda sem exportação</span>');
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
// Idem para a EXPEDIÇÃO: se este dispositivo já viu cargas de expedição nesta
// sessão, um flush que as zera é bloqueado (protege as OEs, que NÃO entram na
// trava geral acima — ela só cobre OS+desenhos).
let _appJaTeveExpedicao = false;
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
  if (!exigirEdicao('ver os snapshots de contingência')) return;
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
  _cloudLoadErro = false; // restauramos dado bom
  cloudCache._device = DEVICE_ID; // este dispositivo é o autor da restauração
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
  _baseline = Object.assign({}, cloudCache);   // restauração é a nova base do merge
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
  if (!exigirEdicao('conectar a pasta de backup')) return;
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
  if (!exigirEdicao('desconectar a pasta de backup')) return;
  await clearBackupFolderHandle();
  backupFolderHandle = null;
  toast('Pasta de backup desconectada', '');
  atualizarBackupFolderStatus();
}

async function escreverBackupJsonAgora() {
  if (!exigirEdicao('gravar o backup na pasta')) return;
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

// Data no formato brasileiro para NOME DE ARQUIVO: DD-MM-AAAA. Usa hífen no
// lugar da barra (/ é inválida em nome de arquivo). Espelha o formatDate da
// tela, que mostra DD/MM/AAAA. ISO inesperada cai de volta como veio.
function dataBrArquivo(iso) {
  if (!iso) return '';
  const [y, m, d] = String(iso).split('-');
  return (d && m && y) ? `${d}-${m}-${y}` : String(iso);
}

// O arquivo da OS na pasta é NOMEADO SÓ PELO NÚMERO, e o número vai sempre com
// os quatro dígitos. Duas razões, as duas vindas de conflito real na pasta:
//
//   • a DATA no nome fazia a mesma OS virar vários arquivos. Regravar a folha
//     depois de mexer na OS criava um irmão em vez de substituir, e mudar a data
//     da OS criava mais um: a 0445 chegou a ter três arquivos (sem data, 23/07 e
//     28/07), com conteúdos diferentes e nenhuma pista de qual valia;
//   • sem o zero à esquerda, "340" e "0340" viravam arquivos diferentes para o
//     mesmo número.
//
// A data continua dentro da folha, que é onde ela informa. No nome ela só
// impedia o arquivo de ser substituído.
function pdfFilenameForOS(o) {
  return `OS-${_numeroOSCanonico(o && o.os)}.pdf`;
}

// Número da OS na forma canônica: quatro dígitos, sem espaço, sem nada que não
// sirva em nome de arquivo. É o que faz "340", "0340" e " 340 " serem o mesmo.
function _numeroOSCanonico(os) {
  const bruto = String(os == null ? '' : os).trim();
  if (!bruto) return 'sem-numero';
  const n = parseInt(bruto, 10);
  return isNaN(n) ? (sanitizeForFilename(bruto) || 'sem-numero') : formatarNumeroOS(n);
}

// Apaga da pasta as versões ANTIGAS do arquivo desta OS: as que trazem a data no
// nome e as que ficaram sem o zero à esquerda. É a mesma OS — manter as duas era
// justamente o conflito. Só mexe em arquivo cujo nome o próprio programa gera.
async function _limparVariantesPdfDaOS(dir, o, manter) {
  const numero = _numeroOSCanonico(o && o.os);
  if (numero === 'sem-numero' || !dir || typeof dir.entries !== 'function') return [];
  const semZero = String(parseInt(numero, 10));
  const re = new RegExp(`^OS-0*(?:${semZero})(?:-\\d{2}-\\d{2}-\\d{4})?\\.pdf$`, 'i');
  const apagados = [];
  try {
    const nomes = [];
    for await (const [n, h] of dir.entries()) {
      if (h.kind === 'file' && n !== manter && re.test(n)) nomes.push(n);
    }
    for (const n of nomes) {
      try { await dir.removeEntry(n); apagados.push(n); } catch (e) { /* segue */ }
    }
  } catch (e) { /* pasta sem iterador: não dá para limpar, e tudo bem */ }
  return apagados;
}

// Mesmo número canônico da folha: sem ele, "340" e "0340" geravam
// `etiqueta-340.pdf` e `etiqueta-0340.pdf` para o mesmo número.
function etiquetaFilenameForOS(o) {
  return `etiqueta-${_numeroOSCanonico(o && o.os)}.pdf`;
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
    // Com mais de uma tonalidade, o tom vai JUNTO do tamanho no destaque —
    // "G tom 1", "G tom 2"… — senão duas etiquetas do mesmo tamanho ficam
    // indistinguíveis no ensaque. Sem linha "TOM:" separada: seria repetição.
    const tomDoPacote = (dados.tonsPacotes || [])[i];
    const tomSuf = (!ehReposicao && (dados.nTons || 1) > 1 && tomDoPacote != null) ? ` tom ${tomDoPacote}` : '';
    const destaque = ehReposicao
      ? { t: ETIQUETA_CONTEUDO_REPOSICAO, s: 1.6, c: true }   // conteúdo (texto longo)
      : { t: (tams[i] || dados.tam) + tomSuf, s: 2, c: true };// tamanho (+ tom) do pacote, dobro
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
    // FALAR quando falha. Antes o erro morria no console e o retorno era
    // descartado por quem chama: a OS era salva, a folha ia pra pasta, e a
    // etiqueta continuava a de dias atrás sem ninguém ficar sabendo — foi o que
    // aconteceu com a 0436, salva em 30/07 com a etiqueta de 22/07 na pasta.
    // Etiqueta velha é pacote ensacado errado, então isto é aviso de erro.
    if (!silent) toast(`A etiqueta da OS ${_numeroOSCanonico(o && o.os)} NÃO foi regravada: ${e && e.message || e}`, 'err');
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
  // MESMO encaixe da impressao: a folha e alargada ate ficar na proporcao da A4,
  // com 15mm de margem esquerda e 8mm nas outras tres. Aqui nao ha zoom — a foto
  // ja sai na proporcao certa, e o addImage abaixo a estica na folha inteira.
  // E o que faz o arquivo da pasta e o papel saírem iguais.
  encaixarFolhaNaA4(sheet);
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
    desfazerEncaixeA4(sheet);
    sheet.style.zoom = prevZoom;
    sheet.style.transform = prevTransform;
    sheet.style.transformOrigin = prevOrigin;
  }
}

// `os` é opcional: quando vem, as versões antigas do mesmo número são
// consolidadas depois de gravar.
async function savePdfToFolder(blob, filename, os) {
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
    if (os) await _limparVariantesPdfDaOS(handle, os, filename);
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
// Instante em que a gravação em curso começou. Se uma captura travar (html2canvas
// já travou em ambiente sem layout), a trava de reentrada ficaria ligada para
// sempre e nenhuma OE seria salva pelo resto da sessão — sem nenhum aviso.
let _oeSalvandoDesde = 0;

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
  if (!exigirEdicao('conectar a pasta das OE')) return;
  const handle = await pickOeFolder();
  if (handle) {
    toast(`Pasta das OE conectada: ${handle.name}`, 'ok');
    atualizarOeFolderStatus();
  }
}
async function desconectarPastaOe() {
  if (!exigirEdicao('desconectar a pasta das OE')) return;
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
  const iniBr = dataBrArquivo(ini), fimBr = dataBrArquivo(fim);
  const base = (ini === fim) ? `OE-${iniBr}` : `OE-${iniBr}_a_${fimBr}`;
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
    const _bottomsDe = sel => Array.from(sheet.querySelectorAll(sel))
      .map(el => (el.getBoundingClientRect().bottom - sheetRect.top) * ratio)
      .filter(v => v > 0 && v <= canvas.height)
      .sort((a, b) => a - b);
    const cortesSeguros = _bottomsDe('.exp-print-bloco');
    // Rede de segurança para o dia que sozinho é mais alto que uma página: em
    // vez do corte duro no meio da folha — que partia um quadro de OS ao meio —
    // o corte cai no fim de um QUADRO. O dia continua na folha seguinte, mas
    // nenhuma OS sai picada.
    const cortesCaixa = _bottomsDe('.exp-print-os');
    // CABEÇALHO FIXO: a faixa do topo (título, período, unidades, emissão) é
    // redesenhada no alto de cada página. Sem isso, da segunda folha em diante o
    // papel chegava à mão de quem confere sem dizer que OE é aquela nem de que
    // período — e uma OE mensal tem várias folhas.
    const headEl = sheet.querySelector('.exp-print-head');
    const cabAlturaPx = headEl
      ? Math.round((headEl.getBoundingClientRect().bottom - sheetRect.top) * ratio)
      : 0;
    const pdf = new _jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
    const pageWmm = 210, pageHmm = 297;
    const pxPorMm = canvas.width / pageWmm;      // px do canvas por mm de largura A4
    const pageHpx = Math.floor(pageHmm * pxPorMm); // px que cabem numa A4 de altura
    // RECUO do 1º quadro nas folhas repetidas: getBoundingClientRect NÃO inclui o
    // margin-bottom do cabeçalho, então nas páginas 2+ o conteúdo era colocado
    // colado em cabAlturaPx e o primeiro quadro saía cortado pela sobreposição do
    // cabeçalho fixo. Reserva o mesmo respiro do fluxo (margin do head) + 3mm.
    const headMb = headEl ? (parseFloat(getComputedStyle(headEl).marginBottom) || 0) : 0;
    const cabGapPx = Math.round(headMb * ratio) + Math.round(3 * pxPorMm);
    let y = 0, pagina = 0;
    while (y < canvas.height - 1) {
      // Na 1ª página o cabeçalho já vem no próprio fluxo; nas demais ele é
      // repetido e consome altura útil da folha.
      const repetirCab = pagina > 0 && cabAlturaPx > 0;
      const alturaCab = repetirCab ? cabAlturaPx + cabGapPx : 0;
      const disponivel = pageHpx - alturaCab;
      const maxY = y + disponivel;
      let cut;
      if (maxY >= canvas.height) {
        cut = canvas.height;
      } else {
        // maior fim-de-bloco que cabe inteiro nesta página
        const cand = cortesSeguros.filter(v => v > y + 1 && v <= maxY);
        if (cand.length) cut = Math.max(...cand);
        else {
          const candCaixa = cortesCaixa.filter(v => v > y + 1 && v <= maxY);
          cut = candCaixa.length ? Math.max(...candCaixa) : maxY; // só então, corte duro
        }
      }
      const sliceH = Math.max(1, Math.round(cut - y));
      const tmp = document.createElement('canvas');
      tmp.width = canvas.width; tmp.height = alturaCab + sliceH;
      const ctx = tmp.getContext('2d');
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, tmp.width, tmp.height);
      if (repetirCab) {
        ctx.drawImage(canvas, 0, 0, canvas.width, cabAlturaPx, 0, 0, canvas.width, cabAlturaPx);
      }
      ctx.drawImage(canvas, 0, y, canvas.width, sliceH, 0, alturaCab, canvas.width, sliceH);
      const imgData = tmp.toDataURL('image/jpeg', 0.95);
      if (pagina > 0) pdf.addPage();
      pdf.addImage(imgData, 'JPEG', 0, 0, pageWmm, tmp.height / pxPorMm, undefined, 'FAST');
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
// A folha da OE vive dentro da section .page, que fica display:none quando o
// usuário está em outra tela. html2canvas mede o elemento no layout: escondido,
// o retângulo é zero e o PDF sairia em branco. Para poder salvar de QUALQUER
// tela, a section é revelada fora da vista só durante a captura e devolvida ao
// estado anterior no fim — o usuário não vê nada piscar.
async function _comFolhaOeRenderizavel(fn) {
  const sec = document.querySelector('section.page[data-page="print-expedicao"]');
  if (!sec || !sec.classList.contains('hidden')) return await fn();
  const styleAntes = sec.getAttribute('style');
  sec.classList.remove('hidden');
  sec.style.position = 'fixed';
  sec.style.left = '-10000px';
  sec.style.top = '0';
  sec.style.width = '260mm';     // folga sobre os 210mm da folha
  sec.style.zIndex = '-1';
  sec.style.pointerEvents = 'none';
  try {
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    return await fn();
  } finally {
    if (styleAntes == null) sec.removeAttribute('style'); else sec.setAttribute('style', styleAntes);
    sec.classList.add('hidden');
  }
}

// Regrava o PDF da OE depois de mexer no plano de expedição — o mesmo hábito da
// OS, que regrava folha e etiqueta ao ser salva. Antes o arquivo em disco só era
// atualizado quando alguém ABRIA a folha do plano: alocar uma OS e não abrir a
// folha deixava a OE do dia desatualizada na pasta.
// Debounce porque montar uma carga são vários cliques seguidos, e cada captura
// html2canvas é cara — interessa o estado final, não cada passo.
let _oeAutoTimer = null;
// Há um auto-save da OE ADIADO porque o usuário está vendo a folha? Grava ao
// sair da folha (aí ela está escondida e a captura é fora da tela, sem
// "subir e descer" na frente do usuário).
let _oeSaveAdiado = false;
function agendarAutoSaveOE() {
  if (_oeAutoTimer) clearTimeout(_oeAutoTimer);
  _oeAutoTimer = setTimeout(() => {
    _oeAutoTimer = null;
    salvarPdfOeNaPasta({ silent: true }).catch(e => console.warn('auto-save OE', e));
  }, 1500);
}

// O auto-save roda de um timer, sem clique do usuário. Nessa situação o
// navegador NÃO permite abrir a caixa de permissão da pasta (requestPermission
// exige gesto do usuário), então a gravação falhava calada e parecia que o
// recurso simplesmente não funcionava. Aqui o motivo é dito uma vez por sessão,
// com a instrução do que fazer — e sem repetir a cada gravação.
let _oeAvisoDado = '';
function _avisarOeUmaVez(chave, msg) {
  if (_oeAvisoDado === chave) return;
  _oeAvisoDado = chave;
  console.warn('auto-save OE:', msg);
  toast(msg, 'err');
}

// Há OE de verdade no período atual? = existe alguma ocorrência não cancelada
// com pelo menos uma OS alocada (ida ou volta). Espelha o filtro de
// renderPrintPlanoExpedicao: sem isso a folha sai só com cabeçalho e o aviso
// "Nenhuma OE produzida". Fonte única pra tela e gravação não divergirem.
function oeTemConteudo() {
  const { ini, fim } = _expRange(expPlanoModo, expPlanoAncora);
  return ocorrenciasExpedicao(ini, fim).some(oc =>
    !oc.cancelada &&
    resumoPernaExpedicao(oc, 'ida').itens.length +
    resumoPernaExpedicao(oc, 'volta').itens.length > 0);
}

async function salvarPdfOeNaPasta({ silent = false } = {}) {
  // Trava de reentrada com validade: uma captura travada não pode calar o
  // auto-save para sempre.
  if (_oeSalvando && (Date.now() - _oeSalvandoDesde) < 60000) return false;
  // Auto-save (silencioso) grava SÓ a OE DIÁRIA. Semanal e mensal não são
  // salvas sozinhas — evita encher a pasta com PDFs de período que o chão não
  // usa (a expedição imprime o diário). O salvar MANUAL (botão, silent=false)
  // continua valendo para qualquer modo.
  if (silent && expPlanoModo !== 'dia') return false;
  // NÃO perturbar a folha enquanto o usuário a está VENDO. O auto-save
  // re-renderiza a folha e liga o pdf-capture (que desloca o layout), fazendo
  // ela "subir e descer". Se a folha está visível, adia: grava quando o usuário
  // sair dela (aí está escondida → captura fora da tela). O salvar MANUAL
  // (silent=false, botão) segue normal — o usuário pediu.
  if (silent) {
    const secOe = document.querySelector('section.page[data-page="print-expedicao"]');
    if (secOe && !secOe.classList.contains('hidden')) { _oeSaveAdiado = true; return false; }
  }
  // Não grava OE VAZIA na pasta como se fosse OE emitida: sem nenhuma OS
  // alocada no período, o PDF sairia só com cabeçalho e "Nenhuma OE produzida".
  if (!oeTemConteudo()) {
    if (silent) return false;
    toast('OE vazia — nenhuma OS alocada neste período. Nada foi salvo.', 'err');
    return false;
  }
  let handle = oeFolderHandle || (await loadOeFolderHandle());
  if (!handle) {
    if (silent) {
      _avisarOeUmaVez('sem-pasta', 'A OE não está sendo salva sozinha: nenhuma pasta conectada. Configurações → Pasta das OE.');
      return false;
    }
    toast('Conecte a pasta das OE em Configurações primeiro.', 'err');
    return false;
  }
  // No modo silencioso só CONSULTA a permissão: pedir exigiria um gesto do
  // usuário que um timer não tem, e a chamada falharia de qualquer jeito.
  let permOk;
  if (silent) {
    try { permOk = (await handle.queryPermission({ mode: 'readwrite' })) === 'granted'; }
    catch (e) { permOk = false; }
    if (!permOk) {
      _avisarOeUmaVez('sem-permissao',
        'A OE não pôde ser salva sozinha: a pasta precisa de permissão renovada. Abra a folha do plano e clique em "Salvar OE na pasta" uma vez.');
      return false;
    }
  } else {
    permOk = await ensureFolderPermission(handle, 'readwrite');
    if (!permOk) { toast('Permissão da pasta das OE negada', 'err'); return false; }
    _oeAvisoDado = '';   // permissão renovada: volta a avisar se falhar de novo
  }
  oeFolderHandle = handle;
  _oeSalvando = true;
  _oeSalvandoDesde = Date.now();
  try {
    if (!silent) toast('Gerando PDF da OE...', '');
    const blob = await _comFolhaOeRenderizavel(async () => {
      renderPrintPlanoExpedicao();
      await new Promise(r => setTimeout(r, 150));
      return await gerarPdfDaSheetExp();
    });
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
  const data = _prepararOSParaSalvar();
  if (!data) return;
  if (_salvandoOS) return toast('Salvando… aguarde', '');
  _salvandoOS = true;
  try { await _salvarEImprimirConfirmada(data); }
  finally { _salvandoOS = false; }
}

async function _salvarEImprimirConfirmada(data) {
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
    const saved = await savePdfToFolder(blob, filename, data);
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
          const okC = await savePdfToFolder(blobC, fnC, conjugada);
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

  // Uma etiqueta por PACOTE — a MESMA regra do volume de expedição:
  // tamanhos × TONALIDADES + 1 (reposição/ribana). Cada tonalidade é ensacada
  // separada, então cada tamanho rende um pacote por tom.
  // Antes o cálculo parava em "tamanhos + 1" e ignorava a tonalidade: uma OS de
  // 7 tamanhos em 2 tons saía com 8 etiquetas para 15 pacotes reais — 7 pacotes
  // iam para a expedição sem etiqueta nenhuma.
  // Ex.: P-G1-G2 em 1 tom = 4 etiquetas; os mesmos 3 tamanhos em 2 tons = 7.
  // Mínimo 1 pra não bloquear OS sem grade.
  const tamanhosBase = _tamanhosDaGradeExpandido(o);
  const tonsAtivos = tonsEfetivos((o.progresso || {}).totalTamanhoTons || {});
  const nTons = Math.max(1, tonsAtivos.length);
  const tamanhosPacotes = [];
  const tonsPacotes = [];
  for (let ti = 0; ti < nTons; ti++) {
    tamanhosBase.forEach(t => {
      tamanhosPacotes.push(t);
      tonsPacotes.push(tonsAtivos[ti] != null ? tonsAtivos[ti] : null);
    });
  }
  const temReposicao = tamanhosPacotes.length > 0;
  const numEtiquetas = temReposicao ? tamanhosPacotes.length + 1 : 1;

  // Moletom: cada etiqueta de pacote (de tamanho) recebe a lista de composição.
  const composicao = temMoletom ? ETIQUETA_COMPOSICAO_MOLETOM : null;

  return { marca, os, qtde, tam, cor, modelo: desenhoNome, numEtiquetas,
           tamanhosPacotes, tonsPacotes, nTons, temReposicao, composicao };
}

// Abre as etiquetas em PDF numa aba, prontas para imprimir. É o caminho EXATO:
// o PDF leva a página de 100x50 mm dentro do arquivo (jsPDF, format [100,50]),
// então a impressora recebe a medida e não precisa adivinhar.
//
// A janela HTML de etiquetas depende do `@page { size: 100mm 50mm }`, e o Chrome
// só o respeita quando a impressora escolhida TEM esse papel. Numa laser A4
// comum ele cai no papel da impressora e a etiqueta sai pequena, no canto de uma
// folha inteira — foi o que aconteceu na OS 0452. Pelo PDF isso não acontece.
function imprimirEtiquetasPdf(osId) {
  const o = STATE.ordens.find(x => x.id === osId);
  if (!o) { toast('OS não encontrada', 'err'); return; }
  try {
    // gerarPdfEtiquetas devolve o BLOB pronto (termina em `pdf.output('blob')`),
    // não o documento jsPDF. Aqui se chamava .output('blob') no que já era um
    // Blob — "pdf.output is not a function" —, e o botão de etiquetas da OS
    // morria nesse erro em TODA OS desde 28/07 (commit 420f7f9). Quem grava na
    // pasta (salvarPdfEtiquetasAuto) sempre tratou o retorno como Blob; era só
    // este ponto fora de compasso.
    const blob = gerarPdfEtiquetas(dadosEtiquetaParaOS(o));
    const url = URL.createObjectURL(blob);
    const w = window.open(url, '_blank');
    if (!w) { toast('Permita pop-ups para abrir o PDF das etiquetas', 'err'); return; }
    // O objeto só é liberado depois de a aba carregar; revogar na hora deixaria
    // a aba em branco.
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    toast('PDF das etiquetas aberto (100 × 50 mm). No diálogo, deixe Margens: Nenhuma e Escala: 100%.', 'ok');
  } catch (e) {
    console.error('imprimirEtiquetasPdf', e);
    toast('Falha ao gerar o PDF das etiquetas: ' + (e.message || e), 'err');
  }
}

function imprimirEtiquetas(osId) {
  const o = STATE.ordens.find(x => x.id === osId);
  if (!o) { toast('OS não encontrada', 'err'); return; }

  // Guarda o objeto INTEIRO, e não só os campos usados na janela: é ele que vai
  // para o auto-save lá embaixo. Montar um objeto novo à mão ali era o que
  // gravava etiqueta capenga por cima da boa (ver o comentário no auto-save).
  const dados = dadosEtiquetaParaOS(o);
  const { marca, os, qtde, tam, cor, modelo: desenhoNome, numEtiquetas,
          tamanhosPacotes, tonsPacotes, nTons, temReposicao, composicao } = dados;

  const escEt = s => String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  // Cada etiqueta é um pacote: as de tamanho mostram o SEU tamanho em destaque
  // (fonte dobrada); a última é o pacote de reposição, com o conteúdo. Moletom:
  // as etiquetas de tamanho recebem a lista de composição do pacote.
  const compHtml = composicao ? composicao.map(c => `<div class="comp">${escEt(c)}</div>`).join('') : '';
  const corpo = Array.from({ length: numEtiquetas }, (_, i) => {
    const ehRep = temReposicao && i === numEtiquetas - 1;
    // Tonalidade junto do tamanho no destaque — "G tom 1", "G tom 2"… — só quando
    // a OS tem mais de um tom; assim duas etiquetas do mesmo tamanho não ficam
    // indistinguíveis no ensaque. A linha "TOM:" separada some: viraria repetição.
    const tomSuf = (nTons > 1 && !ehRep && tonsPacotes[i] != null) ? ` tom ${tonsPacotes[i]}` : '';
    const destaque = ehRep
      ? `<div class="big rep">${escEt(ETIQUETA_CONTEUDO_REPOSICAO)}</div>`
      : `<div class="big">${escEt(((tamanhosPacotes && tamanhosPacotes[i]) || tam) + tomSuf)}</div>${compHtml}`;
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
    <button onclick="window.opener && window.opener.imprimirEtiquetasPdf && window.opener.imprimirEtiquetasPdf('${escEt(osId)}')">📄 Abrir PDF 10×5cm</button>
    <span style="margin-left:12px;color:#555;font-size:13px;">${numEtiquetas} etiqueta${numEtiquetas>1?'s':''} · LOTE 1${numEtiquetas>1?'..'+numEtiquetas:''} · 10×5cm</span>
    <div style="margin-top:8px;font-size:12px;color:#555;line-height:1.45;max-width:760px;">
      Esta janela imprime pelo navegador, e aí o <b>tamanho da folha é o da impressora</b>: numa laser A4 a etiqueta sai
      pequena no canto da folha. Para a <b>impressora de etiquetas 10×5cm</b>, use o <b>PDF</b> — ele carrega a página de
      100&times;50&nbsp;mm dentro do arquivo, e a impressora recebe a medida exata.
      No diálogo do Chrome, deixe <b>Margens: Nenhuma</b> e <b>Escala: 100%</b> (nunca "Ajustar à página").
    </div>
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
  //
  // `dados` INTEIRO. Aqui se montava um objeto novo com sete campos, sem
  // tamanhosPacotes/tonsPacotes/nTons/temReposicao/composicao — e o
  // gerarPdfEtiquetas, sem esses campos, cai nos defaults: o destaque de toda
  // etiqueta vira a lista inteira de tamanhos ("P-M-G-GG-G1-G2-G3") em vez do
  // tamanho daquele pacote, a última deixa de ser a de reposição e a composição
  // do moletom some. Como a janela "etiquetas (tela)" grava por cima do arquivo
  // da pasta, conferir na tela ESTRAGAVA a etiqueta boa que já estava salva.
  // Aconteceu com a 0435 (28/07) e a 0452 (30/07), medido pelo tamanho do PDF.
  salvarPdfEtiquetasAuto(o, dados);
}

function imprimirEtiquetasAtual() {
  if (!printOsAtual) { toast('Abra uma OS antes', 'err'); return; }
  imprimirEtiquetasPdf(printOsAtual.id);
}

async function salvarEImprimirEtiquetas() {
  const data = _prepararOSParaSalvar();
  if (!data) return;
  await _comTravaDeSalvar(() => _salvarEtiquetasConfirmada(data));
}

async function _salvarEtiquetasConfirmada(data) {
  const idx = STATE.ordens.findIndex(o => o.id === data.id);
  if (idx >= 0) STATE.ordens[idx] = data; else STATE.ordens.push(data);
  await saveState('ordens');
  await atualizarCounterOS(data.os);
  osEditId = null;
  imprimirEtiquetasPdf(data.id);
}

// Usado so pelo botao "Imprimir / Salvar PDF" (window.print()). O caminho
// principal — "Salvar e Gerar PDF" — nao passa por aqui: la o jsPDF ja
// encaixa a foto na A4 sozinho.
//
const PX_POR_MM = 3.7795275591;
const A4_L_MM = 210, A4_A_MM = 297;

// As margens da folha A4 moram no styles.css (:root), num lugar so, para o CSS
// da folha do plano e este calculo aqui nao saírem de sincronia.
function _margensA4mm() {
  const cs = getComputedStyle(document.documentElement);
  const num = (nome, padrao) => {
    const n = parseFloat(cs.getPropertyValue(nome));
    return Number.isFinite(n) ? n : padrao;
  };
  return { esq: num('--print-esq', 15), outras: num('--print-margem', 8) };
}

// ENCAIXA A FOLHA DE OS NA A4 INTEIRA.
//
// A folha e desenhada em 210mm de largura e cresce em altura conforme o
// conteudo. Quando passa de 297mm ela tem que encolher para caber em uma folha
// — e encolher reduz a LARGURA junto, deixando uma tira de papel em branco na
// direita e no pe. Era o que se via na impressao: 183mm de conteudo numa folha
// de 210mm.
//
// A saida e alargar o bloco na mesma medida em que ele vai encolher: o conteudo
// reflui mais largo, a altura cai, e a proporcao caminha para a da A4
// (297/210). Quando chega la, encolher para caber na altura devolve exatamente
// 210mm de largura — folha cheia. Como alargar muda a altura, o calculo e
// repetido ate parar de andar (poucas voltas; a folha e quase toda tabela).
//
// As margens sao aplicadas na mesma escala, para caírem nos milimetros pedidos
// depois do encolhimento: 15mm na esquerda, 8mm nas outras tres.
//
// Retorna a escala que a impressao deve aplicar (o PDF da pasta nao usa: la o
// jsPDF encaixa a foto sozinho, e a foto ja sai na proporcao certa).
function encaixarFolhaNaA4(sheet) {
  const m = _margensA4mm();
  const razaoA4 = A4_A_MM / A4_L_MM;

  // Aplica uma largura e devolve o que a folha OCUPA com ela. A medida vem do
  // scrollWidth, e nao da largura pedida: quando o conteudo tem um minimo que
  // nao cabe na caixa (a Camiseta Tricolor tem 5 fases de enfesto e a coluna do
  // CONSUMO nao comprime), ele transborda, e e o transbordo que decide se sai
  // cortado no papel.
  const aplicar = (larguraMm) => {
    const k = larguraMm / A4_L_MM; // o desenho todo vive nesta escala
    sheet.style.width = larguraMm.toFixed(2) + 'mm';
    sheet.style.paddingTop = (m.outras * k).toFixed(2) + 'mm';
    sheet.style.paddingRight = (m.outras * k).toFixed(2) + 'mm';
    sheet.style.paddingBottom = (m.outras * k).toFixed(2) + 'mm';
    sheet.style.paddingLeft = (m.esq * k).toFixed(2) + 'mm';
    void sheet.offsetHeight; // forca reflow pra leitura correta
    return {
      largura: Math.max(sheet.scrollWidth, sheet.offsetWidth) / PX_POR_MM,
      altura: sheet.scrollHeight / PX_POR_MM
    };
  };

  // `medida` descreve SEMPRE o que esta no DOM neste instante — e o ponto em
  // que isto errava. Antes a largura era avancada para o proximo palpite e a
  // escala saia desse palpite, enquanto a folha no DOM continuava na largura
  // anterior. Quando as voltas nao assentavam (e nas tricolores nao assentam,
  // porque a foto e as 5 linhas de enfesto fazem a altura pular), a folha ficava
  // MAIS LARGA do que a escala supunha e o conteudo saia cortado na direita:
  // 225,4mm de folha num papel de 210mm na Camiseta Tricolor — o "ADULTO
  // UNISEX" perdia o X.
  // A escala de uma medida: o quanto a folha pode crescer sem passar da borda,
  // em largura OU em altura — manda a mais apertada das duas.
  const escalaDe = med => Math.min(A4_L_MM / med.largura, A4_A_MM / med.altura);
  // O que se quer e a folha OCUPADA, entao a nota de cada tentativa e a area
  // que ela desenha no papel. Notar pela escala seria outra coisa: escolheria a
  // letra maior, e a letra maior pode vir numa folha estreita, com uma tira de
  // papel sobrando na direita — que e justamente o que se esta tirando daqui.
  const areaDe = med => {
    const e = escalaDe(med);
    return (med.largura * e) * (med.altura * e);
  };

  let larguraMm = A4_L_MM;
  let medida = aplicar(larguraMm);
  let melhor = { largura: larguraMm, area: areaDe(medida) };
  for (let i = 0; i < 8; i++) {
    const alvo = Math.min(Math.max(medida.altura / razaoA4, A4_L_MM), A4_L_MM * 2);
    if (Math.abs(alvo - larguraMm) < 0.3) break;
    // AMORTECIDO: meio caminho entre o palpite e o alvo. Nas tricolores a
    // altura anda em degrau (a foto e as 5 linhas de enfesto nao refluem de
    // pouquinho em pouquinho), e ir direto ao alvo faz a busca oscilar de um
    // extremo ao outro sem nunca assentar. A media assenta.
    larguraMm = (larguraMm + alvo) / 2;
    medida = aplicar(larguraMm);
    const a = areaDe(medida);
    if (a > melhor.area) melhor = { largura: larguraMm, area: a };
  }
  // Guarda-chuva das oscilacoes: vale a MELHOR largura que a busca viu, e nao
  // a ultima que ela por acaso tentou. Reaplicada e remedida, para a escala
  // devolvida descrever exatamente o que ficou no DOM.
  if (Math.abs(melhor.largura - larguraMm) > 0.3) medida = aplicar(melhor.largura);
  // A escala sai do que a folha ocupa DE VERDADE, largura e altura. Assim nada
  // passa da borda, tenha o encaixe assentado ou nao: no pior caso sobra uma
  // tira em branco, que e melhor do que texto cortado ou uma segunda folha.
  return escalaDe(medida);
}

function desfazerEncaixeA4(sheet) {
  if (!sheet) return;
  ['width', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft']
    .forEach(p => sheet.style.removeProperty(p.replace(/[A-Z]/g, c => '-' + c.toLowerCase())));
}

function ajustarImpressaoParaA4() {
  const sheet = document.querySelector('.sheet');
  const scaler = document.querySelector('.sheet-scaler');
  if (!sheet || !scaler) return;

  // Mede a folha na geometria de saida: .pdf-capture zera a ampliacao de
  // leitura, entao scrollHeight vem em mm reais de A4.
  scaler.style.removeProperty('zoom');
  document.body.classList.add('pdf-capture');

  const escala = encaixarFolhaNaA4(sheet);

  document.body.classList.remove('pdf-capture');

  // zoom afeta LAYOUT, entao tabelas, fontes e quebras encolhem juntas.
  scaler.style.setProperty('zoom', escala.toFixed(4), 'important');
}

window.addEventListener('beforeprint', ajustarImpressaoParaA4);
window.addEventListener('afterprint', function() {
  const scaler = document.querySelector('.sheet-scaler');
  desfazerEncaixeA4(document.querySelector('.sheet'));
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
        <button class="edit" onclick="imprimirEtiquetasPdf('${o.id}')" title="Abre as etiquetas em PDF com a página de 100 × 50 mm — é a medida exata que a impressora de etiquetas espera.">etiquetas</button>
        <button class="edit" onclick="imprimirEtiquetas('${o.id}')" title="Abre as etiquetas numa janela do navegador. O tamanho da folha passa a ser o da impressora — use só quando quiser conferir na tela.">etiquetas (tela)</button>
        <button class="edit admin-only" onclick="editarOS('${o.id}')">editar</button>
        <button class="edit admin-only" onclick="duplicarOS('${o.id}')">duplicar</button>
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
  if (!exigirEdicao('marcar etapas da OS')) return;
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
  // Marcar "Ensaque" diz que o lote está PRONTO para expedir — só isso. Entrar
  // numa OE é ato do planejamento da expedição, feito pelo usuário.
  try { await sincronizarPlanoExpedicaoDaOS(os, etapaNome, checked); }
  catch (e) { console.warn('sincronizarPlanoExpedicaoDaOS', e); }
}

async function togglarChecklistTarefa(osId, etapaNome, tarefaNome, checked) {
  if (!exigirEdicao('marcar tarefas da OS')) return;
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
  if (!exigirEdicao('marcar o enfesto da OS')) return;
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
  if (!exigirEdicao('lançar o horário de enfesto')) return;
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.enfestosTempos = os.progresso.enfestosTempos || {};
  os.progresso.enfestosTempos[ordem] = os.progresso.enfestosTempos[ordem] || {};
  const v = (valor || '').trim();
  if (v) os.progresso.enfestosTempos[ordem][campo] = v;
  else delete os.progresso.enfestosTempos[ordem][campo];
  // Confere na hora contra a média das OS de referência: o alerta aparece ao sair
  // do campo, não só na próxima vez que a folha for aberta.
  _atualizarSelosTempoEnfesto(osId);
  try { await saveState('ordens'); } catch (e) { console.warn('salvarTempoEnfesto', e); }
}

// Salva o valor digitado (à mão) de cada TOM em cada fase de enfesto na folha.
// As tonalidades podem variar em qualquer fase, então cada fase tem seus campos
// de Tom 1/2/3. Na fase PRINCIPAL esses campos são as CAMADAS REAIS por tom: ao
// digitá-los, camadas/peças-alvo e o "Total por tamanho" são recalculados (ver
// recalcularDeCamadasPorTom). Nas demais fases é anotação livre.
async function salvarTomEnfesto(osId, ordem, tom, valor) {
  if (!exigirEdicao('lançar as camadas por tonalidade')) return;
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.enfestosTons = os.progresso.enfestosTons || {};
  os.progresso.enfestosTons[ordem] = os.progresso.enfestosTons[ordem] || {};
  const v = (valor || '').trim();
  if (v) os.progresso.enfestosTons[ordem][tom] = v;
  else delete os.progresso.enfestosTons[ordem][tom];
  // Editou a linha de tons da fase PRINCIPAL → recalcula tudo a partir das
  // camadas reais por tom (a função salva e re-renderiza). Fases secundárias
  // seguem como anotação livre.
  if (String(ordem) === String(_ordemFasePrincipal(os))) {
    await recalcularDeCamadasPorTom(osId);
  } else {
    try { await saveState('ordens'); } catch (e) { console.warn('salvarTomEnfesto', e); }
  }
}

// Camadas reais por tom de uma fase (parse numérico de enfestosTons[ord]).
function _camadasPorTomFase(o, ord) {
  const tv = ((o.progresso || {}).enfestosTons || {})[ord] || {};
  const out = {};
  [1, 2, 3, 4].forEach(t => { const n = parseInt(tv[t], 10); if (n > 0) out[t] = n; });
  return out;
}

// Ordem da fase PRINCIPAL do enfesto (primeira não-viés). É a fase cujas camadas
// reais por tom mandam nas camadas/peças-alvo e no "Total por tamanho".
function _ordemFasePrincipal(o) {
  const cons = consumoEnfestoOS(o);
  const p = cons.find(L => !L.ehVies) || cons[0];
  return p ? p.ordem : null;
}

// Recalcula camadas, peças-alvo e o "Total por tamanho" a partir das CAMADAS
// REAIS por tom digitadas na fase principal. As demais fases escalam
// proporcionalmente à nova camada principal. O último tom continua sendo o
// balanceador — e, por construção, cai exatamente em camadas_último × mult.
async function recalcularDeCamadasPorTom(osId) {
  const o = STATE.ordens.find(x => x.id === osId);
  if (!o) return;
  const ordP = _ordemFasePrincipal(o);
  if (ordP == null) return;
  const porTom = _camadasPorTomFase(o, ordP);
  const tomsEntrados = Object.keys(porTom).map(Number).sort((a, b) => a - b);
  if (!tomsEntrados.length) { // nada digitado: mantém o planejado (balanceador)
    try { await saveState('ordens'); } catch (e) {}
    return;
  }
  const valores = tomsEntrados.map(t => porTom[t]);   // camadas por tom, em ordem
  const camadas = valores.reduce((s, v) => s + v, 0);
  if (!(camadas > 0)) return;
  const mult = multiplicadorPecaOS(o);
  const g = o.grade || {};
  const qtds = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].map(k => parseInt(g[k], 10) || 0).filter(q => q > 0);
  const minQtd = qtds.length ? Math.min(...qtds) : 1;

  o.enfesto = o.enfesto || {};
  // Escala as demais fases proporcionalmente à nova camada principal.
  if (Array.isArray(o.enfesto.blocos) && o.enfesto.blocos.length) {
    const bp = o.enfesto.blocos.find(b => (b.ordem || 0) === ordP);
    const antes = (bp && parseInt(bp.camadas, 10) > 0) ? parseInt(bp.camadas, 10)
      : (parseInt(o.enfesto.camadas, 10) || 0);
    const nomeFaseDe = ord => ((o.fases || []).find(f => f.ordem === ord) || {}).nome || '';
    o.enfesto.blocos.forEach(b => {
      const ehVies = /vi[eé]s/i.test(nomeFaseDe(b.ordem) || b.nomeTecido || '');
      if (ehVies) { b.camadas = 1; return; }
      if ((b.ordem || 0) === ordP) { b.camadas = camadas; return; }
      const cur = parseInt(b.camadas, 10) || 0;
      b.camadas = (antes > 0 && cur > 0) ? Math.max(1, Math.round(cur * camadas / antes)) : camadas;
    });
  }
  o.enfesto.camadas = camadas;
  o.enfesto.target = camadas * minQtd * mult;   // peças-alvo (tamanho limitante)
  // As camadas mudaram, então a QUANTIDADE da OS mudou junto: os componentes e o
  // total de peças precisam acompanhar. Sem esta linha eles continuavam com o
  // número do momento em que a OS foi salva — numa OS criada sem camadas, com
  // ZERO —, e a expedição não oferecia a OS para alocar porque a via sem peças.
  recomputarQuantidadesOS(o);

  // "Total por tamanho": N tons contíguos; cada tom editável recebe V =
  // camadas_tom × grade × mult — o V é "peças por tamanho", então precisa
  // multiplicar pela unidade da GRADE (minQtd, ex.: 2 na grade "2X") e pelo
  // multiplicador do tecido (moletom 1 / malha 2). Sem o minQtd, uma grade de 2
  // por tamanho saía com metade no tom (o colTotal e o target já incluíam a
  // grade; só o V dos tons ficava de fora). O último tom é o balanceador.
  o.progresso = o.progresso || {};
  o.progresso.totalTamanhoTons = {};
  o.progresso.totalTamanhoTomValor = o.progresso.totalTamanhoTomValor || {};
  const N = valores.length;
  for (let slot = 1; slot <= 4; slot++) {
    if (slot <= N) o.progresso.totalTamanhoTons[slot] = true;
    if (slot < N) o.progresso.totalTamanhoTomValor[slot] = valores[slot - 1] * minQtd * mult; // editável
    else delete o.progresso.totalTamanhoTomValor[slot];                                        // balanceador/inexistente
  }

  try { await saveState('ordens'); } catch (e) { console.warn('recalcularDeCamadasPorTom', e); }
  if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(o);
}

// Salva o tempo de Início/Fim do corte, mostrado junto da etapa "Corte" em
// Etapas de Produção. Um par único por OS. campo ∈ {ini, fim}.
async function salvarTempoCorte(osId, campo, valor) {
  if (!exigirEdicao('lançar o horário do corte')) return;
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
  if (!exigirEdicao('editar a observação da OS')) return;
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.obs = (valor || '').trim();
  try { await saveState('ordens'); } catch (e) { console.warn('salvarObsOS', e); }
}

// Calcula os tons efetivamente marcados como prefixo consecutivo: cada tom só
// vale se todos os anteriores estiverem marcados (Tom 4 exige Tom 1, 2 e 3).
// Sanitiza dados antigos ou estado inconsistente sem precisar limpar.
function tonsEfetivos(ttTons) {
  const out = [];
  if (ttTons && ttTons[1]) out.push(1);
  if (ttTons && ttTons[1] && ttTons[2]) out.push(2);
  if (ttTons && ttTons[1] && ttTons[2] && ttTons[3]) out.push(3);
  if (ttTons && ttTons[1] && ttTons[2] && ttTons[3] && ttTons[4]) out.push(4);
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

// Reescreve as quantidades CONGELADAS da OS a partir do que ela tem hoje:
// grade × camadas do enfesto × multiplicador do tecido.
//
// Existem duas contas de "quantas peças esta OS tem", e elas se separaram:
//   • a folha de OS calcula AO VIVO (`totaisPorTamanhoTomOS`) — sempre certa;
//   • `componentes[].qtdTotal` e `enfesto.totalPecas` ficam GRAVADOS, escritos
//     uma única vez pelo formulário no momento em que a OS foi salva.
// Quem lê a segunda é o resto do sistema: o estoque de corte e — o que se viu
// aqui — o seletor de OS da expedição, que só oferece OS com peças > 0. Numa OS
// criada sem camadas e completada depois na folha, a folha mostrava 480 peças e
// a expedição enxergava 0, então a OS simplesmente não aparecia para alocar.
// Devolve true quando alguma quantidade mudou.
function recomputarQuantidadesOS(o) {
  if (!o) return false;
  const keys = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
  const g = o.grade || {};
  const camadas = parseInt((o.enfesto || {}).camadas, 10) || 0;
  const mult = multiplicadorPecaOS(o);
  const pecasPorTamanho = {};
  keys.forEach(k => { pecasPorTamanho[k] = (parseInt(g[k], 10) || 0) * camadas * mult; });
  let mudou = false;
  const refazer = lista => (lista || []).forEach(c => {
    const porPeca = Number(c.qtdPorPeca) || 0;
    const novo = {};
    let total = 0;
    keys.forEach(k => { const v = pecasPorTamanho[k] * porPeca; novo[k] = v; total += v; });
    if (JSON.stringify(c.qtdPorTamanho || {}) !== JSON.stringify(novo) || (Number(c.qtdTotal) || 0) !== total) {
      c.qtdPorTamanho = novo;
      c.qtdTotal = total;
      mudou = true;
    }
  });
  refazer(o.componentes);
  refazer(o.aviamentos);
  const totalPecas = (parseInt(g.total, 10) || 0) * camadas;
  if (o.enfesto && (Number(o.enfesto.totalPecas) || 0) !== totalPecas) {
    o.enfesto.totalPecas = totalPecas;
    mudou = true;
  }
  return mudou;
}

// Fonte ÚNICA dos números do "Total por tamanho": quantidade por tamanho, por
// tonalidade, total de cada tom e total geral. A folha de OS e a folha de OE
// (plano de expedição) leem daqui, então não têm como mostrar números diferentes.
//
// Regra do balanceador: o ÚLTIMO tom marcado recebe, em cada tamanho, o total da
// coluna menos o que os tons editáveis levaram — assim a soma das colunas bate
// com a linha "Total geral".
//
// O V digitado é a quantidade do TAMANHO LIMITANTE (o de menor número na grade),
// e cada coluna escala pela grade: v(tamanho) = V × grade[tamanho] / grade[menor].
// Numa grade 2M-4G-2GG com 12 camadas no Tom 1, o V é 24 e a linha sai
// M=24 · G=48 · GG=24 — o G tem o dobro porque a grade pede o dobro.
//
// Antes o V era uniforme na linha inteira (M=24 · G=24 · GG=24). Era o mesmo
// número em grade uniforme, mas em grade desigual repartia errado entre os tons:
// o G do Tom 1 saía com 24 no lugar de 48, e o balanceador absorvia a diferença
// (96 em vez de 72). A soma da coluna continuava fechando — por isso passava
// despercebido —, mas o quanto cabia a cada tonalidade estava trocado.
function totaisPorTamanhoTomOS(o) {
  const keys = ['p','m','g','gg','g1','g2','g3'];
  const g = (o && o.grade) || {};
  const cam = (o && o.enfesto && o.enfesto.camadas) || 0;
  const mult = multiplicadorPecaOS(o);
  const prog = (o && o.progresso) || {};
  const colTotal = k => (g[k] || 0) * cam * mult;
  const tamanhos = keys.filter(k => (g[k] || 0) > 0);
  // Tamanho limitante: é a unidade em que o V é digitado.
  const qtdMin = tamanhos.length ? Math.min(...tamanhos.map(k => g[k] || 0)) : 1;
  const tons = tonsEfetivos(prog.totalTamanhoTons || {});
  const valores = prog.totalTamanhoTomValor || {};
  const balancer = tons.length ? tons[tons.length - 1] : null;
  const vTom = tom => Math.max(0, Number(valores[tom]) || 0);
  // O V daquele tom NAQUELE tamanho, já escalado pela grade.
  const vCel = (tom, k) => Math.round(vTom(tom) * (g[k] || 0) / (qtdMin || 1));
  // Quanto os tons editáveis levam numa coluna — é o que o balanceador desconta.
  const somaEditaveisCel = k => {
    let s = 0;
    tons.forEach(t => { if (t !== balancer) s += vCel(t, k); });
    return s;
  };
  let somaEditaveis = 0;
  tons.forEach(t => { if (t !== balancer) somaEditaveis += vTom(t); });
  // Enquanto NADA foi digitado, as linhas de tom saem VAZIAS: a quantidade por
  // tamanho fica só na linha "Total por tamanho", logo acima. Repetir o mesmo
  // número numa linha de tom duplicava a linha de cima e confundia quem lê. A
  // divisão só aparece quando alguém diz COMO dividir.
  //
  // Com UMA tonalidade só, porém, não há divisão a informar: aquele tom carrega
  // a coluna inteira, e isso não é chute nenhum. Antes ele caía nesta regra —
  // como o único tom é o balanceador, nenhum V editável existia, a soma dava
  // zero e a linha saía em branco. Era o que fazia os números SUMIREM ao lançar
  // as camadas do Tom 1 na folha: o lançamento acertava as camadas da OS, a
  // linha do Tom 1 aparecia e vinha vazia; só ao digitar o Tom 2 os números
  // voltavam.
  const semDigitacao = somaEditaveis === 0 && tons.length > 1;
  const linhas = tons.map(tom => {
    const cels = {};
    let total = 0;
    tamanhos.forEach(k => {
      let v;
      if (semDigitacao) v = 0;
      else if (tom === balancer) v = Math.max(0, colTotal(k) - somaEditaveisCel(k));
      else v = vCel(tom, k);
      cels[k] = v;
      total += v;
    });
    return { tom, cels, total, balanceador: tom === balancer, editavel: tom !== balancer };
  });
  return {
    keys, tamanhos, tons, linhas, colTotal, vTom, vCel, somaEditaveisCel, qtdMin,
    balancer, somaEditaveis, semDigitacao,
    totalGeral: (g.total || 0) * cam * mult,
  };
}

async function togglarTotalTamanhoTom(osId, tom, checked) {
  if (!exigirEdicao('editar o total por tamanho')) return;
  const os = STATE.ordens.find(x => x.id === osId);
  if (!os) return;
  os.progresso = os.progresso || {};
  os.progresso.totalTamanhoTons = os.progresso.totalTamanhoTons || {};
  const tNum = Number(tom);
  const t = os.progresso.totalTamanhoTons;
  if (checked) {
    // Bloqueia se prereq nao atendido (cada tom exige todos os anteriores)
    if (tNum === 2 && !t[1]) {
      if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
      return;
    }
    if (tNum === 3 && (!t[1] || !t[2])) {
      if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
      return;
    }
    if (tNum === 4 && (!t[1] || !t[2] || !t[3])) {
      if (printOsAtual && printOsAtual.id === osId) renderPrintSheet(os);
      return;
    }
    t[tNum] = true;
  } else {
    delete t[tNum];
    // Cascade: desmarcar um tom derruba todos os posteriores.
    if (tNum === 1) { delete t[2]; delete t[3]; delete t[4]; }
    else if (tNum === 2) { delete t[3]; delete t[4]; }
    else if (tNum === 3) { delete t[4]; }
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
// `size` é o tamanho da célula em que o número foi digitado: o usuário digita a
// quantidade DAQUELE tamanho, e o que fica gravado é sempre o V na unidade do
// tamanho limitante (é dele que as outras colunas são derivadas). Sem o `size`,
// digitar 48 no G de uma grade 2M-4G-2GG gravaria V=48 e o M sairia com 48
// também.
async function salvarValorTotalTamanhoTom(osId, tom, valor, size) {
  if (!exigirEdicao('editar o total por tamanho')) return;
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
  // O V é digitado na unidade do TAMANHO LIMITANTE e cada coluna escala por
  // grade[k]/qtdMin. Com isso a trava vira a mesma em todas as colunas —
  // V × g[k]/qtdMin ≤ g[k] × cam × mult ⟺ V ≤ qtdMin × cam × mult —, então
  // basta o teto do tamanho limitante em vez de varrer as colunas.
  const visiveis = ['p','m','g','gg','g1','g2','g3'].map(k => g[k] || 0).filter(q => q > 0);
  const qtdMin = visiveis.length ? Math.min(...visiveis) : 0;
  const max = Math.max(0, qtdMin * cam * mult - somaOutros);
  // O número digitado é do tamanho `size`; converte para a unidade do limitante.
  const qtdSize = (size && (g[size] || 0) > 0) ? g[size] : qtdMin;
  const digitado = Math.max(0, Math.floor(Number(valor) || 0));
  const bruto = qtdSize > 0 ? Math.round(digitado * qtdMin / qtdSize) : digitado;
  const n = Math.max(0, Math.min(max, bruto));
  os.progresso.totalTamanhoTomValor[tNum] = n;
  // Reescreve a linha inteira com o valor de CADA tamanho (o clamp pode ter
  // baixado o V, e cada coluna tem o seu número).
  document.querySelectorAll(`input[data-tt-tom-input="${tNum}"]`).forEach(i => {
    const k = i.dataset.ttSize;
    const q = (k && (g[k] || 0) > 0) ? g[k] : qtdMin;
    const v = qtdMin > 0 ? Math.round(n * q / qtdMin) : n;
    const txt = v > 0 ? String(v) : '';
    if (i.value !== txt) i.value = txt;
  });
  atualizarLinhasTomNoDOM();
  try { await saveState('ordens'); } catch (e) { console.warn('salvarValorTotalTamanhoTom', e); }
}

// Sincroniza visualmente o valor V entre todos os inputs da mesma linha de
// tom (digitar em uma celula preenche as outras com o mesmo numero) e
// recalcula no DOM as celulas do balanceador e as colunas "Total" das
// linhas — sem re-render pra preservar o foco do input.
// Digitar numa célula preenche as outras da linha PROPORCIONALMENTE à grade: o
// número digitado vale para o tamanho daquela célula, e cada outra recebe
// digitado × grade[dela] / grade[digitado]. Numa grade 2M-4G-2GG, digitar 24 no
// M põe 48 no G — antes repetia 24 em todas.
function propagarValorTomTamanho(input, tom) {
  const os = printOsAtual;
  const g = (os && os.grade) || {};
  const kIn = input.dataset.ttSize;
  const qIn = (kIn && (g[kIn] || 0) > 0) ? g[kIn] : 1;
  const digitado = Math.max(0, Number(input.value) || 0);
  document.querySelectorAll(`input[data-tt-tom-input="${tom}"]`).forEach(i => {
    if (i === input) return;
    const k = i.dataset.ttSize;
    const q = (k && (g[k] || 0) > 0) ? g[k] : qIn;
    const v = qIn > 0 ? Math.round(digitado * q / qIn) : digitado;
    i.value = v > 0 ? String(v) : '';
  });
  atualizarLinhasTomNoDOM();
}

function atualizarLinhasTomNoDOM() {
  const os = printOsAtual;
  if (!os) return;
  const balancerCells = document.querySelectorAll('[data-tt-balancer-cell]');
  if (!balancerCells.length && !document.querySelector('[data-tt-row-total]')) return;
  const g = os.grade || {};
  // Tons editaveis. O input guarda a quantidade DO TAMANHO dele; o que interessa
  // aqui é quanto cada tom leva EM CADA COLUNA, então guarda-se o par
  // (valor, tamanho) e a conversão é feita por coluna.
  const vPorTom = {};
  [1,2,3,4].forEach(tt => {
    const first = document.querySelector(`input[data-tt-tom-input="${tt}"]`);
    if (first) vPorTom[tt] = { val: Math.max(0, Number(first.value) || 0), size: first.dataset.ttSize };
  });
  const naColuna = (e, k) => {
    const qOrig = (e.size && (g[e.size] || 0) > 0) ? g[e.size] : 0;
    const qDest = g[k] || 0;
    if (!(qOrig > 0)) return e.val;
    return Math.round(e.val * qDest / qOrig);
  };
  // Atualiza cada celula do balanceador: colTotal(k) - o que os editaveis levam
  // NAQUELA coluna.
  balancerCells.forEach(c => {
    const size = c.dataset.ttBalancerSize;
    const colTotal = calcularColTotalAlvoImpressao(os, size);
    let somaEditaveis = 0;
    Object.values(vPorTom).forEach(e => { somaEditaveis += naColuna(e, size); });
    const v = Math.max(0, colTotal - somaEditaveis);
    c.textContent = v > 0 ? String(v) : '';
  });
  // Atualiza as colunas "Total" de cada linha de tom (soma da linha)
  document.querySelectorAll('[data-tt-row-total]').forEach(c => {
    const tt = Number(c.dataset.ttRowTotal);
    let sum = 0;
    if (vPorTom[tt] != null) {
      // Tom editavel: soma o valor de cada celula visivel (cada uma tem o seu).
      document.querySelectorAll(`input[data-tt-tom-input="${tt}"]`).forEach(i => {
        sum += Math.max(0, Number(i.value) || 0);
      });
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
  // Horário que chegou de outro usuário também tem que ser conferido contra a
  // média — o alerta não pode depender de quem digitou.
  _atualizarSelosTempoEnfesto(os.id);
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
    // Consolida: versões antigas do MESMO número (com data no nome, ou sem o
    // zero à esquerda) saem da pasta. São a mesma OS.
    const velhas = await _limparVariantesPdfDaOS(handle, o, filename);
    toast(`PDF salvo: ${filename}`
      + (velhas.length ? ` · ${velhas.length} versão(ões) antiga(s) do mesmo número removida(s)` : ''), 'ok');
  } catch (e) {
    console.warn('autoSalvarPdfPrintAtual', e);
  }
}

function editarOsAtual() {
  if (printOsAtual) editarOS(printOsAtual.id);
}

function editarOS(id) {
  if (!exigirEdicao('editar OS')) return;
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
  if (!exigirEdicao('duplicar OS')) return;
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
// Camadas que uma FASE do enfesto tem quando a OS não gravou um número próprio
// para ela. Cada fase enfesta um tanto diferente: a ribana moletom cadastrada
// como "2x" rende duas unidades da grade por camada, então 36 camadas de moletom
// pedem 18 de ribana; o forro de capuz vai a metade; o viés é sempre 1.
//
// É a mesma tabela do botão "calcular camadas" do formulário
// (`calcularCamadasParaProducao`). Antes ela só existia lá, e o que sobrava para
// a folha era `b.camadas || camadasGlobal`: numa OS salva sem as camadas por
// fase, TODAS as fases herdavam o número do moletom. A ribana da OS 0443
// aparecia com 36 camadas em vez de 18 — e como o kg sai de
// comp × larg × camadas × gramatura, a baixa de estoque da ribana saía dobrada.
function camadasPadraoDaFase(o, ordem, camadasPrincipal) {
  const cam = parseInt(camadasPrincipal, 10) || 0;
  if (!(cam > 0)) return 0;
  const fasesOS = (o && o.fases) || [];
  const idx = fasesOS.findIndex(f => (f.ordem || 0) === ordem);
  if (idx < 0) return cam;
  // Mesma complementação POR FASE que o formulário faz: quando a fase não tem um
  // tecido que exista de fato no cadastro, vale o da linha de Tecidos da OS
  // naquela posição. Sem isto, a fase caía sem categoria, o papel saía vazio e a
  // gola voltava a herdar as camadas do tecido principal.
  const linhasTec = (o && o.tecidos) || [];
  const fasesEfetivas = fasesOS.map((f, i) => {
    const valido = f.tecidoId && (STATE.tecidos || []).some(t => t.id === f.tecidoId);
    return { ...f, tecidoId: valido ? f.tecidoId : ((linhasTec[i] || {}).tecidoId || '') };
  });
  const faseOS = fasesEfetivas[idx] || {};
  // As "unidades da grade" (o 2x da ribana) moram no CADASTRO da grade — a
  // cópia de fases guardada na OS não as leva.
  const grade = (STATE.grades || []).find(x => x.id === (o && o.gradeId));
  const fasesGrade = (grade && grade.fases) || [];
  const faseGrade = fasesGrade.find(f => (f.ordem || 0) === ordem) || fasesGrade[idx] || {};
  const papel = (calcularPapeisFases(fasesEfetivas)[idx] || {}).papel || '';
  const tec = (STATE.tecidos || []).find(t => t.id === faseOS.tecidoId);
  if (papel === 'forro_capuz') {
    // Mesmo "2x" da ribana, agora cadastrável: a fase da GRADE é que guarda as
    // unidades (a cópia de fases da OS não as leva). Sem cadastro, 2.
    const unidForro = parseInt(faseGrade.unidades, 10) || UNIDADES_PADRAO_FORRO;
    return Math.max(1, Math.ceil(cam / unidForro));
  }
  if (papel.indexOf('ribana_') === 0 && isTecidoRibana(tec)) {
    const unidades = parseInt(faseGrade.unidades, 10) || (MULTIPLICADOR_PECAS.ribana || 2);
    // Ribana MOLETOM escala com a grade: "2 barras + 4 punhos por tamanho"
    // cobre duas blusas/tamanho quando a grade pede 2 por tamanho. As outras
    // ribanas (gola polo, ribana de malha) já têm a grade toda embutida no "10x".
    const escalaComGrade = String((tec && tec.nome) || '').toLowerCase().includes('moletom');
    const g = (o && o.grade) || {};
    const qtds = ['p','m','g','gg','g1','g2','g3'].map(k => parseInt(g[k], 10) || 0).filter(q => q > 0);
    const fator = (escalaComGrade && qtds.length)
      ? qtds.reduce((s, x) => s + x, 0) / qtds.length
      : 1;
    return Math.max(1, Math.ceil(cam * multiplicadorPecaOS(o) * fator / unidades));
  }
  return cam;
}

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
    // A cor CANÔNICA da fase vem da linha de Tecidos da OS (derivada do desenho e
    // mantida em dia). O nomeCor do bloco é um snapshot que pode ficar VELHO —
    // ex.: desenho copiado e a ribana trocada depois: a linha de tecido vira
    // "Vermelho Ribana Moletom" mas o bloco fica "Mostarda". Prefere a cor da
    // linha de tecido quando ela é real e o tecido dela bate com o da fase
    // (evita pegar a linha errada num índice desalinhado).
    const _tRow = tecs[i];
    const _corLinha = (_tRow && _tRow.corNome && _tRow.corNome !== '—'
      && (!fase.tecidoNome || !_tRow.tecidoNome || _tRow.tecidoNome === fase.tecidoNome))
      ? _tRow.corNome : '';
    let cor = _corLinha || b.nomeCor || fase.corNome || '';
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
    // Camadas: as gravadas na fase mandam. Sem elas, cada fase deriva das
    // camadas principais pela sua própria regra — herdar o número do moletom
    // dobrava a ribana cadastrada como "2x".
    const camadas = ehVies ? 1 : (b.camadas || camadasPadraoDaFase(o, ord, camadasGlobal) || camadasGlobal || 0);
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

// Descritor da grade para agrupar OSs "iguais". Prefere a IDENTIDADE da grade
// cadastrada: o texto em o.grade.descricao é um SNAPSHOT do nome na hora em que a
// OS foi criada, então renomear a grade (ou um espaço duplo digitado) fazia a
// MESMA grade virar duas chaves e o histórico de tempo das duas metades nunca se
// encontrava — era o caso das OS 0405 ("P ao G3 | BM.TRICOLOR") e 0435 ("2X P ao
// G3 | BM.TRICOLOR"), a mesma grade no cadastro. OS sem gradeId (antiga) ainda
// entra no grupo quando o texto dela casa com o nome de uma grade cadastrada;
// sem isso, cai no texto normalizado e, na falta dele, nas quantidades.
function _osGradeKey(o) {
  const g = (o && o.grade) || {};
  const txt = _normNome(g.descricao || '');
  let viva = (o && o.gradeId) ? (STATE.grades || []).find(x => x.id === o.gradeId) : null;
  if (!viva && txt) viva = (STATE.grades || []).find(x => _normNome(x.nome || x.descricao || '') === txt);
  if (viva) return 'g:' + viva.id;
  return txt || ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].map(k => g[k] || 0).join('-');
}

// Nome de fase normalizado para COMPARAR entre OSs. Além do que _normNome já faz
// (acento, caixa, espaço duplo), os separadores viram espaço: a mesma fase foi
// cadastrada como "Barra + Punhos" numa OS e "Barra/Punhos" noutra, e sem isso o
// tempo medido numa não somava com o da outra.
function _normFaseNome(nome) {
  return _normNome(String(nome || '').replace(/[+/&,;-]+/g, ' '));
}

// Média histórica da DURAÇÃO de cada fase do enfesto, entre OSs com a MESMA
// grade e o MESMO conjunto de fases, a partir dos tempos Início/Fim já lançados
// (progresso.enfestosTempos). Chave = nome da fase (minúsculo). Alimenta a
// sugestão de "tempo automático" por fase na folha. { nome: {mediaMin, n} }.
function _mediaTempoFasesSimilares(o) {
  const gradeKey = _osGradeKey(o);
  const nomesFase = new Set((o.fases || []).map(f => _normFaseNome(f.nome)).filter(Boolean));
  if (!nomesFase.size) return {};
  const acc = {};
  (STATE.ordens || []).forEach(x => {
    if (!x || x.id === o.id) return;
    if (_osGradeKey(x) !== gradeKey) return;
    const nomesX = new Set((x.fases || []).map(f => _normFaseNome(f.nome)).filter(Boolean));
    if (nomesX.size !== nomesFase.size) return;             // mesmo conjunto de fases
    let igual = true; nomesFase.forEach(n => { if (!nomesX.has(n)) igual = false; });
    if (!igual) return;
    const tempos = (x.progresso || {}).enfestosTempos || {};
    (x.fases || []).forEach(f => {
      const nome = _normFaseNome(f.nome);
      if (!nome) return;
      const t = tempos[f.ordem] || {};
      const ini = _opMin(t.enfIni), fim = _opMin(t.enfFim);
      if (ini == null || fim == null || fim <= ini) return;
      if (!acc[nome]) acc[nome] = { soma: 0, n: 0 };
      acc[nome].soma += (fim - ini); acc[nome].n++;
    });
  });
  const out = {};
  Object.keys(acc).forEach(n => { out[n] = { mediaMin: Math.round(acc[n].soma / acc[n].n), n: acc[n].n }; });
  return out;
}

/* ---- conferência do tempo lançado contra a média das OS de referência ---- */

// Margem aceita antes de acusar conflito: 25% da média, com piso de 10 min — sem
// o piso, uma fase de 30 min viraria alerta por 4 minutos de diferença, que é
// ruído de anotação, não desvio de processo.
const _ENF_TOL_PCT = 0.25, _ENF_TOL_MIN = 10;

// Duração lançada numa fase (minutos), ou null quando falta horário. `invertido`
// marca o caso fim ≤ início, que é erro de digitação e não duração.
function _enfDuracaoFase(o, ord) {
  const t = ((o.progresso || {}).enfestosTempos || {})[ord] || {};
  const ini = _opMin(t.enfIni), fim = _opMin(t.enfFim);
  if (ini == null || fim == null) return { min: null, invertido: false };
  if (fim <= ini) return { min: null, invertido: true };
  return { min: fim - ini, invertido: false };
}

// Confere o tempo lançado na fase contra a média das OS de referência (mesma
// grade, mesmas fases) e devolve o veredito: 'sem-ref' (nada com que comparar),
// 'ok', 'acima', 'abaixo' ou 'invertido'.
function _enfConferirTempoFase(o, ord) {
  // ord pode vir como número (render) ou string (dataset do DOM).
  const fase = (o.fases || []).find(f => String(f.ordem) === String(ord)) || {};
  const { min, invertido } = _enfDuracaoFase(o, ord);
  if (invertido) return { veredito: 'invertido' };
  if (min == null) return { veredito: 'vazio' };
  const med = _mediaTempoFasesSimilares(o)[_normFaseNome(fase.nome)];
  if (!med) return { veredito: 'sem-ref', min };
  const tol = Math.max(_ENF_TOL_MIN, Math.round(med.mediaMin * _ENF_TOL_PCT));
  const dif = min - med.mediaMin;
  const veredito = Math.abs(dif) <= tol ? 'ok' : (dif > 0 ? 'acima' : 'abaixo');
  return { veredito, min, dif, tol, mediaMin: med.mediaMin, n: med.n };
}

// Conteúdo do selo que vai ao lado dos campos de horário da fase: a duração
// lançada e, quando ela foge da média das OS de referência, o conflito em
// vermelho com o tamanho do desvio. Fica num span próprio para ser reescrito
// sozinho a cada digitação, sem redesenhar a folha.
function _enfSeloTempoHtml(o, ord) {
  const r = _enfConferirTempoFase(o, ord);
  const dur = v => esc(_opDurTexto(v));
  const cinza = 'color:#555;font-size:5.5pt;';
  const vermelho = 'color:#c81e1e;font-weight:700;font-size:5.5pt;';
  if (r.veredito === 'vazio') return '';
  if (r.veredito === 'invertido') {
    return `<span style="${vermelho}" title="O horário de fim é anterior (ou igual) ao de início">⚠ fim antes do início</span>`;
  }
  const gasto = `<span style="${cinza}">= ${dur(r.min)}</span>`;
  if (r.veredito === 'sem-ref' || r.veredito === 'ok') return gasto;
  const sinal = r.dif > 0 ? '+' : '−';
  const palavra = r.veredito === 'acima' ? 'acima' : 'abaixo';
  return `${gasto} <span style="${vermelho}" title="Fora da média desta fase: ${_opDurTexto(r.mediaMin)} em ${r.n} OS de referência (mesma grade e mesmas fases), com margem de ${_opDurTexto(r.tol)}">`
    + `⚠ ${sinal}${dur(Math.abs(r.dif))} ${palavra} da média</span>`;
}

// Reescreve os selos de conflito da folha aberta (todas as fases da OS). Chamado
// ao digitar um horário e quando a folha recebe dados de outro usuário.
function _atualizarSelosTempoEnfesto(osId) {
  const o = (STATE.ordens || []).find(x => x.id === osId);
  if (!o) return;
  document.querySelectorAll('[data-enf-selo]').forEach(el => {
    el.innerHTML = _enfSeloTempoHtml(o, el.dataset.enfSelo);
  });
}

// Tempo de cada fase do enfesto de uma GRADE, medido nas OS dela: cruza os
// horários de Início/Fim lançados na folha de OS (progresso.enfestosTempos) de
// todas as OS daquela grade. É número APURADO, não digitado — não existe campo
// para editar, e a cada abertura do cadastro ele é recalculado do zero, então não
// há como ficar desatualizado nem como um usuário alterar.
// Devolve [{ nome, mediaMin, n, minMin, maxMin, osNums }] na ordem das fases
// cadastradas na grade; fase medida que não está mais no cadastro vai no fim.
function temposFasesDaGrade(gradeId) {
  const grade = (STATE.grades || []).find(g => g.id === gradeId);
  if (!grade) return [];
  const chave = 'g:' + grade.id;
  const acc = new Map();     // nome normalizado → { nome, mins:[], osNums:Set }
  const pegar = (norm, nomeExib) => {
    let e = acc.get(norm);
    if (!e) { e = { nome: nomeExib, mins: [], osNums: new Set() }; acc.set(norm, e); }
    return e;
  };
  // As fases do cadastro entram primeiro (mesmo sem medição): a tabela do cadastro
  // tem que mostrar a fase que ainda não foi cronometrada, senão ela parece não
  // existir em vez de parecer não medida.
  (grade.fases || []).forEach(f => {
    const norm = _normFaseNome(f.nome);
    if (norm) pegar(norm, (f.nome || '').trim());
  });
  (STATE.ordens || []).forEach(o => {
    if (_osGradeKey(o) !== chave) return;
    const tempos = (o.progresso || {}).enfestosTempos || {};
    (o.fases || []).forEach(f => {
      const norm = _normFaseNome(f.nome);
      if (!norm) return;
      const t = tempos[f.ordem] || {};
      const ini = _opMin(t.enfIni), fim = _opMin(t.enfFim);
      if (ini == null || fim == null || fim <= ini) return;
      const e = pegar(norm, (f.nome || '').trim());
      e.mins.push(fim - ini);
      const num = (o.os || '').toString().trim();
      if (num) e.osNums.add(num);
    });
  });
  return Array.from(acc.values()).map(e => ({
    nome: e.nome,
    n: e.mins.length,
    mediaMin: e.mins.length ? Math.round(e.mins.reduce((s, m) => s + m, 0) / e.mins.length) : null,
    minMin: e.mins.length ? Math.min(...e.mins) : null,
    maxMin: e.mins.length ? Math.max(...e.mins) : null,
    osNums: Array.from(e.osNums).sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }))
  }));
}

// Bloco só-leitura do cadastro da grade com o tempo apurado por fase.
function _gradeTemposHtml(grade) {
  const cab = `<label style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3);">Tempo por fase (medido nas OS)</label>`;
  const nota = `<div class="field-hint" style="margin-top:4px;margin-bottom:8px;">
      Apurado dos horários de <b>Início</b> e <b>Fim</b> que a folha de OS registra em cada fase do enfesto, somando todas as OS desta grade.
      É <b>histórico medido</b>, não meta: recalculado a cada abertura e <b>sem campo para digitar ou alterar</b> — para mudar, lance os horários na folha da OS.
    </div>`;
  if (!grade || !grade.id) {
    return `<div style="margin-top:14px;">${cab}
      <div class="field-hint" style="margin-top:4px;">Disponível depois de salvar a grade e lançar os horários de enfesto em alguma OS dela.</div>
    </div>`;
  }
  const linhas = temposFasesDaGrade(grade.id);
  const medidas = linhas.filter(l => l.n > 0);
  const corpo = linhas.length ? `
    <table class="table" style="margin-top:2px;">
      <thead><tr>
        <th>Fase do enfesto</th>
        <th style="text-align:right;">Tempo médio</th>
        <th style="text-align:right;">Medições</th>
        <th style="text-align:right;">Menor → maior</th>
        <th>OS medidas</th>
      </tr></thead>
      <tbody>
        ${linhas.map(l => `
          <tr${l.n ? '' : ' style="color:var(--ink-3);"'}>
            <td><strong>${esc(l.nome)}</strong></td>
            <td style="text-align:right;font-family:'IBM Plex Mono',monospace;font-weight:700;">${l.n ? esc(_opDurTexto(l.mediaMin)) : '—'}</td>
            <td style="text-align:right;font-family:'IBM Plex Mono',monospace;">${l.n ? l.n : '—'}</td>
            <td style="text-align:right;font-family:'IBM Plex Mono',monospace;white-space:nowrap;">${l.n > 1 ? esc(_opDurTexto(l.minMin)) + ' → ' + esc(_opDurTexto(l.maxMin)) : '—'}</td>
            <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;">${l.osNums.length ? esc(l.osNums.join(', ')) : '—'}</td>
          </tr>`).join('')}
      </tbody>
    </table>
    ${medidas.length
      ? `<div class="field-hint" style="margin-top:6px;">Total medido nas fases com tempo: <b>${esc(_opDurTexto(medidas.reduce((s, l) => s + l.mediaMin, 0)))}</b>${
          medidas.length < linhas.length ? ` · ${linhas.length - medidas.length} fase(s) ainda sem nenhuma medição` : ''}.</div>`
      : `<div class="field-hint" style="margin-top:6px;">Nenhuma OS desta grade tem horário de enfesto lançado ainda — as fases aparecem sem tempo até a primeira medição.</div>`}`
    : `<div class="field-hint" style="margin-top:4px;">Esta grade ainda não tem fases cadastradas nem OS com horário lançado.</div>`;
  return `<div style="margin-top:14px;">${cab}${nota}${corpo}</div>`;
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
  // ganha campos em branco pros quatro tons (Tom 1/2/3/4), independente do que
  // está marcado no "Total por tamanho" — preenchíveis à mão e persistidos por fase.
  const enfestosTons = (o.progresso && o.progresso.enfestosTons) || {};
  const tomsSelEnf = [1, 2, 3, 4];
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
  // Tempo AUTOMÁTICO por fase: média histórica da duração desta fase entre OSs
  // com a mesma grade e as mesmas fases (dos tempos já lançados). Só sugestão —
  // aparece ao lado, não sobrescreve o que o usuário digita.
  const mediasFase = _mediaTempoFasesSimilares(o);
  const nomeFaseDe = ord => _normFaseNome(((o.fases || []).find(f => f.ordem === ord) || {}).nome);
  const linhaTempo = (lbl, ord, campoIni, campoFim, t) => {
    const med = mediasFase[nomeFaseDe(ord)];
    const hint = med
      ? ` <span style="color:#888;font-size:5.5pt;" title="Tempo médio desta fase em ${med.n} OS com a mesma grade e fases">⌀ ${esc(_opDurTexto(med.mediaMin))}</span>`
      : '';
    // Selo da conferência: a duração lançada e o alerta quando ela foge da média.
    // É reescrito sozinho a cada digitação (_atualizarSelosTempoEnfesto).
    return `<div style="display:flex;align-items:center;gap:5px;padding:1px 0;font-family:'IBM Plex Mono',monospace;font-size:6pt;line-height:1.3;">
      <span style="font-weight:700;min-width:44px;text-transform:uppercase;letter-spacing:.04em;">${lbl}</span>
      <span style="color:#555;">Início</span>${campoTempo(ord, campoIni, t[campoIni])}
      <span style="color:#555;">Fim</span>${campoTempo(ord, campoFim, t[campoFim])}${hint}
      <span data-enf-selo="${esc(String(ord))}">${_enfSeloTempoHtml(o, ord)}</span>
    </div>`;
  };
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
  // Quantas cores o DESENHO realmente tem — pela sequência do desc, senão pelos
  // campos de cor. O banner mostra SÓ essas: um desenho de UMA cor (Básica) cujo
  // enfesto usa moletom + forro + ribana (3 tecidos coloridos) herda 3 cores na
  // variante; sem este limite o banner sairia com 3 cores numa peça de 1 cor só.
  // A regra inteira mora em coresDaPecaOS — a folha de OE lê da mesma função.
  const coresDesenho = coresDaPecaOS(o);
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
              // O V é digitado na unidade do tamanho limitante; cada célula sai
              // escalada pela grade (TT.linhas já traz o número de cada uma).
              // Digitar numa célula propaga pras outras via DOM, na proporção.
              const vTom = TT.vTom;
              const tomRow = (tom) => {
                const isChecked = tomsSel.includes(tom);
                const ck = isChecked ? 'checked' : '';
                let bloqueado = false;
                if (tom === 2 && !ttTons[1]) bloqueado = true;
                if (tom === 3 && (!ttTons[1] || !ttTons[2])) bloqueado = true;
                if (tom === 4 && (!ttTons[1] || !ttTons[2] || !ttTons[3])) bloqueado = true;
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
                  return `<td style="padding:0;"><input type="number" min="0" value="${val > 0 ? val : ''}" data-tt-tom-input="${tom}" data-tt-size="${k}" oninput="propagarValorTomTamanho(this, ${tom})" onchange="salvarValorTotalTamanhoTom('${esc(o.id)}', ${tom}, this.value, '${k}')" style="width:100%;box-sizing:border-box;border:none;background:transparent;text-align:center;font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:8pt;padding:1px 2px;"></td>`;
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
                ${tomRow(1)}${tomRow(2)}${tomRow(3)}${tomRow(4)}`;
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

/* ========================================================= */
/*  IMPORTAR RISCO: ler o PDF do encaixe e cadastrar a grade  */
/* ========================================================= */
// O CAD que alimenta a máquina de corte exporta, para cada encaixe, um relatório
// em PDF com o comprimento e a largura daquele risco. Digitar isso à mão é o que
// deixa dezenas de fases sem medida no cadastro — e fase sem medida não calcula
// consumo de tecido nem tempo de enfesto.
//
// UM PDF = UMA FASE. Uma grade de cinco fases pede cinco relatórios.
//
// A FONTE É O CONTEÚDO, não o nome do arquivo. Verificado nos relatórios reais:
//   • QUAL GRADE — a tabela de tamanhos do relatório identifica UMA única grade
//     entre as cadastradas. O nome do modelo não identificaria: o mesmo
//     "BLUSÃO CANGURU TRICOLOR" serve a três grades diferentes.
//   • QUAL FASE — o campo "Tecido" do relatório muda a cada fase (2, 3, 1,
//     FORRO, RIBANA nos cinco riscos da canguru). Ele é um código do CAD, sem
//     significado para o sistema — até alguém dizer, UMA VEZ, a que fase
//     corresponde. A partir daí o par (modelo + tecido) reconhece sozinho.
//   • Enquanto esse par não foi ensinado, a MEDIDA resolve quando a grade já
//     existe: a largura dá a família do tecido e o comprimento mais próximo dá a
//     fase. Nos cinco riscos isso acertou os cinco, com a fase certa a 0,15 m e a
//     segunda opção a 2,4 m — folga de mais de dez vezes.
//   • O nome do arquivo é o ÚLTIMO recurso, e serve principalmente de conferência.

// EXCEDENTE DE ENFESTO. O comprimento do relatório é a medida de CORTAR — o
// risco propriamente dito. O que se cadastra na grade é a medida de ENFESTAR,
// que é maior: sobra pano nas duas pontas para a enfestadeira segurar e para o
// corte não morrer na borda. O excedente é somado sempre ao COMPRIMENTO. A
// LARGURA não recebe nada — ela é a do tecido, e o tecido não estica.
//
// A sobra é POR TECIDO, e sai do cadastro dele ("Excedente de enfesto (cm)"):
// malha, moletom e ribana não se comportam igual na enfestadeira, e quem enfesta
// sabe de quanto cada um precisa. Este número aqui é só o padrão de quem ainda
// não foi cadastrado.
//
// Os 15 cm conferem com o que já estava cadastrado à mão: no
// "2X P ao G3 | BM.TRICOLOR" o Corpo Parte 1 do PDF é 4,5493 m e o cadastro
// dizia 4,70 — 4,5493 + 0,15 dá 4,6993. Exato. As outras fases erravam de 2 a 5
// cm, que é o arredondamento de quem digitou.
const EXCEDENTE_ENFESTO_PADRAO_CM = 15;

// O excedente de um tecido, em METROS. Vazio (nunca cadastrado) cai no padrão;
// zero cadastrado é zero de verdade — há tecido que não leva sobra.
function excedenteEnfestoM(tecidoId) {
  const t = tecidoId ? (STATE.tecidos || []).find(x => x.id === tecidoId) : null;
  const v = t ? t.excedente : null;
  const cm = (v === '' || v == null || !isFinite(parseFloat(v)))
    ? EXCEDENTE_ENFESTO_PADRAO_CM
    : Math.max(0, parseFloat(v));
  return cm / 100;
}

// Rótulo curto do excedente, para a tela dizer de quanto foi a soma.
const excedenteEnfestoCm = tecidoId => Math.round(excedenteEnfestoM(tecidoId) * 100);

// A medida que vai para o CADASTRO, a partir da que o relatório informa.
// `tecidoId` é o tecido DAQUELA fase — é ele que manda no excedente.
const _riscoCompCadastro = (compPdf, tecidoId) =>
  (compPdf == null ? null : compPdf + excedenteEnfestoM(tecidoId));

let _riscoLeituras = [];

function _riscoNum(txt) {
  const m = String(txt == null ? '' : txt).replace(',', '.').match(/-?\d+(?:\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
}

// A memória do que já foi ensinado: "modelo do risco | código do tecido" → fase.
// Mora em meta, junto do resto do que o programa aprendeu com o uso.
function _riscoAprendidos() {
  STATE.meta = STATE.meta || {};
  if (!STATE.meta.riscoFases || typeof STATE.meta.riscoFases !== 'object') STATE.meta.riscoFases = {};
  return STATE.meta.riscoFases;
}
function _riscoChave(leitura) {
  const m = _normNome(leitura.modelo), t = _normNome(leitura.tecido);
  return (m && t) ? (m + '|' + t) : '';
}

// Lê um relatório de encaixe.
//
// O relatório tem TRÊS COLUNAS de rótulo+valor — ler o texto em sequência não
// serve, porque "Largura:" e o valor dele ficam separados por campos de outra
// coluna. A regra é geométrica: cada VALOR é atribuído ao rótulo mais próximo à
// ESQUERDA dele, preferindo o que está na MESMA linha e só então o que está
// logo acima. Sem fronteira de coluna, sem palpite.
//
// A preferência pela mesma linha não é detalhe. Nestes relatórios o rótulo e o
// valor saem na mesma altura — `Largura:@233  117 cm@271` —, e a regra antiga,
// que exigia o rótulo ACIMA, descartava esse par e ainda pegava o rótulo da
// coluna vizinha, que o CAD desalinha em 1 ponto. Resultado medido nos 13 PDFs
// da pasta CM.LISA: comprimento e largura vinham NULOS em todos, e o campo
// "Tecido" vinha com o valor da Área ("7.86 m²"). Sem comprimento e largura o
// assistente não tem o que levar para a fase — que é a razão de ele existir.
async function _riscoLerPdf(file) {
  const lib = window.pdfjsLib;
  if (!lib) throw new Error('A biblioteca de leitura de PDF não carregou — a primeira vez precisa de internet.');
  if (lib.GlobalWorkerOptions && !lib.GlobalWorkerOptions.workerSrc) {
    lib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  }
  const doc = await lib.getDocument({ data: await file.arrayBuffer() }).promise;
  const pag = await doc.getPage(1);
  const tc = await pag.getTextContent();
  const itens = tc.items
    .map(i => ({ x: i.transform[4], y: i.transform[5], t: String(i.str || '').trim() }))
    .filter(i => i.t);
  if (!itens.length) {
    throw new Error('Este PDF não tem texto — parece ser digitalizado. Só serve o relatório gerado pelo próprio CAD.');
  }

  const ys = itens.map(i => i.y);
  const alt = (Math.max(...ys) - Math.min(...ys)) || 1;
  const janela = alt * 0.02;              // até 2% da altura conta como "logo abaixo"
  const rot = itens.filter(i => i.t.endsWith(':'));
  const vals = itens.filter(i => !i.t.endsWith(':'));

  const MESMA_LINHA = 2.5;   // o CAD desalinha rótulo e valor em ~1 ponto
  const porRotulo = new Map();
  vals.forEach(V => {
    let escolhido = null, melhor = null;
    rot.forEach(L => {
      if (L.x > V.x + 1) return;              // rótulo tem que estar à esquerda
      const dy = L.y - V.y;
      // Duas faixas, e a de mesma linha ganha SEMPRE da de cima — daí o degrau
      // de 1e6 no peso. Dentro de cada faixa vence o rótulo mais perto.
      let faixa;
      if (Math.abs(dy) <= MESMA_LINHA) faixa = 0;
      else if (dy > 0 && dy <= janela) faixa = 1;
      else return;
      // Na mesma linha o desalinho de 1 ponto é ruído e não pesa: manda a
      // distância horizontal, isto é, o rótulo imediatamente à esquerda. Na
      // faixa de cima a altura volta a valer, para não pular uma linha inteira.
      const peso = faixa * 1e6 + (faixa === 0 ? 0 : dy * 1000) + (V.x - L.x);
      if (melhor == null || peso < melhor) { melhor = peso; escolhido = L; }
    });
    if (!escolhido) return;
    const k = escolhido.t.slice(0, -1);
    if (!porRotulo.has(k)) porRotulo.set(k, []);
    porRotulo.get(k).push(V);
  });
  const campo = nome => {
    const chave = Array.from(porRotulo.keys())
      .find(k => _normNome(k) === _normNome(nome));
    if (!chave) return '';
    const vs = porRotulo.get(chave).slice().sort((a, b) => b.y - a.y || a.x - b.x);
    const topo = vs[0].y;
    return vs.filter(v => Math.abs(v.y - topo) < 3).map(v => v.t).join(' ').trim();
  };

  // As FILAS do relatório. Agrupar por y arredondado não serve para a tabela de
  // tamanhos: o CAD desalinha o tamanho em 1 ponto do resto da fila, e "M" cai
  // numa chave e "1  4  MODELO" noutra. A fila é uma faixa de altura.
  const FILA = 3;                       // pontos de tolerância dentro de uma fila
  const filas = [];
  itens.slice().sort((a, b) => b.y - a.y).forEach(i => {
    const f = filas[filas.length - 1];
    if (f && Math.abs(f.y - i.y) <= FILA) { f.itens.push(i); return; }
    filas.push({ y: i.y, itens: [i] });
  });
  filas.forEach(f => f.itens.sort((a, b) => a.x - b.x));
  const txtFila = f => f.itens.map(i => i.t).join(' ').trim();
  // Compatibilidade com o resto da função, que lê por linha.
  const ordem = filas.map(f => f.y);
  const linhas = new Map(filas.map(f => [f.y, f.itens]));
  const txtDe = y => txtFila({ itens: linhas.get(y) });

  // Tabela de tamanhos. Cada fila é "P  1  4  NOME DO MODELO": o tamanho, os
  // COMPLETOS (quantas grades inteiras daquele tamanho o risco traz — é a
  // própria distribuição da grade), os moldes encaixados e o modelo.
  //
  // Lê-se pela fila, não pela linha seguinte. A regra antiga procurava uma linha
  // cujo texto fosse EXATAMENTE o tamanho e pegava o primeiro número da linha
  // DE BAIXO — que nestes relatórios é a fila do PRÓXIMO tamanho. Por isso a
  // pasta P-M-G-GG-G1-G2-G3 era lida como "1xM 1xG" e a 2xP-2xGG como "2xG": a
  // assinatura saía errada, e com ela o casamento com a grade cadastrada.
  const iCab = filas.findIndex(f => /(^|\s)Tamanho(\s|$)/i.test(txtFila(f)) && /Completos/i.test(txtFila(f)));
  const iFim = filas.findIndex((f, k) => k > iCab && iCab >= 0 && /^Encaixe$/i.test(txtFila(f)));
  const tamanhos = {};
  if (iCab >= 0) {
    const ate = iFim > iCab ? iFim : filas.length;
    for (let k = iCab + 1; k < ate; k++) {
      const toks = filas[k].itens;
      const iTam = toks.findIndex(t => /^(P|M|G|GG|G1|G2|G3)$/i.test(t.t));
      if (iTam < 0) continue;
      // COMPLETOS = o primeiro número inteiro à direita do tamanho.
      const num = toks.slice(iTam + 1).find(t => /^\d+$/.test(t.t));
      const n = num ? parseInt(num.t, 10) : 0;
      if (n > 0) tamanhos[toks[iTam].t.toLowerCase()] = n;
    }
  }

  // O nome do modelo sai do PRÓPRIO item "Área usada modelo XXX", e não do texto
  // da fila: a área ("7.86 m² (100.00%)") fica na mesma altura, e juntar a fila
  // colava ela no nome. Isso não é cosmético — o modelo é metade da chave do que
  // o programa aprende (modelo|tecido), e com a área junto a chave mudava a cada
  // arquivo, então nada era aprendido nunca.
  let modelo = '';
  for (const f of filas) {
    const it = f.itens.find(i => /rea usada modelo/i.test(i.t));
    if (!it) continue;
    modelo = it.t.replace(/.*rea usada modelo\s*/i, '').trim();
    if (!modelo) modelo = txtFila(f).replace(/.*rea usada modelo\s*/i, '').trim();
    break;
  }
  // Corta um rabo de área que tenha vindo junto, venha de onde vier.
  modelo = modelo.replace(/\s*\d+(?:[.,]\d+)?\s*m².*$/i, '').trim();

  const comp = _riscoNum(campo('Comprimento'));
  const larg = _riscoNum(campo('Largura'));
  return {
    arquivo: file.name,
    comprimento: comp != null ? comp / 100 : null,    // o relatório dá em cm
    largura: larg != null ? larg / 100 : null,
    area: _riscoNum(campo('Área')),
    aproveitamento: _riscoNum(campo('Aproveitamento')),
    gramatura: _riscoNum(campo('Peso')),              // g/m² (o CAD rotula como kg)
    tecido: campo('Tecido'),
    sentido: campo('Sentido'),
    modelo, tamanhos
  };
}

// A grade com EXATAMENTE esta distribuição de tamanhos. É o casamento mais seguro
// que existe aqui: aplicar o risco na grade errada trocaria 2,51 m por 4,55 m e
// dobraria o consumo de tecido.
function _riscoGradesQueCasam(tamanhos) {
  const keys = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
  const alvo = {};
  keys.forEach(k => { const n = parseInt((tamanhos || {})[k], 10) || 0; if (n > 0) alvo[k] = n; });
  if (!Object.keys(alvo).length) return [];
  return (STATE.grades || []).filter(g => {
    const t = g.tamanhos || {};
    return keys.every(k => (parseInt(t[k], 10) || 0) === (alvo[k] || 0));
  });
}

const _riscoF = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : null; };

// A que fase da grade este risco pertence, e por quê. A ordem das fontes é a
// ordem da confiança: o que já foi ensinado, depois a medida, depois o nome.
function _riscoResolverFase(leitura, grade) {
  const fases = (grade.fases || []).slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0));
  if (!fases.length) return { fase: null, origem: 'sem fases', folga: null };

  // 1) ENSINADO: o par (modelo do risco + código do tecido) já foi apontado uma
  //    vez. Daí em diante não precisa de mais nada — nem do nome do arquivo.
  const chave = _riscoChave(leitura);
  const memoria = _riscoAprendidos();
  if (chave && memoria[chave]) {
    const f = fases.find(x => _normNome(x.nome) === _normNome(memoria[chave]));
    if (f) {
      // Confere contra a largura: se o CAD trocou o código, a medida denuncia.
      const l = _riscoF(f.larg);
      const bate = l == null || leitura.largura == null || Math.abs(l - leitura.largura) < 0.06;
      return { fase: f, origem: bate ? 'aprendido' : 'aprendido (largura não bate)', folga: null, alerta: !bate };
    }
  }

  // 2) MEDIDA: largura dá a família do tecido; dentro dela, o comprimento mais
  //    próximo dá a fase. Só vale quando a fase já tem medida cadastrada.
  const mesmaLarg = fases.filter(f => {
    const l = _riscoF(f.larg);
    return l != null && leitura.largura != null && Math.abs(l - leitura.largura) < 0.06;
  });
  const pool = (mesmaLarg.length ? mesmaLarg : fases).filter(f => _riscoF(f.comp) != null);
  if (pool.length && leitura.comprimento != null) {
    // Compara com a medida DE CADASTRO (já com o excedente), não com a do
    // relatório: senão toda fase pareceria estar o excedente inteiro fora.
    // O excedente é o do tecido de CADA fase — duas fases do mesmo risco podem
    // ser de tecidos diferentes, com sobras diferentes.
    const ord = pool.map(f => ({
      f, d: Math.abs(_riscoF(f.comp) - _riscoCompCadastro(leitura.comprimento, f.tecidoId))
    })).sort((a, b) => a.d - b.d);
    const folga = ord.length > 1 ? ord[1].d : null;
    // Escolha APERTADA (a segunda opção quase tão perto) não decide sozinha.
    const apertada = folga != null && folga < ord[0].d * 3;
    if (!apertada) return { fase: ord[0].f, origem: 'medida', folga, dist: ord[0].d };
  }

  // 3) NOME DO ARQUIVO: último recurso. Casa por palavra inteira da fase.
  const a = _normNome(leitura.arquivo);
  let melhor = null, pontos = 0;
  fases.forEach(f => {
    const palavras = _normNome(f.nome).split(/\s+/).filter(w => w.length > 2 && w !== 'parte');
    const p = palavras.filter(w => a.includes(w)).length;
    if (p > pontos) { pontos = p; melhor = f; }
  });
  if (melhor) return { fase: melhor, origem: 'nome do arquivo', folga: null };

  // 4) fases sem medida ainda: sobra a largura sozinha, se ela isolar uma.
  if (mesmaLarg.length === 1) return { fase: mesmaLarg[0], origem: 'largura', folga: null };
  return { fase: null, origem: 'indefinida', folga: null };
}

/* ---------------- criar grade NOVA a partir dos riscos ---------------- */
// Quando os tamanhos do relatório não batem com nenhuma grade cadastrada, é
// grade nova. O PDF traz o que é medida (tamanhos, comprimento, largura) mas não
// traz o que é DECISÃO DA CASA: como aquele produto se chama aqui dentro, de que
// tecido é cada fase, quantas unidades da grade a fase rende, quantas peças vão
// no pacote. Esta tela pergunta isso — uma vez por produto — e guarda a resposta.
// Da segunda grade do mesmo produto em diante, tudo já vem preenchido.

// A memória do PRODUTO: "modelo do risco" → como ele se chama e o que é aqui.
function _riscoProdutos() {
  STATE.meta = STATE.meta || {};
  if (!STATE.meta.riscoProdutos || typeof STATE.meta.riscoProdutos !== 'object') STATE.meta.riscoProdutos = {};
  return STATE.meta.riscoProdutos;
}
// A memória do TECIDO por fase: "modelo|código do tecido" → id do tecido.
function _riscoTecidos() {
  STATE.meta = STATE.meta || {};
  if (!STATE.meta.riscoTecidos || typeof STATE.meta.riscoTecidos !== 'object') STATE.meta.riscoTecidos = {};
  return STATE.meta.riscoTecidos;
}

// O pedaço do nome da grade que descreve os tamanhos, na convenção da casa:
//   todos iguais a 1 e contíguos  -> "P ao G3"
//   todos iguais a N e contíguos  -> "2X P ao G3"
//   diferentes entre si           -> "2M-4G-2GG"
function _riscoNomeTamanhos(tamanhos) {
  const ordem = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
  const rot = { p: 'P', m: 'M', g: 'G', gg: 'GG', g1: 'G1', g2: 'G2', g3: 'G3' };
  const presentes = ordem.filter(k => (parseInt(tamanhos[k], 10) || 0) > 0);
  if (!presentes.length) return '';
  const qtds = presentes.map(k => parseInt(tamanhos[k], 10));
  const iguais = qtds.every(q => q === qtds[0]);
  const contiguo = presentes.every((k, i) =>
    i === 0 || ordem.indexOf(k) === ordem.indexOf(presentes[i - 1]) + 1);
  if (iguais && contiguo && presentes.length > 2) {
    const faixa = `${rot[presentes[0]]} ao ${rot[presentes[presentes.length - 1]]}`;
    return qtds[0] === 1 ? faixa : `${qtds[0]}X ${faixa}`;
  }
  return presentes.map(k => (qtds[presentes.indexOf(k)] === 1 ? '' : qtds[presentes.indexOf(k)]) + rot[k]).join('-');
}

// Assinatura dos tamanhos, para juntar num só grupo os riscos que são da mesma
// grade nova. Cinco PDFs da mesma peça viram UMA grade de cinco fases.
function _riscoAssinatura(tamanhos) {
  return ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3']
    .map(k => (parseInt((tamanhos || {})[k], 10) || 0)).join('-');
}

// O rascunho de cada grade nova, por assinatura de tamanhos. Precisa viver fora
// do HTML: a tabela é redesenhada quando se troca o tipo de peça (a previsão de
// fases muda com ele), e sem isto o que já foi digitado se perderia a cada
// troca.
let _riscoNovas = {};

// Monta (ou devolve) o rascunho de um grupo. As fases nascem da PREVISÃO do
// destino; cada PDF do grupo é encaixado na fase com que se parece, e o que não
// casa com nenhuma vira linha própria no fim. As fases previstas que ficaram sem
// PDF continuam na lista, sem medida — é o ponto do recurso: a grade nasce
// inteira, e as medidas que faltam entram depois, por outro risco ou à mão.
function _riscoNovaDraft(G) {
  if (_riscoNovas[G.assinatura]) return _riscoNovas[G.assinatura];
  const modelo = (G.itens[0].L.modelo || '').trim();
  const prod = _riscoProdutos()[_normNome(modelo)] || {};
  const d = {
    sku: prod.sku || '',
    tipoPeca: prod.tipoPeca || '',
    variacao: prod.variacao || '',
    pecasPorPacote: prod.pecasPorPacote || '',
    fases: []
  };
  _riscoNovas[G.assinatura] = d;
  _riscoNovaAplicarPrevisao(G, d);
  return d;
}

// Refaz a lista de fases a partir da previsão do destino, preservando o que já
// foi digitado (nome, tecido, unidades) e os PDFs já encaixados.
function _riscoNovaAplicarPrevisao(G, d) {
  const memTec = _riscoTecidos();
  const previstas = previsaoFases(d.tipoPeca, d.variacao, d.sku);
  const antigas = d.fases || [];
  const usadas = new Set();
  const linha = (nome, iPdf) => {
    const velha = antigas.find(f => _normNome(f.nome) === _normNome(nome));
    const L = iPdf != null ? G.itens[iPdf].L : null;
    return {
      nome,
      tecidoId: (velha && velha.tecidoId) || (L ? (memTec[_riscoChave(L)] || '') : ''),
      unidades: (velha && velha.unidades) || 2,
      iPdf
    };
  };
  // 1) cada PDF procura a fase prevista com que se parece
  const destinoDoPdf = new Map();
  G.itens.forEach((it, ip) => {
    const sug = _normNome(_riscoNomeFaseSugerido(it.L));
    const alvo = previstas.find(n => sug && _normNome(n) === sug && !usadas.has(n));
    if (alvo) { usadas.add(alvo); destinoDoPdf.set(ip, alvo); }
  });
  // 2) o que sobrou de PDF ocupa a primeira fase prevista ainda livre
  G.itens.forEach((it, ip) => {
    if (destinoDoPdf.has(ip)) return;
    const livre = previstas.find(n => !usadas.has(n));
    if (livre) { usadas.add(livre); destinoDoPdf.set(ip, livre); }
  });
  const fases = previstas.map(n => {
    const ip = [...destinoDoPdf.entries()].find(([, alvo]) => alvo === n);
    return linha(n, ip ? ip[0] : null);
  });
  // 3) PDF que não coube em nenhuma prevista vira linha própria
  G.itens.forEach((it, ip) => {
    if (destinoDoPdf.has(ip)) return;
    fases.push(linha(_riscoNomeFaseSugerido(it.L) || `Fase ${fases.length + 1}`, ip));
  });
  // 4) sem previsão nenhuma (produto que a regra não conhece): uma linha por PDF
  if (!fases.length) G.itens.forEach((it, ip) => fases.push(linha(_riscoNomeFaseSugerido(it.L), ip)));
  d.fases = fases;
}

// Lê a tela de volta para os rascunhos. Chamado antes de qualquer redesenho e
// antes de criar a grade — sem isto, trocar o tipo de peça apagaria o que foi
// digitado.
function _riscoNovaColetar() {
  const v = id => (document.getElementById(id)?.value ?? '').toString().trim();
  _riscoGruposNovos().forEach((G, gi) => {
    const d = _riscoNovas[G.assinatura];
    if (!d || !document.getElementById(`rn-sku-${gi}`)) return;
    d.sku = v(`rn-sku-${gi}`);
    d.tipoPeca = v(`rn-tipo-${gi}`);
    d.variacao = v(`rn-var-${gi}`);
    d.pecasPorPacote = v(`rn-pac-${gi}`);
    d.fases.forEach((f, fi) => {
      if (!document.getElementById(`rn-f-nome-${gi}-${fi}`)) return;
      f.nome = v(`rn-f-nome-${gi}-${fi}`);
      f.tecidoId = v(`rn-f-tec-${gi}-${fi}`);
      f.unidades = parseInt(v(`rn-f-un-${gi}-${fi}`), 10) || 2;
    });
  });
}

// Trocar o tipo de peça ou a variação refaz a previsão de fases.
function riscoNovaMudouDestino(gi, sel, kind) {
  if (sel && kind) onSelectGradeFolder(sel, kind);
  _riscoNovaColetar();
  const G = _riscoGruposNovos()[gi];
  if (G && _riscoNovas[G.assinatura]) _riscoNovaAplicarPrevisao(G, _riscoNovas[G.assinatura]);
  renderRiscoResultado();
}

function riscoNovaAddFase(gi) {
  _riscoNovaColetar();
  const G = _riscoGruposNovos()[gi];
  const d = G && _riscoNovas[G.assinatura];
  if (d) d.fases.push({ nome: '', tecidoId: '', unidades: 2, iPdf: null });
  renderRiscoResultado();
}

function riscoNovaDelFase(gi, fi) {
  _riscoNovaColetar();
  const G = _riscoGruposNovos()[gi];
  const d = G && _riscoNovas[G.assinatura];
  if (d) d.fases.splice(fi, 1);
  renderRiscoResultado();
}

// O nome sugerido para uma fase que ainda não existe: o que já foi ensinado,
// senão o que o nome do arquivo diz, senão o próprio código do tecido.
function _riscoNomeFaseSugerido(L) {
  const mem = _riscoAprendidos();
  const k = _riscoChave(L);
  if (k && mem[k]) return mem[k];
  const a = L.arquivo || '';
  const m = a.match(/fase\s+([^.\-]+?)\s*(?:-|\.pdf|$)/i);
  if (m) return m[1].trim().replace(/\s+/g, ' ');
  const t = (L.tecido || '').trim();
  return t && !/^\d+$/.test(t) ? t.charAt(0) + t.slice(1).toLowerCase() : '';
}

// PREVISÃO DE FASES. Uma grade não é feita só do que veio em PDF: o risco do
// CORPO chega sozinho, mas a grade precisa nascer com a gola e o viés também,
// senão a OS sai sem eles e alguém tem que lembrar de completar depois.
//
// A previsão vem do TIPO DE PEÇA + VARIAÇÃO, que é o que define de que partes a
// peça é feita. Os nomes são os que a casa já usa nas grades cadastradas — não
// invento vocabulário novo: "Corpo Parte 1" e "Forro de capuz" são como estão
// nas 40 e poucas grades que já existem, e é assim que saem na folha de OS.
//
// A regra não obriga a nada: é o ponto de partida da tabela, e cada linha pode
// ser renomeada, apagada, e outras podem ser somadas.
const PREVISAO_FASES = [
  { tp: 'camiseta',      vr: 'basica',   fases: ['Corpo', 'Gola', 'Viés'] },
  { tp: 'camiseta',      vr: 'bicolor',  fases: ['Corpo Parte 1', 'Corpo Parte 2', 'Gola', 'Viés'] },
  // A "camiseta recortada" (CM.REC) é esta: três recortes, gola e viés.
  { tp: 'camiseta',      vr: 'tricolor', fases: ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Gola', 'Viés'] },
  { tp: 'camiseta polo', vr: null,       fases: ['Corpo', 'Viés'] },
  { tp: 'blusa_moletom', vr: 'basica',   fases: ['Corpo', 'Barra/Punhos', 'Viés'] },
  { tp: 'blusa_moletom', vr: 'tricolor', fases: ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Forro de capuz', 'Barra/Punhos', 'Viés'] }
];

// Fases previstas para um destino. `vr: null` na regra vale para qualquer
// variação daquele tipo. O SKU entra como segunda chance: quem escreve CM.REC
// ou BM.TRICOLOR está dizendo o produto, mesmo que a pasta ainda não diga.
function previsaoFases(tipoPeca, variacao, sku) {
  const tp = _normNome(tipoPeca || ''), vr = _normNome(variacao || '');
  let regra = PREVISAO_FASES.find(r => _normNome(r.tp) === tp && r.vr != null && _normNome(r.vr) === vr)
           || PREVISAO_FASES.find(r => _normNome(r.tp) === tp && r.vr == null);
  if (!regra) {
    const s = String(sku || '').toUpperCase().replace(/\s+/g, '');
    const porSku = [
      [/^CM\.(REC|TRI)/, 'camiseta', 'tricolor'],
      [/^CM\.BIC/,       'camiseta', 'bicolor'],
      [/^CM\./,          'camiseta', 'basica'],
      [/^PM\./,          'camiseta polo', null],
      [/^BM\.(TRI|REC)/, 'blusa_moletom', 'tricolor'],
      [/^BM\./,          'blusa_moletom', 'basica']
    ].find(([re]) => re.test(s));
    if (porSku) regra = PREVISAO_FASES.find(r => _normNome(r.tp) === _normNome(porSku[1])
      && (porSku[2] == null ? r.vr == null : (r.vr != null && _normNome(r.vr) === _normNome(porSku[2]))));
  }
  return regra ? regra.fases.slice() : [];
}

// Blocos de "grade nova" a montar: um por assinatura de tamanhos. Entram os
// riscos sem grade candidata E os que quem está importando mandou virar grade
// nova mesmo havendo candidata (`forcarNova`) — a grade dos mesmos tamanhos pode
// existir para outro produto, e é decisão de quem cadastra, não do programa.
function _riscoGruposNovos() {
  const grupos = new Map();
  _riscoLeituras.forEach((L, i) => {
    if (L.erro) return;
    if (!L.forcarNova && L.grades && L.grades.length) return;
    if (!L.tamanhos || !Object.keys(L.tamanhos).length) return;
    const a = _riscoAssinatura(L.tamanhos);
    if (!grupos.has(a)) grupos.set(a, { assinatura: a, tamanhos: L.tamanhos, itens: [] });
    grupos.get(a).itens.push({ L, i });
  });
  return Array.from(grupos.values());
}

function _riscoHtmlGradesNovas() {
  const grupos = _riscoGruposNovos();
  if (!grupos.length) return '';
  const produtos = _riscoProdutos();
  const memTec = _riscoTecidos();
  const tecOpts = sel => '<option value="">— escolher —</option>' + (STATE.tecidos || [])
    .map(t => `<option value="${esc(t.id)}" ${t.id === sel ? 'selected' : ''}>${esc(t.nome)}</option>`).join('');

  return grupos.map((G, gi) => {
    const modelo = (G.itens[0].L.modelo || '').trim();
    const d = _riscoNovaDraft(G);
    const nomeTam = _riscoNomeTamanhos(G.tamanhos);
    const tamTxt = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3']
      .filter(k => (parseInt(G.tamanhos[k], 10) || 0) > 0)
      .map(k => `${k.toUpperCase()}=${G.tamanhos[k]}`).join(' · ');
    const previstas = previsaoFases(d.tipoPeca, d.variacao, d.sku);
    const previstasTxt = previstas.length
      ? `Fases previstas para este produto: <b>${esc(previstas.join(' · '))}</b>.`
      : 'Este destino ainda não tem previsão de fases — as linhas abaixo saíram dos PDFs.';
    const forcada = G.itens.some(it => it.L.forcarNova && it.L.grades && it.L.grades.length);
    return `
    <div class="card" style="margin-top:12px;border-left:3px solid var(--accent);">
      <div class="card-title">Grade nova — ${esc(modelo) || 'produto sem nome no risco'}</div>
      <div class="field-hint" style="margin-bottom:8px;">
        ${G.itens.length} risco(s) com os mesmos tamanhos (<b>${esc(tamTxt)}</b>)${forcada
          ? ' — há grade(s) com essa distribuição, mas você pediu para <b>criar uma nova</b>.'
          : ' e nenhuma grade cadastrada com essa distribuição.'}
        Os campos abaixo são os que o PDF <b>não</b> traz — eles são decisão da casa. Preenchidos uma vez, ficam guardados para este produto.
      </div>
      <div class="form-grid cols-3">
        <div class="field">
          <label>SKU da grade *</label>
          <input type="text" id="rn-sku-${gi}" value="${esc(d.sku)}" placeholder="Ex.: BM.TRICOLOR" oninput="riscoAtualizarNome(${gi})">
          <div class="field-hint">O que vem depois do "|" no nome.</div>
        </div>
        <div class="field">
          <label>Tipo de peça (pasta) *</label>
          <select id="rn-tipo-${gi}" data-prev="${esc(d.tipoPeca)}" onchange="riscoNovaMudouDestino(${gi}, this, 'pasta')">
            ${opcoesPastaGrade('pasta', d.tipoPeca)}
          </select>
          <div class="field-hint">Decide a <b>pasta</b> da grade — não o SKU — e a <b>previsão de fases</b> abaixo.</div>
        </div>
        <div class="field">
          <label>Variação (subpasta)</label>
          <select id="rn-var-${gi}" data-prev="${esc(d.variacao)}" onchange="riscoNovaMudouDestino(${gi}, this, 'subpasta')">
            ${opcoesPastaGrade('subpasta', d.variacao)}
          </select>
        </div>
        <div class="field">
          <label>Peças por pacote</label>
          <input type="number" min="0" id="rn-pac-${gi}" value="${esc(d.pecasPorPacote)}" placeholder="0">
        </div>
        <div class="field full">
          <label>Nome da grade</label>
          <input type="text" id="rn-nome-${gi}" value="${esc(nomeTam + (d.sku ? ' | ' + d.sku : ''))}" readonly class="is-auto">
          <div class="field-hint">Montado dos tamanhos do risco + o SKU. Os tamanhos vêm do PDF e não se digitam.</div>
        </div>
      </div>
      <div class="field-hint" style="margin:10px 0 4px;">
        ${previstasTxt}
        Fase <b>sem risco</b> nasce sem medida — entra depois, por outro PDF ou à mão. Dá para renomear, tirar e acrescentar.
      </div>
      <table class="table" style="font-size:12px;margin-top:4px;">
        <thead><tr><th style="width:26px;">#</th><th>Nome da fase *</th><th>Tecido *</th><th>Unid.</th><th>Medida (risco + excedente do tecido)</th><th>Arquivo</th><th style="width:28px;"></th></tr></thead>
        <tbody>
          ${d.fases.map((f, fi) => {
            const L = f.iPdf != null && G.itens[f.iPdf] ? G.itens[f.iPdf].L : null;
            const medida = L
              ? `${esc((_riscoCompCadastro(L.comprimento, f.tecidoId) || 0).toFixed(2))} × ${esc((L.largura || 0).toFixed(3))}
                 <div style="color:var(--ink-3);font-size:10px;">tecido no risco: ${esc(L.tecido || '—')}</div>`
              : `<span style="color:var(--ink-3);">— sem risco —</span>`;
            return `<tr>
              <td style="text-align:center;font-family:'IBM Plex Mono',monospace;">${fi + 1}</td>
              <td><input type="text" id="rn-f-nome-${gi}-${fi}" value="${esc(f.nome)}" placeholder="Ex.: Corpo Parte 1" style="font-size:12px;"></td>
              <td><select id="rn-f-tec-${gi}-${fi}" style="font-size:12px;">${tecOpts(f.tecidoId)}</select></td>
              <td><input type="number" min="1" id="rn-f-un-${gi}-${fi}" value="${esc(f.unidades)}" style="width:56px;font-size:12px;"></td>
              <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;">${medida}</td>
              <td style="font-size:10px;color:var(--ink-3);">${L ? esc(L.arquivo) : ''}</td>
              <td><button class="del" onclick="riscoNovaDelFase(${gi}, ${fi})" title="Tirar esta fase">✕</button></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div style="margin-top:8px;display:flex;gap:8px;align-items:center;">
        <button class="btn primary" onclick="criarGradeDoRisco(${gi})">+ Criar esta grade com ${d.fases.length} fase(s)</button>
        <button class="btn" onclick="riscoNovaAddFase(${gi})">+ Acrescentar fase</button>
      </div>
    </div>`;
  }).join('');
}

function riscoAtualizarNome(gi) {
  const G = _riscoGruposNovos()[gi];
  if (!G) return;
  const sku = (document.getElementById(`rn-sku-${gi}`)?.value || '').trim();
  const el = document.getElementById(`rn-nome-${gi}`);
  if (el) el.value = _riscoNomeTamanhos(G.tamanhos) + (sku ? ' | ' + sku : '');
}

async function criarGradeDoRisco(gi) {
  if (!exigirEdicao('criar grade pelo risco')) return;
  _riscoNovaColetar();
  const G = _riscoGruposNovos()[gi];
  if (!G) return;
  const d = _riscoNovas[G.assinatura];
  if (!d) return;
  const v = id => (document.getElementById(id)?.value || '').trim();
  const sku = d.sku;
  if (!sku) return toast('Informe o SKU da grade', 'err');
  const nome = v(`rn-nome-${gi}`);
  if ((STATE.grades || []).some(g => _normNome(g.nome) === _normNome(nome))) {
    return toast(`Já existe uma grade chamada "${nome}"`, 'err');
  }
  // As linhas são as da TABELA, não os PDFs: a grade nasce com as fases
  // previstas para o produto, tenha risco para todas ou não.
  const linhas = d.fases.map(f => ({
    L: (f.iPdf != null && G.itens[f.iPdf]) ? G.itens[f.iPdf].L : null,
    nome: f.nome, tecidoId: f.tecidoId, unidades: f.unidades
  }));
  if (!linhas.length) return toast('A grade precisa de pelo menos uma fase', 'err');
  const semNome = linhas.filter(x => !x.nome);
  if (semNome.length) return toast(`${semNome.length} fase(s) sem nome`, 'err');
  const semTec = linhas.filter(x => !x.tecidoId);
  if (semTec.length && !confirm(`${semTec.length} fase(s) sem tecido escolhido.\n\nFase sem tecido não entra no cálculo de consumo nem na baixa de estoque. Criar mesmo assim?`)) return;
  const semRisco = linhas.filter(x => !x.L);
  if (semRisco.length && !confirm(`${semRisco.length} fase(s) sem risco: ${semRisco.map(x => x.nome).join(', ')}.\n\nElas entram no cadastro SEM medida, para serem preenchidas depois — por outro PDF ou à mão. Criar assim?`)) return;

  const tamanhos = {};
  ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].forEach(k => { tamanhos[k] = parseInt(G.tamanhos[k], 10) || 0; });
  tamanhos.total = Object.keys(tamanhos).reduce((s, k) => s + (k === 'total' ? 0 : tamanhos[k]), 0);
  tamanhos.descricao = nome;

  const nova = {
    id: uid(), nome, tamanhos,
    tipoPeca: v(`rn-tipo-${gi}`), variacao: v(`rn-var-${gi}`),
    pecasPorPacote: parseInt(v(`rn-pac-${gi}`), 10) || 0,
    fases: linhas.map((x, i) => ({
      ordem: i + 1, nome: x.nome, tecidoId: x.tecidoId, unidades: x.unidades,
      comp: _riscoCompCadastro(x.L.comprimento, x.tecidoId).toFixed(2),
      larg: x.L.largura.toFixed(3),
      // Bobinas previstas o relatório não traz. Em malha algodão a regra da casa
      // dá o número pelo comprimento; nos outros tecidos fica em branco, para
      // ser preenchida quando alguém souber.
      bobinas: sugestaoBobinasFase(x.tecidoId, _riscoCompCadastro(x.L.comprimento, x.tecidoId)) ?? ''
    }))
  };
  // Campos legados (espelho da 1ª fase), como o cadastro manual grava.
  nova.enfestoComprimento = nova.fases[0] ? nova.fases[0].comp : '';
  nova.enfestoLargura = nova.fases[0] ? nova.fases[0].larg : '';

  if (!confirm(`Criar a grade "${nome}" com ${nova.fases.length} fase(s)?\n\n`
    + nova.fases.map(f => `  ${f.ordem}. ${f.nome}: ${f.comp} × ${f.larg}`).join('\n'))) return;

  STATE.grades.push(nova);
  // APRENDE: o produto e, por fase, o código do tecido. A próxima grade deste
  // mesmo produto já vem com tudo preenchido.
  const produtos = _riscoProdutos(), memFase = _riscoAprendidos(), memTec = _riscoTecidos();
  const modeloKey = _normNome(G.itens[0].L.modelo || '');
  if (modeloKey) {
    produtos[modeloKey] = { sku, tipoPeca: nova.tipoPeca, variacao: nova.variacao, pecasPorPacote: nova.pecasPorPacote };
  }
  linhas.forEach(x => {
    const k = _riscoChave(x.L);
    if (!k) return;
    memFase[k] = x.nome;
    if (x.tecidoId) memTec[k] = x.tecidoId;
  });
  await saveState('grades');
  await saveState('meta');
  toast(`Grade "${nome}" criada com ${nova.fases.length} fase(s)`, 'ok');
  // Relê os mesmos PDFs: agora eles encontram a grade recém-criada.
  _riscoLeituras.forEach(L => {
    if (L.erro) return;
    L.grades = _riscoGradesQueCasam(L.tamanhos);
    L.grade = L.grades.length === 1 ? L.grades[0] : null;
    L.res = L.grade ? _riscoResolverFase(L, L.grade) : { fase: null, origem: L.grades.length ? 'escolher grade' : 'sem grade' };
    L.aplicar = false;   // já foi gravado na criação
  });
  renderRiscoResultado();
  renderGrades();
}

/* ---------------- tela de importação do risco ---------------- */

function abrirModalRisco() {
  if (!exigirEdicao('importar risco de PDF')) return;
  _riscoLeituras = [];
  document.getElementById('modal-risco-fields').innerHTML = `
    <div class="info-box">
      Escolha os <b>relatórios de encaixe</b> gerados pelo CAD (um PDF por fase).
      O programa lê o comprimento, a largura, a tabela de tamanhos e o código do tecido de cada um, soma ao comprimento o <b>excedente de enfesto cadastrado no tecido</b> (${EXCEDENTE_ENFESTO_PADRAO_CM} cm nos que não têm) — a diferença entre a medida de <b>cortar</b>, que é a do relatório, e a de <b>enfestar</b>, que é a que se cadastra; a largura não recebe nada —,
      descobre <b>a qual grade</b> pertencem (pelos tamanhos) e <b>a qual fase</b> (pelo código do tecido,
      pela medida, ou pelo nome do arquivo — nessa ordem). Nada é gravado antes de você conferir.
    </div>
    <label class="file-label" style="font-size:13px;font-weight:500;">
      📄 Escolher os PDFs do encaixe
      <input type="file" accept="application/pdf" multiple onchange="lerRiscosEscolhidos(event)">
    </label>
    <div id="risco-resultado" style="margin-top:12px;"></div>`;
  const btn = document.getElementById('btn-risco-aplicar');
  if (btn) btn.style.display = 'none';
  openModal('modal-risco');
}

async function lerRiscosEscolhidos(ev) {
  const files = Array.from(ev.target.files || []);
  if (!files.length) return;
  const box = document.getElementById('risco-resultado');
  box.innerHTML = `<div class="empty" style="padding:16px;">Lendo ${files.length} arquivo(s)…</div>`;
  _riscoLeituras = [];
  for (const f of files) {
    try {
      const L = await _riscoLerPdf(f);
      L.grades = _riscoGradesQueCasam(L.tamanhos);
      L.grade = L.grades.length === 1 ? L.grades[0] : null;
      L.res = L.grade ? _riscoResolverFase(L, L.grade) : { fase: null, origem: L.grades.length ? 'escolher grade' : 'sem grade' };
      L.aplicar = !!(L.grade && L.res.fase && L.comprimento && L.largura);
      _riscoLeituras.push(L);
    } catch (e) {
      _riscoLeituras.push({ arquivo: f.name, erro: e.message || String(e) });
    }
  }
  renderRiscoResultado();
}

// A célula da GRADE de uma linha. Três situações, e a do meio era um beco sem
// saída até aqui:
//   • uma só grade com aqueles tamanhos → mostra o nome, e pronto;
//   • VÁRIAS grades → agora um seletor para escolher em qual lançar. Antes a
//     linha dizia "7 grades com estes tamanhos" e a fase dizia "grade não
//     encontrada", com o quadradinho desabilitado e sem botão de aplicar: o
//     risco não tinha como virar cadastro por esta tela. E o bloco de "grade
//     nova" também não aparecia, porque grade candidata existia. Acontece em
//     todo risco de P ao G3, que é a grade mais comum da casa — sete cadastros
//     têm exatamente esses tamanhos (CM.BÁSICA, CM.TRICOLOR, BM.BÁSICA,
//     PM.BÁSICA, PM.TRICOLOR, SM.LISO, SM. ESPARTANA).
//     O assistente de PASTA já resolvia assim; esta tela ficou para trás.
//   • nenhuma grade → o aviso de sempre, e o bloco de grade nova abaixo.
function _riscoCelulaGrade(L, i) {
  const n = (L.grades || []).length;
  if (n === 1 && L.grade) return esc(L.grade.nome);
  if (n > 1) {
    const opts = '<option value="">— escolher entre ' + n + ' —</option>' + L.grades
      .map(g => `<option value="${esc(g.id)}" ${L.grade && L.grade.id === g.id ? 'selected' : ''}>${esc(g.nome)}</option>`)
      .join('');
    return `<select onchange="riscoTrocarGrade(${i}, this.value)" style="font-size:12px;max-width:100%;">${opts}</select>`;
  }
  return '<span style="color:var(--alert);">nenhuma grade com estes tamanhos</span>';
}

function renderRiscoResultado() {
  const box = document.getElementById('risco-resultado');
  const btn = document.getElementById('btn-risco-aplicar');
  if (!box) return;
  const fmt = v => v == null ? '—' : String(v).replace('.', ',');
  const linhas = _riscoLeituras.map((L, i) => {
    if (L.erro) {
      return `<tr><td colspan="8" style="color:var(--alert);"><b>${esc(L.arquivo)}</b> — ${esc(L.erro)}</td></tr>`;
    }
    // De onde veio a decisão da fase: é o que o usuário precisa julgar.
    const selo = {
      'aprendido': '<span class="exp-badge ok" title="O par modelo + código do tecido já foi ensinado numa importação anterior.">aprendido</span>',
      'medida': `<span class="exp-badge ok" title="Largura dá a família do tecido; o comprimento mais próximo dá a fase. Diferença de ${fmt((L.res.dist || 0).toFixed(2))} m contra ${fmt((L.res.folga || 0).toFixed(2))} m da segunda opção.">pela medida</span>`,
      'nome do arquivo': '<span class="exp-badge baixo" title="O conteúdo não decidiu — valeu o nome do arquivo. Confira.">pelo nome</span>',
      'largura': '<span class="exp-badge baixo" title="Só a largura isolou esta fase.">pela largura</span>',
      'indefinida': '<span class="exp-badge alto">não identificada</span>',
      'escolher grade': '<span class="exp-badge baixo" title="Mais de uma grade tem exatamente estes tamanhos. Escolha ao lado em qual lançar.">escolha a grade</span>',
      'sem grade': '<span class="exp-badge alto">grade não encontrada</span>',
      'sem fases': '<span class="exp-badge alto">a grade não tem fases</span>'
    }[L.res.origem] || `<span class="exp-badge alto">${esc(L.res.origem)}</span>`;

    const fases = L.grade ? (L.grade.fases || []).slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0)) : [];
    const selFase = L.grade ? `<select onchange="riscoTrocarFase(${i}, this.value)" style="font-size:12px;">
        <option value="">— escolher —</option>
        ${fases.map(f => `<option value="${esc(f.nome)}" ${L.res.fase && f.nome === L.res.fase.nome ? 'selected' : ''}>${esc(f.nome)}</option>`).join('')}
      </select>` : '—';

    const atual = L.res.fase ? `${fmt(L.res.fase.comp || '—')} × ${fmt(L.res.fase.larg || '—')}` : '—';
    const compCad = _riscoCompCadastro(L.comprimento, (L.res.fase || {}).tecidoId);
    const novo = `${fmt(compCad != null ? compCad.toFixed(2) : null)} × ${fmt(L.largura != null ? L.largura.toFixed(3) : null)}`;
    const mudou = L.res.fase && (_riscoF(L.res.fase.comp) !== (compCad == null ? null : +compCad.toFixed(2))
      || _riscoF(L.res.fase.larg) !== (L.largura == null ? null : +L.largura.toFixed(3)));
    return `<tr>
      <td style="text-align:center;"><input type="checkbox" ${L.aplicar ? 'checked' : ''} ${L.grade && L.res.fase ? '' : 'disabled'} onchange="riscoMarcar(${i}, this.checked)"></td>
      <td style="font-size:11px;">${esc(L.arquivo)}</td>
      <td style="font-size:12px;">${_riscoCelulaGrade(L, i)}</td>
      <td>${selFase}<div style="margin-top:2px;">${selo}</div></td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;">${esc(L.tecido || '—')}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-3);">${esc(atual)}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;${mudou ? 'color:var(--alert);' : ''}">${esc(novo)}
        <div style="font-weight:400;color:var(--ink-3);font-size:10px;">risco ${esc(fmt(L.comprimento != null ? L.comprimento.toFixed(2) : '—'))} + ${excedenteEnfestoCm((L.res.fase || {}).tecidoId)} cm</div></td>
      <td style="font-size:11px;">${L.gramatura ? esc(L.gramatura + ' g/m²') : ''}${L.aproveitamento ? ' · ' + esc(L.aproveitamento + '%') : ''}</td>
    </tr>`;
  }).join('');

  const nOk = _riscoLeituras.filter(L => L.aplicar).length;
  box.innerHTML = `
    <table class="table" style="font-size:12px;">
      <thead><tr>
        <th style="width:30px;"></th><th>Arquivo</th><th>Grade (pelos tamanhos)</th>
        <th>Fase</th><th>Tecido</th><th>Cadastro</th><th>Do PDF</th><th>Extra</th>
      </tr></thead>
      <tbody>${linhas}</tbody>
    </table>
    ${_riscoHtmlGradesNovas()}
    <div class="field-hint" style="margin-top:8px;">
      A coluna <b>Do PDF</b> em vermelho é o que vai <b>mudar</b> no cadastro. Desmarque o que não quiser aplicar.
      Ao aplicar, o programa <b>aprende</b> o par <i>modelo do risco + código do tecido</i> de cada linha —
      na próxima importação daquele produto a fase é reconhecida sozinha, sem depender do nome do arquivo.
    </div>`;
  if (btn) { btn.style.display = nOk ? '' : 'none'; btn.textContent = `Aplicar ${nOk} fase(s) nas grades`; }
}

// O checkbox da linha. Precisa ser função: `_riscoLeituras` é `let` no escopo do
// arquivo, e um `onchange` inline resolve nomes no `window`, onde ela não está.
function riscoMarcar(i, marcado) {
  if (_riscoLeituras[i]) _riscoLeituras[i].aplicar = !!marcado;
  const btn = document.getElementById('btn-risco-aplicar');
  const n = _riscoLeituras.filter(L => L.aplicar).length;
  if (btn) { btn.style.display = n ? '' : 'none'; btn.textContent = `Aplicar ${n} fase(s) nas grades`; }
}

// Escolher a grade quando várias têm os mesmos tamanhos. Trocar a grade refaz a
// escolha da FASE: as fases são de cada grade, e a que estava selecionada não
// existe na nova.
function riscoTrocarGrade(i, id) {
  const L = _riscoLeituras[i];
  if (!L) return;
  L.grade = (L.grades || []).find(g => g.id === id) || null;
  L.res = L.grade
    ? _riscoResolverFase(L, L.grade)
    : { fase: null, origem: (L.grades || []).length ? 'escolher grade' : 'sem grade' };
  L.aplicar = !!(L.grade && L.res.fase);
  renderRiscoResultado();
}

function riscoTrocarFase(i, nome) {
  const L = _riscoLeituras[i];
  if (!L || !L.grade) return;
  const f = (L.grade.fases || []).find(x => x.nome === nome);
  L.res = { fase: f || null, origem: f ? 'escolhida por você' : 'indefinida', folga: null };
  L.aplicar = !!f;
  renderRiscoResultado();
}

async function aplicarRiscoNasGrades() {
  if (!exigirEdicao('importar risco de PDF')) return;
  const alvo = _riscoLeituras.filter(L => L.aplicar && L.grade && L.res.fase);
  if (!alvo.length) return toast('Nada marcado para aplicar', 'err');
  const mudancas = alvo.map(L =>
    `${L.grade.nome} · ${L.res.fase.nome}: ${L.res.fase.comp || '—'}×${L.res.fase.larg || '—'} → ${_riscoCompCadastro(L.comprimento, L.res.fase.tecidoId).toFixed(2)}×${L.largura.toFixed(3)}`);
  if (!confirm(`Aplicar ${alvo.length} medida(s) no cadastro das grades?\n\n${mudancas.join('\n')}\n\n`
    + 'O programa também vai guardar a que fase corresponde cada código de tecido, para reconhecer sozinho na próxima vez.')) return;

  const memoria = _riscoAprendidos();
  let n = 0;
  alvo.forEach(L => {
    const f = (L.grade.fases || []).find(x => x.nome === L.res.fase.nome);
    if (!f) return;
    f.comp = _riscoCompCadastro(L.comprimento, f.tecidoId).toFixed(2);   // + o excedente do TECIDO da fase
    f.larg = L.largura.toFixed(3);                            // largura vai como veio
    const k = _riscoChave(L);
    if (k) memoria[k] = f.nome;
    n++;
  });
  // Campo legado da grade (espelho da 1ª fase), para não ficar divergindo.
  new Set(alvo.map(L => L.grade)).forEach(g => {
    const f1 = (g.fases || []).slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0))[0];
    if (f1) { g.enfestoComprimento = f1.comp || ''; g.enfestoLargura = f1.larg || ''; }
  });
  await saveState('grades');
  await saveState('meta');
  closeModal('modal-risco');
  toast(`${n} fase(s) atualizada(s) pelo risco · ${Object.keys(memoria).length} vínculo(s) de tecido aprendidos`, 'ok');
  renderGrades();
}

/* ========================================================= */
/*   ASSISTENTE: uma PASTA inteira de riscos, uma por vez    */
/* ========================================================= */
// A tela de cima ("Importar risco") resolve um punhado de PDFs de uma vez, mas
// obriga a escolher arquivo por arquivo e a decidir tudo numa tabela só. Quando
// chega a pasta inteira do CAD — dezenas de relatórios, de vários produtos — o
// que se quer é outra coisa: apontar A PASTA e ser conduzido, uma grade de cada
// vez, até não sobrar PDF.
//
// O assistente faz isso em três tempos:
//   1. LÊ tudo. Varre a pasta (e as subpastas), lê cada relatório e mostra o
//      progresso. PDF que não é relatório do CAD fica de lado, com o motivo.
//   2. AGRUPA por distribuição de tamanhos. Os cinco riscos da canguru têm os
//      mesmos tamanhos: são UMA grade de cinco fases, não cinco cadastros. Cada
//      grupo já nasce sabendo se é grade NOVA (nenhuma cadastrada com aqueles
//      tamanhos) ou CORREÇÃO de uma existente.
//   3. PERGUNTA, grupo a grupo, só o que o PDF não traz — nome, tecido, unidades
//      e o consumo previsto em bobinas. Salva e passa para o próximo. Termina
//      quando o último grupo foi salvo ou pulado, com um resumo do que mudou.
//
// Tudo o que é respondido aqui vira memória (produto, fase, tecido, bobinas): a
// próxima pasta do mesmo produto chega quase toda preenchida.

let _pastaWiz = null;

// A memória do CONSUMO: "modelo do risco | código do tecido" → bobinas previstas.
// É o único dos quatro campos decisivos que ninguém consegue deduzir do desenho —
// vem do que a casa gastou nas últimas produções — e por isso vale guardar.
function _riscoBobinasMem() {
  STATE.meta = STATE.meta || {};
  if (!STATE.meta.riscoBobinas || typeof STATE.meta.riscoBobinas !== 'object') STATE.meta.riscoBobinas = {};
  return STATE.meta.riscoBobinas;
}

// Escolher a pasta. O Chrome e o Edge no desktop abrem o seletor nativo de
// pastas e deixam varrer as subpastas; nos outros sobra o <input webkitdirectory>,
// que entrega a mesma lista de arquivos por outro caminho.
async function _escolherPastaDeRiscos() {
  if ('showDirectoryPicker' in window) {
    const dir = await window.showDirectoryPicker({ mode: 'read', id: 'riscos-cad' });
    const arquivos = [];
    const anda = async (h, prefixo) => {
      for await (const [nome, ent] of h.entries()) {
        if (ent.kind === 'directory') { await anda(ent, prefixo + nome + '/'); continue; }
        if (!/\.pdf$/i.test(nome)) continue;
        arquivos.push({ file: await ent.getFile(), caminho: prefixo + nome });
      }
    };
    await anda(dir, '');
    return { pasta: dir.name, arquivos };
  }
  return new Promise(resolve => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.multiple = true;
    inp.webkitdirectory = true;
    inp.style.display = 'none';
    document.body.appendChild(inp);
    inp.onchange = () => {
      const fs = Array.from(inp.files || []).filter(f => /\.pdf$/i.test(f.name));
      const raiz = (fs[0]?.webkitRelativePath || '').split('/')[0] || 'pasta escolhida';
      inp.remove();
      resolve({ pasta: raiz, arquivos: fs.map(f => ({ file: f, caminho: f.webkitRelativePath || f.name })) });
    };
    // Cancelar o seletor não dispara evento nenhum: o Promise fica pendente e o
    // input morre com a página. Não trava nada — a tela nem chegou a abrir.
    inp.click();
  });
}

async function abrirAssistentePasta() {
  if (!exigirEdicao('importar uma pasta de riscos')) return;
  let escolha = null;
  try {
    escolha = await _escolherPastaDeRiscos();
  } catch (e) {
    if (e && (e.name === 'AbortError' || e.name === 'NotAllowedError')) return;   // desistiu
    return toast('Não foi possível abrir a pasta: ' + (e.message || e), 'err');
  }
  if (!escolha) return;
  if (!escolha.arquivos.length) return toast('Nenhum PDF nesta pasta', 'err');
  escolha.arquivos.sort((a, b) => a.caminho.localeCompare(b.caminho, 'pt-BR', { numeric: true }));

  _pastaWiz = {
    pasta: escolha.pasta, arquivos: escolha.arquivos, leituras: [],
    grupos: [], idx: 0, etapa: 'lendo', lidos: 0,
    feito: { criadas: [], atualizadas: [], fases: 0 }
  };
  openModal('modal-risco-pasta');
  renderPastaWiz();

  // `W` é a identidade desta importação. Se quem está usando fechar a janela no
  // meio da leitura (ou abrir outra pasta), `_pastaWiz` deixa de ser `W` e o
  // laço para — sem isso ele continuaria empurrando leituras num objeto morto.
  const W = _pastaWiz;
  for (const it of escolha.arquivos) {
    try {
      const L = await _riscoLerPdf(it.file);
      L.caminho = it.caminho;
      W.leituras.push(L);
    } catch (e) {
      W.leituras.push({ arquivo: it.file.name, caminho: it.caminho, erro: e.message || String(e) });
    }
    if (_pastaWiz !== W) return;
    W.lidos++;
    renderPastaWiz();
    await new Promise(r => setTimeout(r, 0));   // deixa a tela respirar
    if (_pastaWiz !== W) return;
  }

  _pastaMontarGrupos();
  _pastaWiz.etapa = _pastaWiz.grupos.length ? 'passo' : 'fim';
  renderPastaWiz();
}

// Um grupo por distribuição de tamanhos: é o que define UMA grade.
function _pastaMontarGrupos() {
  const grupos = new Map();
  _pastaWiz.leituras.forEach(L => {
    if (L.erro) return;
    const a = _riscoAssinatura(L.tamanhos);
    if (a === '0-0-0-0-0-0-0') {
      L.erro = 'Não deu para ler a tabela de tamanhos deste relatório';
      return;
    }
    if (!grupos.has(a)) grupos.set(a, { assinatura: a, tamanhos: L.tamanhos, itens: [] });
    grupos.get(a).itens.push(L);
  });
  _pastaWiz.grupos = Array.from(grupos.values()).map(G => {
    G.itens.sort((x, y) => (x.caminho || '').localeCompare(y.caminho || '', 'pt-BR', { numeric: true }));
    G.candidatas = _riscoGradesQueCasam(G.tamanhos);
    G.gradeId = G.candidatas.length === 1 ? G.candidatas[0].id : '';
    G.modelo = (G.itens.find(L => L.modelo) || {}).modelo || '';
    G.status = 'pendente';
    _pastaIniciarRascunho(G);
    return G;
  });
}

// O rascunho do grupo: o que a tela mostra e o que o "Salvar" grava. Nasce do
// que já foi aprendido antes; o que não foi, nasce vazio para ser respondido.
function _pastaIniciarRascunho(G) {
  const prod = _riscoProdutos()[_normNome(G.modelo)] || {};
  G.draft = {
    sku: prod.sku || '',
    tipoPeca: prod.tipoPeca || 'camiseta',
    variacao: prod.variacao || '',
    pecasPorPacote: prod.pecasPorPacote || '',
    fases: []
  };
  _pastaResetFases(G);
}

// As linhas de fase dependem do destino: numa grade existente cada PDF aponta
// para uma fase dela; numa grade nova, cada PDF é uma fase a criar.
function _pastaResetFases(G) {
  const memTec = _riscoTecidos(), memBob = _riscoBobinasMem();
  const grade = (STATE.grades || []).find(g => g.id === G.gradeId) || null;
  G.draft.fases = G.itens.map(L => {
    const k = _riscoChave(L);
    const res = grade ? _riscoResolverFase(L, grade) : { fase: null, origem: 'grade nova' };
    const f = res.fase;
    const bobMem = memBob[k];
    const d = {
      alvo: f ? f.nome : '__nova__',
      origem: res.origem,
      nome: f ? f.nome : _riscoNomeFaseSugerido(L),
      tecidoId: (f && f.tecidoId) || memTec[k] || '',
      unidades: (f && parseInt(f.unidades, 10)) || 2,
      bobinas: (f && f.bobinas !== '' && f.bobinas != null) ? String(f.bobinas).replace('.', ',')
             : (bobMem != null && bobMem !== '' ? String(bobMem).replace('.', ',') : ''),
      aplicar: !!(L.comprimento != null && L.largura != null)
    };
    _pastaSugerirBobinas(d, L);
    return d;
  });
}

// A regra da malha algodão preenchendo o rascunho. Só entra onde o campo está
// vazio — o que já estava no cadastro, ou o que foi respondido antes para este
// produto, vale mais do que a regra geral. `sugerida` marca o que foi ela que
// escreveu, para que trocar o tecido recalcule sem apagar o que foi digitado.
function _pastaSugerirBobinas(d, L) {
  const sug = sugestaoBobinasFase(d.tecidoId, _riscoCompCadastro(L.comprimento, d.tecidoId));
  d.regra = sug;
  if (sug == null) return;
  if (d.bobinas === '' || d.bobinas === d.sugerida) {
    d.bobinas = String(sug);
    d.sugerida = String(sug);
  }
}

// Lê a tela de volta para o rascunho. Chamado antes de qualquer redesenho e
// antes de salvar — sem isso, trocar o destino apagaria o que foi digitado.
function _pastaColetar() {
  const G = _pastaWiz && _pastaWiz.grupos[_pastaWiz.idx];
  if (!G || !G.draft) return;
  const v = id => (document.getElementById(id)?.value ?? '').trim();
  if (document.getElementById('pw-sku')) {
    G.draft.sku = v('pw-sku');
    G.draft.tipoPeca = v('pw-tipo');
    G.draft.variacao = v('pw-var');
    G.draft.pecasPorPacote = v('pw-pac');
  }
  G.draft.fases.forEach((d, fi) => {
    const sel = document.getElementById(`pw-f-alvo-${fi}`);
    if (sel) d.alvo = sel.value;
    const nome = document.getElementById(`pw-f-nome-${fi}`);
    if (nome) d.nome = nome.value.trim();
    const tec = document.getElementById(`pw-f-tec-${fi}`);
    if (tec) d.tecidoId = tec.value;
    const un = document.getElementById(`pw-f-un-${fi}`);
    if (un) d.unidades = parseInt(un.value, 10) || 2;
    const bob = document.getElementById(`pw-f-bob-${fi}`);
    if (bob) d.bobinas = bob.value.trim();
    const ap = document.getElementById(`pw-f-ap-${fi}`);
    if (ap) d.aplicar = ap.checked;
  });
}

function renderPastaWiz() {
  const box = document.getElementById('modal-risco-pasta-fields');
  if (!box || !_pastaWiz) return;
  if (_pastaWiz.etapa === 'lendo') return void (box.innerHTML = _pastaHtmlLendo());
  if (_pastaWiz.etapa === 'fim') return void (box.innerHTML = _pastaHtmlResumo());
  box.innerHTML = _pastaHtmlPasso();
}

function _pastaHtmlLendo() {
  const tot = _pastaWiz.arquivos.length, n = _pastaWiz.lidos;
  const pct = tot ? Math.round(n / tot * 100) : 0;
  const erros = _pastaWiz.leituras.filter(L => L.erro).length;
  return `
    <div class="info-box">Pasta <b>${esc(_pastaWiz.pasta)}</b> — ${tot} arquivo(s) PDF.</div>
    <div style="margin:14px 0;">
      <div style="height:8px;background:var(--line-2);border-radius:4px;overflow:hidden;">
        <div style="height:100%;width:${pct}%;background:var(--accent);transition:width .15s;"></div>
      </div>
      <div style="margin-top:6px;font-size:13px;color:var(--ink-2);">
        Lendo ${n} de ${tot}…${erros ? ` · ${erros} sem leitura` : ''}
      </div>
    </div>`;
}

function _pastaHtmlPasso() {
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  const tot = _pastaWiz.grupos.length, pos = _pastaWiz.idx + 1;
  const grade = (STATE.grades || []).find(g => g.id === G.gradeId) || null;
  const nomeTam = _riscoNomeTamanhos(G.tamanhos);
  const tamTxt = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3']
    .filter(k => (parseInt(G.tamanhos[k], 10) || 0) > 0)
    .map(k => `${k.toUpperCase()}=${G.tamanhos[k]}`).join(' · ');

  const casaId = new Set(G.candidatas.map(g => g.id));
  const opts = '<option value="">➕ Criar uma grade NOVA com estes tamanhos</option>'
    + (STATE.grades || []).slice().sort((a, b) => (a.nome || '').localeCompare(b.nome || '', 'pt-BR'))
        .map(g => `<option value="${esc(g.id)}" ${g.id === G.gradeId ? 'selected' : ''}>${casaId.has(g.id) ? '✓ ' : ''}${esc(g.nome)}</option>`).join('');

  const dica = grade
    ? (casaId.has(grade.id)
        ? 'Os tamanhos deste risco batem com esta grade. As medidas abaixo vão <b>corrigir</b> o cadastro dela.'
        : '<b>Atenção:</b> os tamanhos deste risco não batem com a distribuição desta grade.')
    : (G.candidatas.length > 1
        ? `<b>${G.candidatas.length} grades</b> têm exatamente estes tamanhos — escolha acima em qual lançar, ou crie uma nova.`
        : 'Nenhuma grade cadastrada tem estes tamanhos. Responda abaixo o que o PDF não traz e ela será criada.');

  const produto = grade ? '' : `
    <div class="form-grid cols-3" style="margin-top:10px;">
      <div class="field">
        <label>SKU da grade *</label>
        <input type="text" id="pw-sku" value="${esc(G.draft.sku)}" placeholder="Ex.: BM.TRICOLOR" oninput="pastaAtualizarNome()">
        <div class="field-hint">O que vem depois do "|" no nome.</div>
      </div>
      <div class="field">
        <label>Tipo de peça (pasta) *</label>
        <select id="pw-tipo" data-prev="${esc(G.draft.tipoPeca || '')}" onchange="onSelectGradeFolder(this,'pasta')">
          ${opcoesPastaGrade('pasta', G.draft.tipoPeca || '')}
        </select>
        <div class="field-hint">Decide a <b>pasta</b> da grade — não o SKU.</div>
      </div>
      <div class="field">
        <label>Variação (subpasta)</label>
        <select id="pw-var" data-prev="${esc(G.draft.variacao || '')}" onchange="onSelectGradeFolder(this,'subpasta')">
          ${opcoesPastaGrade('subpasta', G.draft.variacao || '')}
        </select>
      </div>
      <div class="field">
        <label>Peças por pacote</label>
        <input type="number" min="0" id="pw-pac" value="${esc(G.draft.pecasPorPacote)}" placeholder="0">
      </div>
      <div class="field full">
        <label>Nome da grade</label>
        <input type="text" id="pw-nome" value="${esc(nomeTam + (G.draft.sku ? ' | ' + G.draft.sku : ''))}" readonly class="is-auto">
        <div class="field-hint">Montado dos tamanhos do risco + o SKU. Os tamanhos vêm do PDF e não se digitam.</div>
      </div>
    </div>`;

  const fasesGrade = grade ? (grade.fases || []).slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0)) : [];
  const tecOpts = sel => '<option value="">— escolher —</option>' + (STATE.tecidos || [])
    .map(t => `<option value="${esc(t.id)}" ${t.id === sel ? 'selected' : ''}>${esc(t.nome)}</option>`).join('');

  const linhas = G.itens.map((L, fi) => {
    const d = G.draft.fases[fi];
    const compCad = _riscoCompCadastro(L.comprimento, d.tecidoId);
    const semMedida = compCad == null || L.largura == null;
    const alvoFase = grade ? fasesGrade.find(f => f.nome === d.alvo) : null;
    const atual = alvoFase ? `${esc(String(alvoFase.comp || '—').replace('.', ','))} × ${esc(String(alvoFase.larg || '—').replace('.', ','))}` : '—';
    const nova = semMedida ? '—' : `${compCad.toFixed(2).replace('.', ',')} × ${L.largura.toFixed(3).replace('.', ',')}`;
    const mudou = alvoFase && !semMedida &&
      (_riscoF(alvoFase.comp) !== +compCad.toFixed(2) || _riscoF(alvoFase.larg) !== +L.largura.toFixed(3));

    const selAlvo = grade ? `<select id="pw-f-alvo-${fi}" onchange="pastaTrocarAlvo(${fi})" style="font-size:12px;width:100%;">
        ${fasesGrade.map(f => `<option value="${esc(f.nome)}" ${f.nome === d.alvo ? 'selected' : ''}>${esc(f.nome)}</option>`).join('')}
        <option value="__nova__" ${d.alvo === '__nova__' ? 'selected' : ''}>➕ criar fase nova</option>
      </select>
      <div style="font-size:10px;color:var(--ink-3);margin-top:2px;">${esc(d.origem)}</div>` : '';

    const campoNome = (!grade || d.alvo === '__nova__')
      ? `<input type="text" id="pw-f-nome-${fi}" value="${esc(d.nome)}" placeholder="Ex.: Corpo Parte 1" style="font-size:12px;width:100%;">`
      : '';

    return `<tr>
      <td style="text-align:center;"><input type="checkbox" id="pw-f-ap-${fi}" ${d.aplicar ? 'checked' : ''} ${semMedida ? 'disabled' : ''}></td>
      <td style="font-size:10px;color:var(--ink-3);max-width:150px;word-break:break-all;">${esc(L.caminho || L.arquivo)}
        <div style="font-family:'IBM Plex Mono',monospace;">tecido: ${esc(L.tecido || '—')}</div></td>
      ${grade ? `<td>${selAlvo}</td>` : ''}
      <td>${campoNome}</td>
      <td><select id="pw-f-tec-${fi}" onchange="pastaTecidoMudou()" style="font-size:12px;">${tecOpts(d.tecidoId)}</select></td>
      <td><input type="number" min="1" id="pw-f-un-${fi}" value="${esc(d.unidades)}" style="width:52px;font-size:12px;"></td>
      <td><input type="text" id="pw-f-bob-${fi}" value="${esc(d.bobinas)}" placeholder="14 · 1/2 · 0" style="width:76px;font-size:12px;">
        ${d.regra == null ? '' : `<div style="font-size:9px;color:var(--ink-3);" title="${esc(textoRegraBobinas(compCad))}">regra: ${d.regra}</div>`}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-3);">${atual}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;${mudou ? 'color:var(--alert);' : ''}${semMedida ? 'color:var(--alert);' : ''}">${semMedida ? 'sem medida' : esc(nova)}</td>
    </tr>`;
  }).join('');

  const pct = Math.round(_pastaWiz.idx / _pastaWiz.grupos.length * 100);
  return `
    <div style="display:flex;align-items:baseline;gap:10px;">
      <div style="font-weight:700;">Grade ${pos} de ${tot}</div>
      <div style="font-size:12px;color:var(--ink-3);">pasta ${esc(_pastaWiz.pasta)} · ${G.itens.length} risco(s) com os mesmos tamanhos</div>
    </div>
    <div style="height:6px;background:var(--line-2);border-radius:3px;overflow:hidden;margin:6px 0 12px;">
      <div style="height:100%;width:${pct}%;background:var(--accent);"></div>
    </div>

    <div class="field">
      <label>Onde lançar estes riscos</label>
      <select id="pw-destino" onchange="pastaTrocarDestino(this.value)">${opts}</select>
      <div class="field-hint">
        Produto no risco: <b>${esc(G.modelo || '—')}</b> · tamanhos <b>${esc(tamTxt)}</b> (${esc(nomeTam)}). ${dica}
      </div>
    </div>
    ${produto}

    <table class="table" style="font-size:12px;margin-top:10px;">
      <thead><tr>
        <th style="width:26px;"></th><th>Arquivo</th>
        ${grade ? '<th style="width:160px;">Fase do cadastro</th>' : ''}
        <th>Nome da fase${grade ? '' : ' *'}</th><th style="width:150px;">Tecido *</th>
        <th style="width:60px;" title="Quantas peças por camada esta fase rende (ribana)">Unid.</th>
        <th style="width:88px;" title="Consumo previsto: quantas bobinas deste tecido a grade gasta nesta fase. Aparece na coluna Consumo da folha de OS.">Bobinas</th>
        <th>Cadastro</th><th>Do risco + excedente</th>
      </tr></thead>
      <tbody>${linhas}</tbody>
    </table>

    <div style="display:flex;gap:8px;margin-top:14px;align-items:center;">
      <button class="btn primary" onclick="pastaSalvarPasso()">${grade ? 'Aplicar e continuar' : 'Criar grade e continuar'} →</button>
      <button class="btn" onclick="pastaPularPasso()">Pular esta grade</button>
      <span style="flex:1;"></span>
      <span style="font-size:12px;color:var(--ink-3);">${tot - pos} grade(s) depois desta</span>
    </div>`;
}

function _pastaHtmlResumo() {
  const F = _pastaWiz.feito;
  const erros = _pastaWiz.leituras.filter(L => L.erro);
  const pulados = _pastaWiz.grupos.filter(G => G.status === 'pulado');
  const li = (t, arr) => arr.length
    ? `<div style="margin-top:10px;"><b>${t}</b><ul style="margin:4px 0 0 18px;font-size:13px;">${arr.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
  return `
    <div class="info-box">
      Pasta <b>${esc(_pastaWiz.pasta)}</b> — ${_pastaWiz.arquivos.length} PDF(s) lidos,
      ${_pastaWiz.grupos.length} grade(s) identificada(s).
      <b>${F.criadas.length}</b> criada(s), <b>${F.atualizadas.length}</b> atualizada(s), <b>${F.fases}</b> fase(s) gravada(s).
    </div>
    ${li('Grades criadas', F.criadas)}
    ${li('Grades atualizadas', F.atualizadas)}
    ${li('Puladas por você', pulados.map(G => `${G.modelo || 'sem nome no risco'} — ${_riscoNomeTamanhos(G.tamanhos)} (${G.itens.length} risco)`))}
    ${li('Arquivos sem leitura', erros.map(L => `${L.caminho || L.arquivo}: ${L.erro}`))}
    <div style="margin-top:16px;">
      <button class="btn primary" onclick="pastaFechar()">Concluir</button>
    </div>`;
}

function pastaAtualizarNome() {
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  const sku = (document.getElementById('pw-sku')?.value || '').trim();
  const el = document.getElementById('pw-nome');
  if (el) el.value = _riscoNomeTamanhos(G.tamanhos) + (sku ? ' | ' + sku : '');
}

function pastaTrocarDestino(valor) {
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  _pastaColetar();
  G.gradeId = valor;
  _pastaResetFases(G);      // as fases só fazem sentido em relação ao destino
  renderPastaWiz();
}

function pastaTrocarAlvo(fi) {
  _pastaColetar();
  renderPastaWiz();          // "criar fase nova" abre o campo de nome
}

// Trocar o tecido muda a regra: malha algodão tem previsão de bobinas pelo
// comprimento, moletom e ribana não têm.
function pastaTecidoMudou() {
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  _pastaColetar();
  G.draft.fases.forEach((d, fi) => _pastaSugerirBobinas(d, G.itens[fi]));
  renderPastaWiz();
}

function pastaPularPasso() {
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  G.status = 'pulado';
  _pastaAvancar();
}

function _pastaAvancar() {
  _pastaWiz.idx++;
  if (_pastaWiz.idx >= _pastaWiz.grupos.length) _pastaWiz.etapa = 'fim';
  renderPastaWiz();
}

async function pastaSalvarPasso() {
  if (!exigirEdicao('cadastrar grade pelo risco')) return;
  _pastaColetar();
  const G = _pastaWiz.grupos[_pastaWiz.idx];
  const linhas = G.itens.map((L, fi) => ({ L, d: G.draft.fases[fi] })).filter(x => x.d.aplicar);
  if (!linhas.length) return toast('Nenhum risco marcado — use "Pular esta grade"', 'err');

  const grade = (STATE.grades || []).find(g => g.id === G.gradeId) || null;
  const memProd = _riscoProdutos(), memFase = _riscoAprendidos(),
        memTec = _riscoTecidos(), memBob = _riscoBobinasMem();
  const bob = s => { const n = parseBobinas(s); return n == null ? '' : n; };

  if (!grade) {
    /* ---- grade NOVA ---- */
    const sku = G.draft.sku;
    if (!sku) return toast('Informe o SKU da grade', 'err');
    const nome = _riscoNomeTamanhos(G.tamanhos) + ' | ' + sku;
    if ((STATE.grades || []).some(g => _normNome(g.nome) === _normNome(nome))) {
      return toast(`Já existe uma grade chamada "${nome}"`, 'err');
    }
    if (linhas.some(x => !x.d.nome)) return toast('Toda fase precisa de nome', 'err');
    const semTec = linhas.filter(x => !x.d.tecidoId).length;
    if (semTec && !confirm(`${semTec} fase(s) sem tecido escolhido.\n\nFase sem tecido não entra no cálculo de consumo nem na baixa de estoque. Criar mesmo assim?`)) return;

    const tamanhos = {};
    ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].forEach(k => { tamanhos[k] = parseInt(G.tamanhos[k], 10) || 0; });
    tamanhos.total = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'].reduce((s, k) => s + tamanhos[k], 0);
    tamanhos.descricao = nome;

    const nova = {
      id: uid(), nome, tamanhos,
      tipoPeca: G.draft.tipoPeca, variacao: G.draft.variacao,
      pecasPorPacote: parseInt(G.draft.pecasPorPacote, 10) || 0,
      fases: linhas.map((x, i) => ({
        ordem: i + 1, nome: x.d.nome, tecidoId: x.d.tecidoId, unidades: x.d.unidades,
        comp: _riscoCompCadastro(x.L.comprimento, x.d.tecidoId).toFixed(2),
        larg: x.L.largura.toFixed(3),
        bobinas: bob(x.d.bobinas)
      }))
    };
    nova.enfestoComprimento = nova.fases[0] ? nova.fases[0].comp : '';
    nova.enfestoLargura = nova.fases[0] ? nova.fases[0].larg : '';
    STATE.grades.push(nova);
    _pastaWiz.feito.criadas.push(`${nome} — ${nova.fases.length} fase(s)`);
    _pastaWiz.feito.fases += nova.fases.length;
    G.gradeId = nova.id;
  } else {
    /* ---- CORRIGIR grade existente ---- */
    const alvos = linhas.filter(x => x.d.alvo !== '__nova__').map(x => x.d.alvo);
    const repetido = alvos.find((a, i) => alvos.indexOf(a) !== i);
    if (repetido) return toast(`Dois riscos apontam para a fase "${repetido}" — corrija antes de aplicar`, 'err');
    if (linhas.some(x => x.d.alvo === '__nova__' && !x.d.nome)) return toast('A fase nova precisa de nome', 'err');

    grade.fases = grade.fases || [];
    let maiorOrdem = grade.fases.reduce((m, f) => Math.max(m, parseInt(f.ordem, 10) || 0), 0);
    const mudancas = [];
    linhas.forEach(x => {
      const comp = _riscoCompCadastro(x.L.comprimento, x.d.tecidoId).toFixed(2);
      const larg = x.L.largura.toFixed(3);
      let f = x.d.alvo === '__nova__' ? null : grade.fases.find(y => y.nome === x.d.alvo);
      if (!f) {
        f = { ordem: ++maiorOrdem, nome: x.d.nome, tecidoId: '', unidades: 2, comp: '', larg: '', bobinas: '' };
        grade.fases.push(f);
        mudancas.push(`+ ${f.nome}: ${comp} × ${larg}`);
      } else {
        mudancas.push(`${f.nome}: ${f.comp || '—'}×${f.larg || '—'} → ${comp}×${larg}`);
      }
      f.comp = comp;
      f.larg = larg;
      if (x.d.tecidoId) f.tecidoId = x.d.tecidoId;
      f.unidades = x.d.unidades;
      // Bobinas em branco não apaga o que já estava: quem deixou vazio não quis
      // dizer "zero", quis dizer "não sei" — e zero se digita como zero.
      if (x.d.bobinas !== '') f.bobinas = bob(x.d.bobinas);
      _pastaWiz.feito.fases++;
    });
    const f1 = grade.fases.slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0))[0];
    if (f1) { grade.enfestoComprimento = f1.comp || ''; grade.enfestoLargura = f1.larg || ''; }
    _pastaWiz.feito.atualizadas.push(`${grade.nome} — ${mudancas.join(' · ')}`);
  }

  // APRENDE: produto, fase, tecido e consumo. A próxima pasta deste mesmo
  // produto chega preenchida.
  const modeloKey = _normNome(G.modelo);
  if (modeloKey && G.draft.sku) {
    memProd[modeloKey] = {
      sku: G.draft.sku, tipoPeca: G.draft.tipoPeca, variacao: G.draft.variacao,
      pecasPorPacote: parseInt(G.draft.pecasPorPacote, 10) || 0
    };
  }
  linhas.forEach(x => {
    const k = _riscoChave(x.L);
    if (!k) return;
    const nomeFinal = x.d.alvo === '__nova__' || !grade ? x.d.nome : x.d.alvo;
    if (nomeFinal) memFase[k] = nomeFinal;
    if (x.d.tecidoId) memTec[k] = x.d.tecidoId;
    if (x.d.bobinas !== '') memBob[k] = bob(x.d.bobinas);
  });

  G.status = 'ok';
  await saveState('grades');
  await saveState('meta');
  toast(`${grade ? 'Grade atualizada' : 'Grade criada'} · ${linhas.length} fase(s)`, 'ok');
  _pastaAvancar();
}

function pastaFechar() {
  closeModal('modal-risco-pasta');
  _pastaWiz = null;
  renderGrades();
}

/* ========================================================= */
/*        PLANILHA DAS GRADES (.xlsx, sem biblioteca)        */
/* ========================================================= */
// Um .xlsx é um ZIP com alguns XML dentro. O app não carrega biblioteca de
// planilha, e não vale a pena carregar uma: o arquivo aqui é simples (duas abas
// de texto e número), e escrever o ZIP à mão são poucas linhas. Em troca, a
// exportação não depende de CDN nenhum e não quebra se um dia a rede cair.
//
// O ZIP sai SEM compressão (método 0, "stored"). Some a necessidade de um
// compressor, o Excel abre igual, e o arquivo continua pequeno — algumas
// dezenas de KB para as 64 grades.

const _CRC_TAB = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c >>> 0;
  }
  return t;
})();

function _crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = _CRC_TAB[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

// Data/hora no formato do MS-DOS, que é o que o cabeçalho do ZIP guarda.
function _zipDataDos(d) {
  return {
    hora: ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF,
    data: (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF
  };
}

// Monta o ZIP a partir de [{nome, texto}]. Devolve um Blob.
function _zipMontar(arquivos) {
  const enc = new TextEncoder();
  const agora = _zipDataDos(new Date());
  const partes = [], central = [];
  let offset = 0;
  const u16 = n => [n & 0xFF, (n >>> 8) & 0xFF];
  const u32 = n => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF];
  arquivos.forEach(f => {
    const nome = enc.encode(f.nome);
    const dados = enc.encode(f.texto);
    const crc = _crc32(dados);
    const local = new Uint8Array([
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0),
      ...u16(agora.hora), ...u16(agora.data),
      ...u32(crc), ...u32(dados.length), ...u32(dados.length),
      ...u16(nome.length), ...u16(0), ...nome
    ]);
    partes.push(local, dados);
    central.push(new Uint8Array([
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0),
      ...u16(agora.hora), ...u16(agora.data),
      ...u32(crc), ...u32(dados.length), ...u32(dados.length),
      ...u16(nome.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(0), ...u32(offset), ...nome
    ]));
    offset += local.length + dados.length;
  });
  const tamCentral = central.reduce((s, c) => s + c.length, 0);
  const fim = new Uint8Array([
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(arquivos.length), ...u16(arquivos.length),
    ...u32(tamCentral), ...u32(offset), ...u16(0)
  ]);
  return new Blob([...partes, ...central, fim], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

function _xmlEsc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    // Caracteres de controle não são válidos em XML e derrubam o Excel inteiro.
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
}

function _colLetra(n) {           // 0 → A, 25 → Z, 26 → AA
  let s = '';
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
  return s;
}

// Uma aba a partir de [[celula, ...], ...]. A primeira linha é o cabeçalho e sai
// em negrito (estilo 1). Número vira número de verdade — senão o Excel não soma
// nem ordena, e a planilha existe justamente para conferir números.
function _xlsxAba(linhas) {
  const corpo = linhas.map((linha, li) => {
    const cels = linha.map((val, ci) => {
      const ref = _colLetra(ci) + (li + 1);
      const est = li === 0 ? ' s="1"' : '';
      if (typeof val === 'number' && isFinite(val)) {
        return `<c r="${ref}"${est}><v>${val}</v></c>`;
      }
      const txt = _xmlEsc(val);
      if (txt === '') return `<c r="${ref}"${est}/>`;
      return `<c r="${ref}"${est} t="inlineStr"><is><t xml:space="preserve">${txt}</t></is></c>`;
    }).join('');
    return `<row r="${li + 1}">${cels}</row>`;
  }).join('');
  const nCols = linhas.reduce((m, l) => Math.max(m, l.length), 0);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>${Array.from({ length: nCols }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="18" customWidth="1"/>`).join('')}</cols>
<sheetData>${corpo}</sheetData>
<autoFilter ref="A1:${_colLetra(nCols - 1)}${linhas.length}"/>
</worksheet>`;
}

// Monta o .xlsx com N abas: [{nome, linhas}]. `sheetView` congela a primeira
// linha para o cabeçalho não sumir ao rolar — numa planilha de conferência com
// 200 linhas isso é a diferença entre dar para usar e não dar.
function _xlsxMontar(abas) {
  const arquivos = [
    { nome: '[Content_Types].xml', texto:
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${abas.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
</Types>` },
    { nome: '_rels/.rels', texto:
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>` },
    { nome: 'xl/workbook.xml', texto:
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${abas.map((a, i) => `<sheet name="${_xmlEsc(a.nome).slice(0, 31)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>
</workbook>` },
    { nome: 'xl/_rels/workbook.xml.rels', texto:
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${abas.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n')}
<Relationship Id="rId${abas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>` },
    { nome: 'xl/styles.xml', texto:
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>` }
  ];
  abas.forEach((a, i) => {
    arquivos.push({ nome: `xl/worksheets/sheet${i + 1}.xml`, texto: _xlsxAba(a.linhas) });
  });
  return _zipMontar(arquivos);
}

// A planilha das GRADES, para conferir os cadastros em lote fora do programa.
// Duas abas, porque são duas leituras diferentes do mesmo cadastro:
//   • "Grades" — uma linha por grade. É a visão de quem quer bater os tamanhos,
//     o tipo de peça e quantas fases cada uma tem, correndo o olho pela lista.
//   • "Fases das grades" — uma linha por FASE. É onde moram os números que mais
//     erram (tecido, unidades da grade, comprimento, largura, bobinas), e um por
//     linha permite filtrar e ordenar no Excel — "me mostre toda fase de ribana
//     sem comprimento", que na visão por grade seria impossível.
// Os dois trazem o que o programa já APUROU (uso em OS, tempo medido do
// enfesto), porque conferir cadastro sem ver o que ele produziu é meio caminho.
// O SKU de uma grade é o que vem depois do "|" no nome ("2M-4G-2GG | BM.TRICOLOR"
// → "BM.TRICOLOR"): antes da barra ficam os tamanhos, depois dela o produto.
// É por ele que as grades se parecem — todas as CM.BÁSICA deveriam ter as mesmas
// fases, os mesmos tecidos e o mesmo jeito de embalar, mudando só a distribuição
// de tamanhos. Agrupar por SKU é o que faz a divergência saltar.
function _skuDaGrade(g) {
  const n = String((g && g.nome) || '');
  const i = n.indexOf('|');
  const bruto = (i >= 0 ? n.slice(i + 1) : '').replace(/\s+/g, ' ').trim();
  return bruto || '(sem SKU no nome)';
}

function _linhasPlanilhaGrades() {
  const tamKeys = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
  // Ordenadas por SKU e, dentro dele, por nome: as semelhantes ficam juntas, uma
  // embaixo da outra, que é como se confere em lote. Ordenar só pelo nome punha
  // "2G-2G1 | CM.BICOLOR" ao lado de "2G-2G1 | CM.TRICOLOR" — vizinhas pela
  // grade de tamanhos, que é justamente o que NÃO precisa bater entre elas.
  const grades = (STATE.grades || []).slice().sort((a, b) =>
    _normNome(_skuDaGrade(a)).localeCompare(_normNome(_skuDaGrade(b)), 'pt-BR')
    || String(a.nome || '').localeCompare(String(b.nome || ''), 'pt-BR'));
  // Quantas OS usam cada grade — o cadastro que ninguém usa é o primeiro
  // candidato a estar errado, e o mais usado é o que urge conferir.
  const usoPorGrade = new Map();
  (STATE.ordens || []).forEach(o => {
    const k = _gradeIdDaOS(o);
    if (k) usoPorGrade.set(k, (usoPorGrade.get(k) || 0) + 1);
  });
  const rotuloTipo = { camiseta: 'Camiseta', blusa_moletom: 'Blusa moletom', outro: 'Outro' };
  const rotuloVar = { basica: 'Básica', bicolor: 'Bicolor', tricolor: 'Tricolor' };

  const abaGrades = [[
    'SKU', 'Grade', 'Tipo de peça', 'Variação',
    'P', 'M', 'G', 'GG', 'G1', 'G2', 'G3', 'Total da grade',
    'Peças por pacote', 'Nº de fases', 'Fases (nomes)',
    'Comprimento 1ª fase (m)', 'Largura 1ª fase (m)', 'Medidas por fase (comp x larg)',
    'Fases sem medida',
    'OS que usam', 'Tempo medido do enfesto (min)', 'Fases medidas'
  ]];
  const abaFases = [[
    'SKU', 'Grade', 'Tipo de peça', 'Variação', 'Fase nº', 'Nome da fase',
    'Tecido', 'Categoria do tecido', 'Unidades da grade',
    'Comprimento (m)', 'Largura (m)', 'Bobinas previstas',
    'Tempo médio medido (min)', 'Medições'
  ]];

  const porSku = new Map();
  grades.forEach(g => {
    const sku = _skuDaGrade(g);
    const chaveSku = _normNome(sku);
    if (!porSku.has(chaveSku)) {
      porSku.set(chaveSku, {
        sku, grades: 0, fases: 0, os: 0, porGrade: [], nomesFase: new Set(),
        tecidos: new Set(), pacotes: new Set(), semComp: 0, semTecido: 0, semFase: 0
      });
    }
    const R = porSku.get(chaveSku);
    const t = g.tamanhos || {};
    const totalGrade = tamKeys.reduce((s, k) => s + (parseInt(t[k], 10) || 0), 0);
    const fases = Array.isArray(g.fases) ? g.fases.slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0)) : [];
    let medidas = [];
    try { medidas = temposFasesDaGrade(g.id) || []; } catch (e) { medidas = []; }
    const medidaDe = nome => medidas.find(l => l.n > 0 && _normFaseNome(l.nome) === _normFaseNome(nome));
    const comTempo = medidas.filter(l => l.n > 0);
    const numOu = v => { const n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : ''; };
    // As medidas de TODAS as fases numa coluna só, para conferir a grade inteira
    // sem trocar de aba. Fase sem medida sai como "—", que é o que se procura.
    const medidasPorFase = fases.map(f => {
      const c = numOu(f.comp), l = numOu(f.larg);
      const rot = (f.nome || '').trim() || `F${f.ordem}`;
      return (c === '' && l === '')
        ? `${rot}: —`
        : `${rot}: ${c === '' ? '?' : String(c).replace('.', ',')}×${l === '' ? '?' : String(l).replace('.', ',')}`;
    }).join(' · ');
    const semMedida = fases.filter(f => numOu(f.comp) === '' || numOu(f.larg) === '').length;
    // O comprimento/largura gravado NA GRADE é espelho da 1ª fase (campo antigo,
    // mantido por compatibilidade). Sai identificado como tal, para ninguém o ler
    // como "a medida do enfesto inteiro" — as outras fases têm as suas.
    const compGrade = numOu(g.enfestoComprimento) !== '' ? numOu(g.enfestoComprimento) : numOu((fases[0] || {}).comp);
    const largGrade = numOu(g.enfestoLargura) !== '' ? numOu(g.enfestoLargura) : numOu((fases[0] || {}).larg);

    R.grades++;
    R.fases += fases.length;
    R.os += usoPorGrade.get(g.id) || 0;
    R.porGrade.push(fases.length);
    R.pacotes.add(parseInt(g.pecasPorPacote, 10) || 0);
    if (!fases.length) R.semFase++;
    fases.forEach(f => {
      const nf = (f.nome || '').trim();
      if (nf) R.nomesFase.add(nf);
      const tf = (STATE.tecidos || []).find(x => x.id === f.tecidoId);
      R.tecidos.add(tf ? tf.nome : (f.tecidoId ? '(tecido excluído)' : '(sem tecido)'));
      if (numOu(f.comp) === '' || numOu(f.larg) === '') R.semComp++;
      if (!tf) R.semTecido++;
    });

    abaGrades.push([
      sku,
      g.nome || '(sem nome)',
      rotuloTipo[g.tipoPeca] || g.tipoPeca || '',
      rotuloVar[g.variacao] || g.variacao || '',
      ...tamKeys.map(k => parseInt(t[k], 10) || 0),
      totalGrade,
      parseInt(g.pecasPorPacote, 10) || 0,
      fases.length,
      fases.map(f => (f.nome || '').trim() || `F${f.ordem}`).join(' · '),
      compGrade, largGrade, medidasPorFase, semMedida,
      usoPorGrade.get(g.id) || 0,
      comTempo.length ? comTempo.reduce((s, l) => s + l.mediaMin, 0) : '',
      comTempo.length ? `${comTempo.length} de ${fases.length}` : 'nenhuma'
    ]);

    if (!fases.length) {
      // Grade sem fase cadastrada aparece assim mesmo, com a linha vazia: numa
      // conferência em lote, o que FALTA é tão importante quanto o que está lá.
      abaFases.push([
        sku, g.nome || '(sem nome)', rotuloTipo[g.tipoPeca] || g.tipoPeca || '',
        rotuloVar[g.variacao] || g.variacao || '',
        '', '(nenhuma fase cadastrada)', '', '', '', '', '', '', '', ''
      ]);
      return;
    }
    fases.forEach(f => {
      const tec = (STATE.tecidos || []).find(x => x.id === f.tecidoId);
      const med = medidaDe(f.nome);
      const num = v => { const n = parseFloat(String(v).replace(',', '.')); return isFinite(n) ? n : ''; };
      abaFases.push([
        sku,
        g.nome || '(sem nome)',
        rotuloTipo[g.tipoPeca] || g.tipoPeca || '',
        rotuloVar[g.variacao] || g.variacao || '',
        parseInt(f.ordem, 10) || '',
        (f.nome || '').trim(),
        tec ? tec.nome : (f.tecidoId ? '(tecido excluído)' : ''),
        tec ? (categoriaEfetivaTecido(tec) || '') : '',
        parseInt(f.unidades, 10) || '',
        num(f.comp),
        num(f.larg),
        (f.bobinas === '' || f.bobinas == null) ? '' : (num(f.bobinas) === '' ? '' : num(f.bobinas)),
        med ? med.mediaMin : '',
        med ? med.n : 0
      ]);
    });
  });
  // ABA DE RESUMO — uma linha por SKU. É a visão de conferência por semelhança:
  // dentro de um SKU as grades só deveriam diferir na distribuição de tamanhos.
  // Fases, tecidos e peças por pacote deveriam ser os mesmos em todas. Quando um
  // SKU mostra "3 a 5 fases", ou dois valores de peças por pacote, ou um tecido a
  // mais que os outros, é aí que está o cadastro errado — e é isso que a linha
  // única por SKU põe na cara, sem precisar comparar 20 grades uma a uma.
  const abaSku = [[
    'SKU', 'Grades', 'Fases (total)', 'Fases por grade', 'Divergem?',
    'Nomes de fase usados no SKU', 'Tecidos usados no SKU',
    'Peças por pacote', 'Fases sem medida', 'Fases sem tecido',
    'Grades sem fase', 'OS que usam'
  ]];
  const listar = set => Array.from(set).sort((a, b) => String(a).localeCompare(String(b), 'pt-BR')).join(' · ');
  Array.from(porSku.values())
    .sort((a, b) => b.grades - a.grades || String(a.sku).localeCompare(String(b.sku), 'pt-BR'))
    .forEach(R => {
      const min = R.porGrade.length ? Math.min(...R.porGrade) : 0;
      const max = R.porGrade.length ? Math.max(...R.porGrade) : 0;
      const alertas = [];
      if (min !== max) alertas.push('nº de fases');
      if (R.pacotes.size > 1) alertas.push('peças por pacote');
      if (R.semComp) alertas.push('medida faltando');
      if (R.semTecido) alertas.push('tecido faltando');
      if (R.semFase) alertas.push('grade sem fase');
      abaSku.push([
        R.sku, R.grades, R.fases,
        min === max ? min : `${min} a ${max}`,
        alertas.length ? alertas.join(', ') : '',
        listar(R.nomesFase), listar(R.tecidos),
        listar(R.pacotes), R.semComp, R.semTecido, R.semFase, R.os
      ]);
    });

  return [
    { nome: 'Resumo por SKU', linhas: abaSku },
    { nome: 'Grades', linhas: abaGrades },
    { nome: 'Fases das grades', linhas: abaFases }
  ];
}

// Gera a planilha e grava na PASTA DO PROGRAMA (a mesma da cópia de dados, em
// Configurações). Sem pasta conectada, cai no download — a planilha é o
// objetivo, a pasta é a conveniência.
async function exportarGradesExcel() {
  if (!exigirEdicao('exportar a planilha das grades')) return;
  if (!(STATE.grades || []).length) return toast('Nenhuma grade cadastrada', 'err');
  let blob;
  try {
    blob = _xlsxMontar(_linhasPlanilhaGrades());
  } catch (e) {
    console.error('exportarGradesExcel', e);
    return toast('Falha ao montar a planilha: ' + (e.message || e), 'err');
  }
  const nome = 'grades-cadastradas.xlsx';
  // Nome FIXO, sem data: reexportar substitui o arquivo em vez de encher a pasta
  // de versões — é a mesma lição do PDF da OS.
  const pasta = backupFolderHandle || (await loadBackupFolderHandle())
    || pdfFolderHandle || (await loadPdfFolderHandle());
  if (pasta && await ensureFolderPermission(pasta, 'readwrite')) {
    try {
      const fh = await pasta.getFileHandle(nome, { create: true });
      const w = await fh.createWritable();
      await w.write(blob);
      await w.close();
      const nGrades = (STATE.grades || []).length;
      const nFases = (STATE.grades || []).reduce((s, g) => s + ((g.fases || []).length || 1), 0);
      // Caminho COMPLETO no aviso: dizer só o nome da pasta ("Gerador-OS")
      // mandava procurar na pasta do código, que tem o mesmo nome e não é onde
      // o arquivo cai. `resolve` devolve os nomes desde a raiz concedida.
      let onde = pasta.name;
      try {
        const partes = await pasta.resolve(fh);
        if (partes) onde = [pasta.name].concat(partes.slice(0, -1)).join('/');
      } catch (e) { /* navegador sem resolve: fica o nome da pasta */ }
      toast(`Planilha salva em ${onde}/${nome} — ${nGrades} grades e ${nFases} fases. `
        + 'É a pasta de backup conectada em Configurações.', 'ok');
      return;
    } catch (e) { console.warn('exportarGradesExcel (pasta)', e); }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nome; a.click();
  URL.revokeObjectURL(url);
  toast('Planilha das grades baixada (conecte a pasta em Configurações para gravá-la junto do programa)', 'ok');
}

async function exportarDados() {
  if (!exigirEdicao('exportar todos os dados')) return;
  const data = { exportadoEm: new Date().toISOString() };
  // Exporta TUDO (CAD_KEYS), inclusive estoqueMov (estoque de tecido) e os
  // movimentos de fase + osCounter — ALL_KEYS sozinho deixava o estoque de fora.
  CAD_KEYS.forEach(k => { data[k] = STATE[k]; });
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  // Nome DATADO: cada exportação é um ponto no tempo e não substitui a anterior.
  // É o oposto da planilha de grades e do PDF da OS, que têm nome fixo de
  // propósito — lá o que interessa é o estado atual, aqui é a história.
  const iso = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const nome = `BACKUP-COMPLETO-${iso}.json`;
  const n = CAD_KEYS.reduce((s, k) => s + (Array.isArray(STATE[k]) ? STATE[k].length : 0), 0);
  const mb = (blob.size / 1048576).toFixed(2);

  const pasta = exportFolderHandle || (await loadExportFolderHandle());
  if (pasta && await ensureFolderPermission(pasta, 'readwrite')) {
    try {
      const fh = await pasta.getFileHandle(nome, { create: true });
      const w = await fh.createWritable();
      await w.write(blob);
      await w.close();
      toast(`Exportado para ${pasta.name}/${nome} — ${n} registros, ${mb} MB`, 'ok');
      atualizarExportFolderStatus();
      return;
    } catch (e) {
      console.warn('exportarDados (pasta)', e);
      toast('Não deu para gravar na pasta — vai para os Downloads.', 'err');
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nome;
  a.click();
  URL.revokeObjectURL(url);
  toast(`Backup exportado (${n} registros, ${mb} MB) — conecte uma pasta em Configurações para gravar direto nela`, 'ok');
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

// Restauração CIRÚRGICA das OEs (expedição) a partir de um snapshot baixado.
// Ao contrário do "Restaurar snapshot" (tudo-ou-nada, reverte OS e cadastros),
// aqui só as chaves de expedição são mescladas — e por UNIÃO por id: adiciona o
// que falta no atual e NUNCA apaga o que já existe. Feito para recuperar OEs
// perdidas sem jogar fora o trabalho recente nas OSs. Aceita o formato do
// "Baixar snapshot"/"Exportar" (chaves com arrays reais).
async function restaurarExpedicaoDeArquivo(e) {
  if (!exigirAdmin('restaurar expedição')) { e.target.value = ''; return; }
  const file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  let data;
  try { data = JSON.parse(await file.text()); }
  catch { toast('Arquivo inválido (não é um JSON de snapshot).', 'err'); return; }
  const KEYS = ['expedicaoCargas', 'expedicaoJanelas', 'expedicaoExcecoes'];
  const plano = {};
  let totalAdd = 0;
  const resumo = [];
  KEYS.forEach(k => {
    let doArq = data[k];
    if (typeof doArq === 'string') { try { doArq = JSON.parse(doArq); } catch { doArq = []; } }
    doArq = Array.isArray(doArq) ? doArq : [];
    const atual = Array.isArray(STATE[k]) ? STATE[k] : [];
    const idsAtuais = new Set(atual.map(x => x && x.id).filter(Boolean));
    const faltando = doArq.filter(x => x && x.id && !idsAtuais.has(x.id));
    plano[k] = atual.concat(faltando);
    totalAdd += faltando.length;
    resumo.push(`• ${k.replace('expedicao', '')}: ${atual.length} agora + ${faltando.length} a restaurar (arquivo tem ${doArq.length})`);
  });
  if (!totalAdd) {
    toast('Nada a restaurar — o arquivo não tem OEs que faltem no atual (talvez seja posterior à perda; tente um snapshot mais antigo).', 'err');
    return;
  }
  const snapInfo = (data.__snapshot && data.__snapshot.data) ? `\nSnapshot de ${data.__snapshot.data}.` : '';
  if (!confirm(`Restaurar ${totalAdd} item(ns) de expedição (OE) que faltam?${snapInfo}\n\n${resumo.join('\n')}\n\nOs itens atuais NÃO são apagados — só adiciona o que falta.`)) return;
  try {
    for (const k of KEYS) { STATE[k] = plano[k]; await saveState(k); }
    toast(`${totalAdd} item(ns) de expedição restaurado(s).`, 'ok');
    const sec = document.querySelector('section.page[data-page="expedicao"]');
    if (sec && !sec.classList.contains('hidden') && typeof renderExpedicaoPlano === 'function') {
      trocarAbaExpedicao('plano');
    }
  } catch (err) {
    toast('Falha ao restaurar: ' + (err.message || err), 'err');
  }
}

// O APAGAR TUDO FOI REMOVIDO do programa (a pedido).
//
// Era o único caminho que zerava a base inteira de uma vez — e numa base
// compartilhada por toda a equipe, um clique errado levava junto as OS, os
// cadastros, o estoque e a expedição de todo mundo. Não havia motivo para ele
// existir: dado desta base se corrige RESTAURANDO um snapshot ou IMPORTANDO um
// backup, que são operações que trazem algo no lugar do que sai. Zerar não
// conserta nada.
//
// A função fica aqui, sem chamador nenhum, apenas para recusar: se um dia algo
// antigo (um atalho salvo, um botão esquecido, o console) ainda a chamar, a
// resposta é não. Apagar a função inteira faria essa chamada virar erro de
// JavaScript, que é uma forma pior de dizer a mesma coisa.
async function limparTudo() {
  toast('Apagar tudo não existe mais no programa. Para desfazer um estrago, '
    + 'restaure um snapshot em Configurações ou importe um backup.', 'err');
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
    await migrarLimpezaDesenho0023();  // remove componentes duplicados do 0023 (1×, admin)
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
window.calcularAlvoDeCamadas = calcularAlvoDeCamadas;
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
window.recarregarForcado = recarregarForcado;
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
window.imprimirEtiquetasPdf = imprimirEtiquetasPdf;
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
window.conectarPastaExport = conectarPastaExport;
window.desconectarPastaExport = desconectarPastaExport;
window.exportarGradesExcel = exportarGradesExcel;
window.abrirModalRisco = abrirModalRisco;
window.lerRiscosEscolhidos = lerRiscosEscolhidos;
window.riscoTrocarFase = riscoTrocarFase;
window.riscoTrocarGrade = riscoTrocarGrade;
window.aplicarRiscoNasGrades = aplicarRiscoNasGrades;
window.criarGradeDoRisco = criarGradeDoRisco;
window.riscoAtualizarNome = riscoAtualizarNome;
window.riscoMarcar = riscoMarcar;
window.importarDados = importarDados;
window.restaurarExpedicaoDeArquivo = restaurarExpedicaoDeArquivo;
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
window.baixarSnapshot = baixarSnapshot;
window.listarSnapshotsLocais = listarSnapshotsLocais;
window.restaurarSnapshotLocal = restaurarSnapshotLocal;
window.esconderAlertaSalvamento = esconderAlertaSalvamento;
window.setUserRole = setUserRole;
window.listarUsuariosComPapel = listarUsuariosComPapel;
window.duplicarCadastro = duplicarCadastro;
window.toggleFolderGrade = toggleFolderGrade;
window.moverEtapaForm = moverEtapaForm;
window.moverEtapaDesenho = moverEtapaDesenho;
