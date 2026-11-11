/* Rode com:  node testes/arquivo-os.js
   Exercita o ARQUIVO DE OS do app.js real: a escolha das OS, a gravação só das
   ativas, a volta da OS mexida e a rotina inteira contra um Supabase falso.
   Com um backup como argumento, mostra também quantas OS reais iriam:
     node testes/arquivo-os.js backups/shared_data-antes-....json */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('/* =========================================================\n   ARQUIVO DE OS');
const fim = src.indexOf('/* ===================== fim do ARQUIVO DE OS');
if (ini < 0 || fim < 0) { console.error('nao achei os limites do trecho'); process.exit(1); }
const trecho = src.slice(ini, fim);

let falhas = 0;
const ok = (cond, msg) => { console.log((cond ? '  ok   ' : '  FALHA ') + msg); if (!cond) falhas++; };

/* ---- Supabase falso: a tabela os_arquivadas em memória ---- */
function criarSupa(tab, opts = {}) {
  let seq = 0;
  return {
    from() {
      const q = { _f: [], _lim: 1e9 };
      q.select = () => q;
      q.gt = (c, v) => { q._f.push(r => r[c] > v); return q; };
      q.in = (c, vs) => { q._f.push(r => vs.includes(r[c])); return q; };
      q.order = () => q;
      q.limit = n => { q._lim = n; return q; };
      q.then = (res, rej) => {
        if (opts.semTabela) return Promise.resolve({ data: null, error: { code: 'PGRST205' } }).then(res, rej);
        let linhas = [...tab.values()].filter(r => q._f.every(f => f(r))).sort((a, b) => a.seq - b.seq).slice(0, q._lim);
        if (q._upd) { linhas.forEach(r => Object.assign(r, q._upd, { seq: ++seq })); linhas = null; }
        return Promise.resolve({ data: linhas && JSON.parse(JSON.stringify(linhas)), error: null }).then(res, rej);
      };
      q.upsert = rows => {
        rows.forEach(r => tab.set(r.id, { ...r, dados: opts.corromper ? r.dados + ' ' : r.dados, seq: ++seq }));
        return Promise.resolve({ error: null });
      };
      q.update = u => { q._upd = u; return q; };
      return q;
    }
  };
}

function montar(ordens, supa, extra = {}) {
  const gravacoes = [];
  const ctx = {
    console: { info() {}, warn() {}, log: console.log },
    setTimeout: (f, ms) => (ms >= 1000 ? 0 : setTimeout(f, ms)), clearTimeout() {},
    Promise, Map, Set, JSON, Date, Math, Number, String, parseInt,
    STATE: { ordens }, supa, currentRole: 'admin', currentUser: { id: 'u1' },
    podeGravar: () => true, printOsAtual: null, osEditId: null, _servidorUrlAtivo: 'x',
    _statusOS: o => o.statusOS || 'nao-iniciado',
    toast() {}, document: { querySelector: () => null },
    saveState: async k => { gravacoes.push(JSON.parse(ctx._osArqSerializarOrdens())); },
    ...extra
  };
  vm.createContext(ctx);
  vm.runInContext(trecho + '\nthis._osArq = _osArq; this._osArqSerializarOrdens = _osArqSerializarOrdens;'
    + 'this._osArqEscolher = _osArqEscolher; this._osArqJuntarNoState = _osArqJuntarNoState;'
    + 'this.arquivarOSAntigas = arquivarOSAntigas; this._osArqFiltrarLista = _osArqFiltrarLista;'
    + 'this._osArqLinhaFinalHtml = _osArqLinhaFinalHtml; this.verArq = v => { _osArqVerNaLista = v; };', ctx);
  ctx.gravacoes = gravacoes;
  return ctx;
}

const os = (n, st, extra = {}) => ({ id: 'id' + n, os: String(n).padStart(4, '0'), statusOS: st, criadoEm: '2026-01-01', ...extra });

(async () => {
  console.log('1. escolha');
  const lista = [];
  // 'estoque' é o fim do caminho hoje (11/11/2026); 'finalizado' é carimbo antigo.
  for (let n = 1; n <= 150; n++) lista.push(os(n, 'estoque'));
  lista[4].statusOS = 'parado';                        // 0005 antiga, ainda ativa
  lista[9].statusOS = 'cancelado';                     // 0010 cancelada: vai
  lista[14].statusOS = 'finalizado';                   // 0015 carimbo antigo: vai
  lista[19].conjugadaStatusPaiId = 'id21';             // 0020 segue a 0021...
  lista[20].statusOS = 'ensacado';                     // ...que ainda anda: 0020 fica
  lista[29].conjugadaPaiId = 'id31';                   // 0030 e 0031 finalizadas: vão juntas
  lista.push({ id: 'semnum', os: '', statusOS: 'finalizado' });
  let ctx = montar(lista, null);
  const esc = ctx._osArqEscolher(lista, o => o.statusOS, 100, new Set());
  ok(esc.length === 47, `47 escolhidas (50 antigas - 0005 ativa - 0020/0021 família ativa): ${esc.length}`);
  ok(!esc.some(id => Number(id.slice(2)) > 50), 'nenhuma das 100 mais recentes (0051-0150)');
  ok(!esc.includes('id5') && !esc.includes('id20') && !esc.includes('id21'), 'ativa e família de ativa ficam');
  ok(esc.includes('id10') && esc.includes('id30') && esc.includes('id31'), 'cancelada e família toda em Estoque vão');
  ok(esc.includes('id1') && esc.includes('id15'), 'OS em Estoque vai; carimbo antigo "finalizado" também');
  ok(!ctx._osArqEscolher([os(1, 'ensacado'), os(2, 'transito-volta'), os(3, 'estoque-fio')], o => o.statusOS, 0, new Set()).length,
    'ensacada, em trânsito e com fio ficam ativas');
  ok(!esc.includes('semnum'), 'OS sem número fica');
  ok(esc[0] === 'id1', 'da mais antiga para a mais nova');

  console.log('2. rotina inteira');
  const tab = new Map();
  const ordens = JSON.parse(JSON.stringify(lista));
  ctx = montar(ordens, criarSupa(tab));
  await ctx.arquivarOSAntigas();
  ok(tab.size === 47, `47 linhas na tabela: ${tab.size}`);
  ok(ctx.gravacoes.length === 1 && ctx.gravacoes[0].length === 151 - 47, `blob gravado com ${151 - 47} OS: ${ctx.gravacoes[0] && ctx.gravacoes[0].length}`);
  ok(ctx.STATE.ordens.length === 151, 'memória continua com as 151 (estoque e relatórios inalterados)');
  ok(ctx._osArqFiltrarLista(ctx.STATE.ordens, false).length === 104, 'lista mostra 104 ativas por padrão');
  ok(ctx._osArqFiltrarLista(ctx.STATE.ordens, true).length === 151, 'com busca, mostra todas');
  // A última linha da lista (11/11/2026): o texto clicável.
  ok(/ver OS arquivadas \(47\)/.test(ctx._osArqLinhaFinalHtml(false)), 'última linha: "ver OS arquivadas (47)"');
  ok(ctx._osArqLinhaFinalHtml(true) === '', 'com busca ou grupo elas já aparecem: sem a linha');
  ctx.verArq(true);
  ok(/esconder as OS arquivadas \(47\)/.test(ctx._osArqLinhaFinalHtml(false)), 'clicado: vira "esconder"');
  ok(ctx._osArqFiltrarLista(ctx.STATE.ordens, false).length === 151, 'clicado: a lista mostra as arquivadas também');
  ctx.verArq(false);
  await ctx.arquivarOSAntigas();
  ok(ctx.gravacoes.length === 1, 'segunda rodada não tem o que arquivar');

  console.log('3. OS arquivada mexida volta; excluída some');
  const o1 = ctx.STATE.ordens.find(o => o.id === 'id1');
  o1.obs = 'mexi depois de arquivada';
  ctx.STATE.ordens = ctx.STATE.ordens.filter(o => o.id !== 'id2');
  const gravado = JSON.parse(ctx._osArqSerializarOrdens());
  ok(gravado.some(o => o.id === 'id1'), 'a mexida volta ao blob');
  ok(!ctx._osArq.porId.has('id1') && !ctx._osArq.porId.has('id2'), 'saem do arquivo da memória');
  await new Promise(r => setTimeout(r, 20));
  ok(tab.get('id1').removida === true && tab.get('id2').removida === true, 'marcadas como removidas na tabela');

  console.log('4. carga em outro computador');
  const blob = JSON.parse(JSON.stringify(gravado));
  const ctx2 = montar(blob, criarSupa(tab));
  ctx2.STATE.ordens = blob;
  const { _osArq } = ctx2;
  await vm.runInContext('_osArqSincronizar()', ctx2);
  ctx2._osArqJuntarNoState();
  ok(_osArq.porId.size === 45, `45 arquivadas lidas (47 - 2 removidas): ${_osArq.porId.size}`);
  ok(ctx2.STATE.ordens.length === 150, `memória com 150 (151 - a excluída): ${ctx2.STATE.ordens.length}`);
  ok(new Set(ctx2.STATE.ordens.map(o => o.id)).size === 150, 'sem duplicadas');
  // A gravação de quem arquivou ainda não chegou: a OS está no blob E no arquivo.
  const ctx3 = montar(JSON.parse(JSON.stringify(lista)), criarSupa(tab));
  await vm.runInContext('_osArqSincronizar()', ctx3);
  ctx3._osArqJuntarNoState();
  ok(ctx3.STATE.ordens.length === 151 && ctx3._osArq.porId.size === 45, 'blob atrasado: sem duplicar e sem desarquivar');

  console.log('5. sem tabela, ou gravação que não confere: nada sai do blob');
  const ctx4 = montar(JSON.parse(JSON.stringify(lista)), criarSupa(new Map(), { semTabela: true }));
  await ctx4.arquivarOSAntigas();
  ok(ctx4.gravacoes.length === 0 && ctx4._osArq.porId.size === 0, 'sem tabela: nada arquivado');
  const ctx5 = montar(JSON.parse(JSON.stringify(lista)), criarSupa(new Map(), { corromper: true }));
  await ctx5.arquivarOSAntigas();
  ok(ctx5.gravacoes.length === 0 && ctx5._osArq.porId.size === 0, 'texto não confere: nada arquivado');

  const bkp = process.argv[2];
  if (bkp) {
    console.log('6. dados reais:', path.basename(bkp));
    const d = JSON.parse(fs.readFileSync(bkp, 'utf8'));
    const reais = JSON.parse((d.data || d).ordens);
    const c = montar(reais, null);
    const ids = c._osArqEscolher(reais, o => o.statusOS || 'nao-iniciado', 100, new Set());
    const set = new Set(ids);
    const mb = s => (s.length / 1048576).toFixed(2) + ' MB';
    console.log(`  ${reais.length} OS; arquivaria ${ids.length}; ficam ${reais.length - ids.length}`);
    console.log(`  ordens no blob: ${mb(JSON.stringify(reais))} -> ${mb(JSON.stringify(reais.filter(o => !set.has(String(o.id)))))}`);
  }

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo ok');
  process.exit(falhas ? 1 : 0);
})();
