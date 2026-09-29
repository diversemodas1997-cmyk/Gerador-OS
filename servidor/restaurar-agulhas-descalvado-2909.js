/*
 * DEVOLVE AS AGULHAS DA UNIDADE DESCALVADO AO QUE ERAM ANTES DAS EDICOES DE 29/09.
 *
 * Junior, 29/09/2026: "Restaure os tipos e quantidades de agulhas na Unidade
 * Descalvado que estava antes das edicoes de hoje". As edicoes foram feitas de
 * 13:57 a 14:04 pela conta admin: B27 nº 11 e nº 14 e UY128GAS nº 11 mudaram de
 * estoque (e a UY128GAS nº 11 perdeu a descricao), e nasceram tres tipos novos
 * (UY128GAS nº 9, 10 e 12).
 *
 * A FONTE e o backup do repositorio privado de 28/09 20:05 (commit f6a741e), o
 * ultimo antes das edicoes. So as agulhas de DESCALVADO sao tocadas:
 *   - agulha que existia no backup volta inteira ao que era;
 *   - agulha que nao existia no backup sai do cadastro;
 *   - os movimentos de 29/09 dessas agulhas saem do historico.
 * Sao Carlos (unidade 'sc') fica como esta, inclusive a B27 nº 11 cadastrada la hoje.
 *
 *   node servidor/restaurar-agulhas-descalvado-2909.js            (so mostra)
 *   node servidor/restaurar-agulhas-descalvado-2909.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const RAIZ = path.join(__dirname, '..');
const REPO = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados';
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const COMMIT = 'f6a741e';
const DIA = '2026-09-29';
const GRAVAR = process.argv.includes('--gravar');

async function conectar() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  return { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
}
async function lerBlob(cab) {
  const l = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  return { data: l[0].data, updatedAt: l[0].updated_at };
}
async function gravarBlob(cab, data, updatedAt) {
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  }
}

const arr = v => { try { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; } catch (e) { return []; } };
const ehAgulhaDesc = p => /agulha/i.test(p.nome || '') && p.unidade !== 'sc';

(async () => {
  const velho = JSON.parse(execFileSync('C:/Program Files/Git/mingw64/bin/git.exe',
    ['show', COMMIT + ':dados-supabase.json'], { cwd: REPO, maxBuffer: 1 << 28 }).toString('utf8'));
  const antes = arr(velho.pecasCad).filter(ehAgulhaDesc);
  const antesPorId = new Map(antes.map(p => [p.id, p]));

  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const pecas = arr(data.pecasCad), movs = arr(data.pecasMov);

  const idsAgora = new Set(pecas.filter(ehAgulhaDesc).map(p => p.id));
  let mudou = 0;
  const novasPecas = [];
  pecas.forEach(p => {
    if (!ehAgulhaDesc(p)) { novasPecas.push(p); return; }
    const v = antesPorId.get(p.id);
    if (!v) { console.log('  sai do cadastro: ' + p.nome + ' (estoque ' + p.emEstoque + ')'); mudou++; return; }
    if (JSON.stringify(v) !== JSON.stringify(p)) {
      console.log('  volta: ' + p.nome + ' · estoque ' + p.emEstoque + ' -> ' + v.emEstoque
        + ' · uso ' + p.emUso + ' -> ' + v.emUso + (p.desc !== v.desc ? ' · descricao restaurada' : ''));
      mudou++;
    }
    novasPecas.push(v);
  });
  antes.forEach(v => {
    if (!idsAgora.has(v.id)) { console.log('  volta ao cadastro: ' + v.nome); novasPecas.push(v); mudou++; }
  });
  const idsAg = new Set([...idsAgora, ...antesPorId.keys()]);
  const novosMovs = movs.filter(m => {
    const tira = idsAg.has(m.itemId) && m.unidade !== 'sc' && String(m.data || '') === DIA;
    if (tira) { console.log('  sai do historico: ' + m.nome + ' · ' + m.motivo + ' · estoque ' + (m.dEstoque >= 0 ? '+' : '') + m.dEstoque + ' (' + m.em + ')'); mudou++; }
    return !tira;
  });
  if (!mudou) { console.log('  nada a fazer'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-restaurar-agulhas-descalvado-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { pecasCad: JSON.stringify(novasPecas), pecasMov: JSON.stringify(novosMovs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const ag2 = arr(d2.pecasCad).filter(ehAgulhaDesc);
  console.log('gravado. Descalvado agora tem ' + ag2.length + ' agulhas:');
  ag2.forEach(p => console.log('  ' + p.nome + ' · uso ' + p.emUso + ' · estoque ' + p.emEstoque));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
