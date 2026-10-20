/*
 * CORPO DAS OS 0670 E 0673 = 2,33 M.
 *
 * Junior, 20/10/2026: "Passe a 0670 e a 0673 para 2,33". As duas conjugadas
 * (M-G CM.LISA da P-G-GG-G1-G3 CM.REC) nasceram com o Corpo 2,35 copiado da
 * grade antes de ela ser acertada para 2 cm a menos que o Corpo 2 da REC
 * (corpo-lisa-conjugada-menos-2cm-2010.js). Aqui vai o mesmo número para a
 * fase Corpo da OS e para o enfesto (comprimento e o bloco da fase).
 * As conjugadas não reservam pano, então não há movimento de estoque a refazer.
 *
 *   node servidor/corpo-0670-0673-2010.js            (só mostra)
 *   node servidor/corpo-0670-0673-2010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
// A conjugada da 0669 era a 0673 e foi renumerada para 0672 no app (20/10); a
// 0673 de agora é uma CM.TRI. A trava abaixo (conjugadaPaiId) para se o número
// apontar outra OS.
const OSS = ['0670', '0672'];
const COMP = '2.33';

async function req(method, caminho, body) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: { apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' },
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t;
}
const arr = s => (typeof s === 'string' ? JSON.parse(s || '[]') : (s || []));

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const ordens = arr(data.ordens);
  for (const n of OSS) {
    const o = ordens.find(x => x.os === n);
    if (!o) throw new Error('OS ' + n + ' não encontrada');
    if (!o.conjugadaPaiId) throw new Error('OS ' + n + ' não é conjugada — parei');
    const i = (o.fases || []).findIndex(f => /^corpo/i.test(f.nome));
    if (i < 0) throw new Error('OS ' + n + ' sem fase Corpo');
    const antes = [o.fases[i].comp, o.enfesto && o.enfesto.comprimento, o.enfesto && o.enfesto.blocos && o.enfesto.blocos[i] && o.enfesto.blocos[i].comp];
    o.fases[i].comp = COMP;
    if (o.enfesto) {
      if (i === 0) o.enfesto.comprimento = parseFloat(COMP);
      if (Array.isArray(o.enfesto.blocos) && o.enfesto.blocos[i]) o.enfesto.blocos[i].comp = parseFloat(COMP);
    }
    console.log(`OS ${n} (${o.grade && o.grade.descricao})\n   fase Corpo ${antes[0]} -> ${COMP} · enfesto ${antes[1]} -> ${o.enfesto && o.enfesto.comprimento} · bloco ${antes[2]} -> ${o.enfesto && o.enfesto.blocos && o.enfesto.blocos[i] && o.enfesto.blocos[i].comp}`);
  }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-corpo-0670-0673-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    ordens: typeof data.ordens === 'string' ? JSON.stringify(ordens) : ordens,
    _device: 'corpo-0670-0673-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const o2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.ordens);
  OSS.forEach(n => {
    const o = o2.find(x => x.os === n);
    console.log(`gravado OS ${n}: Corpo ${(o.fases.find(f => /^corpo/i.test(f.nome)) || {}).comp} · enfesto ${o.enfesto.comprimento}`);
  });
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
