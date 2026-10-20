/*
 * CORPO DA CM.LISA CONJUGADA = CORPO 2 DA CM.REC MENOS 2 CM.
 *
 * Junior, 20/10/2026: "Acerte as três com 2 cm a menos". Cinco das oito duplas
 * já tinham a CM.LISA 2 cm mais curta que a fase Corpo 2 da CM.REC; três não:
 * 2G-2G1 (+3 cm), M-G-G1-G3 (+3 cm) e P-G-GG-G1-G3 (igual — criada hoje copiando
 * a REC sem descontar). Só o cadastro das grades CM.LISA: OS já emitidas não
 * são tocadas.
 *
 *   node servidor/corpo-lisa-conjugada-menos-2cm-2010.js            (só mostra)
 *   node servidor/corpo-lisa-conjugada-menos-2cm-2010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const ATIVAS = ['2G-2G1 | CM.REC | 117cm', 'M-G-G1-G3 | CM.REC | 117cm', 'P-G-GG-G1-G3 | CM.REC | 117cm'];

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
  const grades = arr(data.grades);
  const mexidas = [];
  for (const nome of ATIVAS) {
    const rec = grades.find(g => g.nome.split(' (')[0].trim() === nome && g.conjugadaGradeId);
    if (!rec) throw new Error('não achei a ativa ' + nome);
    const lisa = grades.find(g => g.id === rec.conjugadaGradeId);
    const c2 = (rec.fases || []).find(f => /corpo.*2/i.test(f.nome));
    const corpo = (lisa.fases || []).find(f => /^corpo/i.test(f.nome));
    if (!c2 || !corpo) throw new Error('fase Corpo 2 / Corpo não achada em ' + nome);
    const novo = (Math.round(parseFloat(c2.comp) * 100) - 2) / 100;
    const txt = novo.toFixed(2);
    console.log(`${lisa.nome}\n   Corpo ${corpo.comp} -> ${txt}   (Corpo 2 da REC: ${c2.comp})`);
    // O comprimento de enfesto da grade é o da 1ª fase; acompanha quando era igual.
    if (String(lisa.enfestoComprimento) === String(corpo.comp)) lisa.enfestoComprimento = txt;
    corpo.comp = txt;
    mexidas.push(lisa.id);
  }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-corpo-lisa-menos-2cm-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    grades: typeof data.grades === 'string' ? JSON.stringify(grades) : grades,
    _device: 'corpo-lisa-menos-2cm-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const g2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.grades);
  mexidas.forEach(id => {
    const g = g2.find(x => x.id === id);
    console.log(`gravado: ${g.nome} · Corpo ${(g.fases.find(f => /^corpo/i.test(f.nome)) || {}).comp} · enfesto ${g.enfestoComprimento}`);
  });
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
