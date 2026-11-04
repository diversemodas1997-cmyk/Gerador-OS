/*
 * RISCO DA CM.LISA G-GG-G1-G2 COM O NOME NO PADRAO.
 *
 * Junior, 04/11/2026: "renomeie o PDF da G-GG-G1-G2 para o padrao".
 *
 *   CM.LISA/G-GG-G1-G2/CAMISETA BAS MASC - G GG G1 G2.pdf
 *   -> CM.LISA/G-GG-G1-G2/CM.LISA - CORPO - G-GG-G1-G2.pdf
 *
 * O arquivo foi renomeado no git; este script reaponta a fase Corpo da grade
 * "G-GG-G1-G2 | CM.LISA | 117 CM", que guarda o caminho de quando a medida foi
 * importada (fase.risco). Sem isto o ✓ da coluna Riscos some e a grade parece
 * nao ter vindo de risco nenhum (ver corrigir-risco-de-pasta-renomeada.js).
 * Troca so o valor IGUAL ao caminho antigo, em qualquer grade.
 *
 *   node servidor/risco-cmlisa-g-gg-g1-g2-0411.js            (so mostra)
 *   node servidor/risco-cmlisa-g-gg-g1-g2-0411.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const ANTES = 'CM.LISA/G-GG-G1-G2/CAMISETA BAS MASC - G GG G1 G2.pdf';
const DEPOIS = 'CM.LISA/G-GG-G1-G2/CM.LISA - CORPO - G-GG-G1-G2.pdf';

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

(async () => {
  const idx = JSON.parse(fs.readFileSync(path.join(RAIZ, 'dados', 'riscos-pdf.json'), 'utf8'));
  if (!idx.arquivos.includes(DEPOIS)) throw new Error('o PDF novo nao esta no indice: rode node servidor/indexar-riscos.js');

  const cab = await conectar();
  const l = await (await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  const data = l[0].data, updatedAt = l[0].updated_at;
  const grades = typeof data.grades === 'string' ? JSON.parse(data.grades) : data.grades;

  let n = 0;
  grades.forEach(g => (g.fases || []).forEach(f => {
    if (f.risco === ANTES) { f.risco = DEPOIS; n++; console.log(`grade ${g.nome} · fase ${f.nome}: reapontada`); }
  }));
  if (!n) { console.log('nenhuma fase aponta o caminho antigo - nada a mudar'); return; }
  if (!GRAVAR) { console.log(`\n${n} fase(s) a reapontar (so mostrou; rode com --gravar para gravar)`); return; }

  const bk = path.join(RAIZ, 'backups', `shared_data-antes-risco-cmlisa-g-gg-g1-g2-${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
  fs.writeFileSync(bk, JSON.stringify(l[0]));
  console.log('backup:', bk);

  data.grades = typeof data.grades === 'string' ? JSON.stringify(grades) : grades;
  const r = await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
    body: JSON.stringify({ data, updated_at: new Date().toISOString() })
  });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  try {
    await fetch(`${SUPA}/rest/v1/sync_signal?on_conflict=id`, {
      method: 'POST',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates' }),
      body: JSON.stringify({ id: 'main', updated_at: new Date().toISOString(), device_id: 'script-risco-cmlisa' })
    });
  } catch (e) { /* o polling do app acha a mudanca */ }
  console.log('gravado.');
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
