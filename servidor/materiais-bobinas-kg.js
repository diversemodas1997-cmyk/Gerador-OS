/*
 * PAPEL E FILME EM BOBINAS E KG: OS DOIS NUMEROS DE CONVERSAO.
 *
 * Junior, 28/09/2026: "estoque materiais deve mostrar plastico e papel no
 * formato de kg/bobinas". A conta segue em metros; a tela converte com:
 *   papel kraft: 250 m por bobina; 32 g/m2 x 1,60 m = 0,0512 kg por metro
 *                (bobina de 12,8 kg)
 *   filme:       360 m por bobina (ESTIMATIVA, ver materiais-enfesto-metros.js);
 *                25,8 kg / 360 m = 0,0717 kg por metro (bobina de 25,8 kg)
 *
 *   node servidor/materiais-bobinas-kg.js            (so mostra)
 *   node servidor/materiais-bobinas-kg.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const normNome = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
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


const CONV = {
  'papel kraft aerado': { mPorBobina: 250, kgPorM: 0.0512 },
  'filme plastico': { mPorBobina: 360, kgPorM: Math.round(25.8 / 360 * 10000) / 10000 }
};
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  let est = [];
  try { est = typeof data.materiaisEstCad === 'string' ? JSON.parse(data.materiaisEstCad) : (data.materiaisEstCad || []); } catch (e) { est = []; }
  let n = 0;
  est.forEach(x => {
    const c = CONV[normNome(x.nome)];
    if (!c || x.unidade === 'sc') return;
    if (x.mPorBobina === c.mPorBobina && x.kgPorM === c.kgPorM) { console.log('  ja tem: ' + x.nome); return; }
    Object.assign(x, c);
    n++;
    const bob = (x.emEstoque || 0) / c.mPorBobina;
    console.log('  ' + x.nome + ': ' + c.mPorBobina + ' m/bobina, ' + c.kgPorM + ' kg/m -> em estoque (contagem) ' + x.emEstoque + ' m = '
      + bob.toFixed(2) + ' bob. = ' + ((x.emEstoque || 0) * c.kgPorM).toFixed(1) + ' kg; em uso ' + x.emUso + ' m = ' + ((x.emUso || 0) / c.mPorBobina).toFixed(2) + ' bob.');
  });
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-materiais-bobinas-kg-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiaisEstCad: JSON.stringify(est) }), updatedAt);
  const e2 = JSON.parse((await lerBlob(cab)).data.materiaisEstCad);
  e2.filter(x => x.mPorBobina).forEach(x => console.log('gravado: ' + x.nome + ' ' + x.mPorBobina + ' m/bobina, ' + x.kgPorM + ' kg/m'));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
