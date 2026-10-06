/*
 * LINHA DE CADA GRADE CADASTRADA.
 *
 * Junior, 06/10/2026: "Determine como linha adulta todas as grades cadastradas
 * até o momento, com exceção das grades 2 ao 16 e 2-4-6-8-10".
 *
 * A ficha da grade ganhou o campo Linha (Adulto / Infantil) no mesmo dia. Este
 * script grava esse campo nas grades que já existiam: Infantil nas grades cujo
 * nome começa com "2 ao 16" ou "2-4-6-8-10", Adulto em todas as outras. Grade
 * que já tem a linha gravada não é tocada.
 *
 * Confere também os tamanhos: uma grade marcada Adulto com quantidade em 2 ao
 * 16 (ou Infantil com P ao G3) é listada como CONFLITO e não é gravada — a
 * ficha travaria a fileira dela e o número sumiria sem ninguém decidir.
 *
 *   node servidor/linha-das-grades-0610.js            (so mostra)
 *   node servidor/linha-das-grades-0610.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const ADULTO = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
const INFANTIL = ['t2', 't4', 't6', 't8', 't10', 't12', 't14', 't16'];
const norm = s => (s || '').toString().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .toLowerCase().replace(/[|-]/g, ' ').replace(/\s+/g, ' ').trim();
// Em QUALQUER parte do nome, não só no começo: a grade 2 ao 16 se chama
// "2x | 2 ao 16 | CM.LISA | 117cm".
const ehInfantilPeloNome = nome => ['2 ao 16', '2 4 6 8 10'].some(p => (' ' + norm(nome) + ' ').includes(' ' + p + ' '));

async function conectar() {
  if (!CREDS) throw new Error('nao achei supa-creds.json no Google Drive');
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^﻿/, '')).key;
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

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const eraTexto = typeof data.grades === 'string';
  const grades = arr(data.grades);
  if (!grades.length) throw new Error('nenhuma grade no servidor -- nada feito');

  const tem = (g, ks) => ks.some(k => (parseInt((g.tamanhos || {})[k], 10) || 0) > 0);
  let jaTinham = 0, adulto = 0, infantil = 0;
  const conflitos = [];
  grades.forEach(g => {
    if (g.linha) { jaTinham++; return; }
    const linha = ehInfantilPeloNome(g.nome) ? 'Infantil Unissex' : 'Adulto Unissex';
    const contra = linha === 'Infantil Unissex' ? tem(g, ADULTO) : tem(g, INFANTIL);
    if (contra) { conflitos.push(`${g.nome}  -> seria ${linha}, mas tem tamanhos da outra linha`); return; }
    g.linha = linha;
    if (linha === 'Infantil Unissex') { infantil++; console.log('  INFANTIL  ' + g.nome); } else adulto++;
  });
  console.log(`\n  ${grades.length} grades: ${adulto} Adulto, ${infantil} Infantil, ${jaTinham} ja tinham linha, ${conflitos.length} em conflito`);
  conflitos.forEach(c => console.log('  CONFLITO  ' + c));
  if (!adulto && !infantil) { console.log('  nada a gravar'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-linha-das-grades-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { grades: eraTexto ? JSON.stringify(grades) : grades }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.grades);
  const conta = l => d2.filter(g => g.linha === l).length;
  console.log(`gravado. No servidor agora: ${conta('Adulto Unissex')} Adulto, ${conta('Infantil Unissex')} Infantil, ${d2.filter(g => !g.linha).length} sem linha`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
