/*
 * UM CADASTRO POR COR DE FIO E DE LINHA.
 *
 * Junior, 28/09/2026: "Corrija os cadastros de fios e linhas, pois cada cor
 * deve ser um cadastro diferente". O cadastro era um por tipo (Fio TP200,
 * Linha BR02C) com a lista de cores dentro; cada cor vira um cadastro seu:
 * { item, codigo (tipo), corNome, corCodigo, desc }. As 26 cores da lista da
 * fabrica (12 de fio, 14 de linha) viram 26 cadastros. O cadastro antigo, de
 * tipo com lista, sai.
 *
 *   node servidor/cadastro-por-cor-fio-linha.js            (so mostra)
 *   node servidor/cadastro-por-cor-fio-linha.js --gravar   (grava)
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


(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const tipos = ler('aviamentoTipos');
  const agora = new Date().toISOString();
  const novos = [];
  let n = 0;
  tipos.forEach(t => {
    if (!Array.isArray(t.cores)) { novos.push(t); return; }   // ja e um por cor
    t.cores.forEach((c, i) => {
      novos.push({ id: t.id + '_' + (i + 1), item: t.item, codigo: t.codigo, corNome: c.nome || '', corCodigo: c.codigo || '',
        desc: t.desc || '', atualizadoPor: 'Junior', atualizadoEm: agora });
      n++;
    });
    console.log('  ' + t.item + ' ' + t.codigo + ': ' + t.cores.length + ' cores -> ' + t.cores.length + ' cadastros');
  });
  // Nenhuma cor em dobro (mesmo item + tipo + cor).
  const vistos = new Set();
  const finais = novos.filter(x => { const k = [x.item, x.codigo, x.corNome].map(normNome).join('||'); if (vistos.has(k)) return false; vistos.add(k); return true; });
  console.log('\n' + n + ' cadastro(s) por cor; total ' + finais.length);
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cadastro-por-cor-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(finais) }), updatedAt);
  const t2 = JSON.parse((await lerBlob(cab)).data.aviamentoTipos);
  const conta = {};
  t2.forEach(t => { const k = t.item + ' ' + t.codigo; conta[k] = (conta[k] || 0) + 1; });
  console.log('gravado: ' + Object.entries(conta).map(([k, v]) => k + ' = ' + v + ' cadastros').join(', ') + (t2.some(t => t.cores) ? ' -- ATENCAO: sobrou cadastro com lista' : ''));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
