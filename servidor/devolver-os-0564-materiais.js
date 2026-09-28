/*
 * DEVOLVE OS 7,05 m DA OS 0564 AO PAPEL KRAFT E AO FILME.
 *
 * Junior, 28/09/2026: "A OS 0564 foi enfestada antes da contagem, tire os
 * 7,05 m". A contagem de 4 bobinas e meia ja veio sem o papel e o filme dela,
 * e a baixa automatica (lida da OS, ver _matDasOS) descontou de novo. Entra um
 * AJUSTE de +7,05 m em cada, que compensa a baixa: o historico mostra as duas
 * coisas. Rodar de novo nao soma outra vez (procura o ajuste pela observacao).
 *
 *   node servidor/devolver-os-0564-materiais.js            (so mostra)
 *   node servidor/devolver-os-0564-materiais.js --gravar   (grava)
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


const OBS = 'OS 0564 enfestada antes da contagem';
const METROS = 7.05;
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const cad = ler('materiaisEstCad'), movs = ler('materiaisEstMov');
  const agora = new Date().toISOString();
  const hoje = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  let n = 0;
  cad.filter(x => x.baixaOS && x.unidade !== 'sc').forEach((x, i) => {
    if (movs.some(m => m.itemId === x.id && m.obs === OBS)) { console.log('  ja devolvido: ' + x.nome); return; }
    const antes = Number(x.emEstoque) || 0;
    x.emEstoque = Math.round((antes + METROS) * 100) / 100;
    x.atualizadoPor = 'Junior'; x.atualizadoEm = agora;
    movs.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      itemId: x.id, nome: x.nome, desc: x.desc || '', unidade: 'desc', setor: x.setor || '', data: hoje,
      obs: OBS, tipo: 'ajuste', motivo: 'correcao', qtd: 0, dUso: 0, dEstoque: METROS, por: 'Junior', em: agora });
    n++;
    console.log('  ' + x.nome + ': contagem gravada ' + antes + ' -> ' + x.emEstoque + ' m');
  });
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-devolver-0564-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiaisEstCad: JSON.stringify(cad), materiaisEstMov: JSON.stringify(movs) }), updatedAt);
  const c2 = JSON.parse((await lerBlob(cab)).data.materiaisEstCad);
  c2.filter(x => x.baixaOS).forEach(x => console.log('gravado: ' + x.nome + ' ' + x.emEstoque + ' m (menos os 7,05 da OS 0564 = ' + Math.round((x.emEstoque - METROS) * 100) / 100 + ' m na tela)'));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
