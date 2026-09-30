/*
 * ALGODÃO CRU DIVIDIDO POR LARGURA DE BOBINA.
 *
 * Junior, 30/09/2026: "corrija as quantidades e tipos de tecido algodao cru no
 * estoque de tecidos. Os tecidos disponiveis nessa cor deve ser diferenciado
 * entre 7 bobinas com 1,19m, 4 bobinas com 1,17m, 3 bobinas com 1,15m, 54
 * bobinas com 80cm." Pesos: as medias que ja estavam lancadas (decidido com ele).
 *
 * O que havia (Malha Algodao · Algodao cru, 04/09):
 *   - 14 bobinas / 252 kg num lancamento so  -> 18 kg por bobina
 *   - 40 bobinas / 520 kg, obs "Bobinas de 80 cm" -> 13 kg por bobina
 * O que fica:
 *   - o de 14 vira TRES lancamentos na mesma data: 7 x 119 cm (126 kg),
 *     4 x 117 cm (72 kg), 3 x 115 cm (54 kg) -- o total de 252 kg nao muda;
 *   - o de 40 ganha largura 80 cm;
 *   - entra um lancamento de CORRECAO, hoje: 14 bobinas x 80 cm, 182 kg
 *     (54 contadas - 40 lancadas, a 13 kg), com a observacao dizendo por que.
 *
 *   node servidor/corrigir-algodao-cru-3009.js            (so mostra)
 *   node servidor/corrigir-algodao-cru-3009.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const ID_14 = 'id_1788529628218_284';   // 14 bobinas / 252 kg
const ID_40 = 'id_1788529911602_899';   // 40 bobinas / 520 kg, 80 cm
const ID_CORRECAO = 'cru-80cm-correcao-2026-09-30';

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
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = arr(data.estoqueMov);
  if (movs.some(m => m.id === ID_CORRECAO)) { console.log('  ja corrigido antes -- nada a fazer'); return; }
  const m14 = movs.find(m => m.id === ID_14), m40 = movs.find(m => m.id === ID_40);
  if (!m14 || !m40) throw new Error('nao achei os dois lancamentos de 04/09 do algodao cru');
  if (+m14.kg !== 252 || +m14.fechados !== 14 || +m40.kg !== 520 || +m40.fechados !== 40) {
    throw new Error('os lancamentos mudaram desde a analise -- confira antes de rodar: '
      + JSON.stringify([m14.kg, m14.fechados, m40.kg, m40.fechados]));
  }
  const obsDiv = 'Dividido por largura em 30/09 (era um lancamento so: 14 bobinas / 252 kg, 18 kg por bobina)';
  const base = { tipo: 'entrada', tecidoNome: m14.tecidoNome, corNome: m14.corNome, abertos: 0,
                 data: m14.data, origem: 'manual', osId: '', osNumero: '' };
  const largas = [
    Object.assign({}, base, { id: ID_14, fechados: 7, largura: 119, kg: 126, obs: obsDiv }),
    Object.assign({}, base, { id: ID_14 + '-117', fechados: 4, largura: 117, kg: 72, obs: obsDiv }),
    Object.assign({}, base, { id: ID_14 + '-115', fechados: 3, largura: 115, kg: 54, obs: obsDiv })
  ];
  const novo40 = Object.assign({}, m40, { largura: 80 });
  const correcao = Object.assign({}, base, { id: ID_CORRECAO, fechados: 14, largura: 80, kg: 182,
    data: '2026-09-30',
    obs: 'Correcao de contagem 30/09: sao 54 bobinas de 80 cm (havia 40 lancadas); 14 x 13 kg, a media das 40' });

  const saida = [];
  movs.forEach(m => {
    if (m.id === ID_14) saida.push(...largas);
    else if (m.id === ID_40) saida.push(novo40);
    else saida.push(m);
  });
  saida.push(correcao);

  const cru = l => l.filter(m => m.origem !== 'os' && norm(m.tecidoNome) === 'malha algodao' && /cru/.test(norm(m.corNome)));
  const resumo = l => cru(l).map(m => `    ${m.data}  ${String(m.fechados).padStart(2)} bob  ${String(m.largura || '-').padStart(3)} cm  ${String(m.kg).padStart(5)} kg  ${m.obs || ''}`).join('\n');
  const soma = (l, k) => cru(l).reduce((a, m) => a + (+m[k] || 0), 0);
  console.log('ANTES:\n' + resumo(movs) + `\n    total: ${soma(movs, 'fechados')} bobinas, ${soma(movs, 'kg')} kg`);
  console.log('DEPOIS:\n' + resumo(saida) + `\n    total: ${soma(saida, 'fechados')} bobinas, ${soma(saida, 'kg')} kg`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-algodao-cru-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(saida) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.estoqueMov);
  console.log(`gravado. Agora: ${soma(d2, 'fechados')} bobinas, ${soma(d2, 'kg')} kg de algodao cru.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
