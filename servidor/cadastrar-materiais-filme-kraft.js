/*
 * CADASTRA O FILME PLASTICO E O PAPEL KRAFT NO ESTOQUE DE MATERIAIS.
 *
 * Junior, 28/09/2026: "Analise as imagens etiqueta filme plastico e etiqueta
 * papel kraft e cadastre como tipos de materiais em Estoque de materiais".
 *
 *   Filme plastico  -- Queiplas; bobina, folha; medidas 260 x 0,03 (largura em
 *                      cm x espessura em mm); o rolo da foto: 25,8 kg
 *   Papel kraft     -- KAE.32G.250.F3160: kraft aerado F03, 32 g/m2, rolo de
 *                      250 m x 160 cm, tubo de 3", 100% celulose
 *
 * Setor Corte nos dois (o kraft aerado e o filme vao na mesa de corte).
 * Unidade Descalvado, quantidades zeradas: a contagem entra depois.
 * Tipo que ja existe (mesmo nome + descricao + unidade + setor) fica como esta.
 *
 *   node servidor/cadastrar-materiais-filme-kraft.js            (so mostra)
 *   node servidor/cadastrar-materiais-filme-kraft.js --gravar   (grava)
 *
 * Antes de gravar salva o blob inteiro em backups/. A escrita so passa se
 * ninguem gravou no servidor desde a leitura (carimbo updated_at).
 * Depois de gravar: F5 no programa, senao aba aberta regrava o estado velho.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const normNome = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();

const MATERIAIS = [
  { nome: 'Filme plástico', setor: 'Corte',
    desc: 'Queiplas · bobina, folha · 260 × 0,03 (largura cm × espessura mm)',
    obs: 'Etiqueta do rolo: 25,8 kg' },
  { nome: 'Papel kraft aerado', setor: 'Corte',
    desc: 'KAE.32G.250.F3160 · F03 32 g/m² · rolo 250 m × 160 cm · tubo 3" · 100% celulose · fabricante CNPJ 92.791.243/0002-94 · EAN 7898921274654',
    obs: '' }
];

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
  let cad = [];
  try { cad = typeof data.materiaisEstCad === 'string' ? JSON.parse(data.materiaisEstCad) : (data.materiaisEstCad || []); } catch (e) { cad = []; }
  if (!Array.isArray(cad)) cad = [];
  // A mesma chave do app (salvarEstoqueItem com setor).
  const chaveDe = i => [i.nome, i.desc || '', i.unidade === 'sc' ? 'sc' : 'desc', i.setor || ''].map(normNome).join('||');
  const agora = new Date().toISOString();
  let n = 0;
  console.log('Estoque de materiais hoje: ' + cad.length + ' cadastro(s).\n');
  MATERIAIS.forEach((m, i) => {
    const x = { nome: m.nome, desc: m.desc, unidade: 'desc', setor: m.setor };
    if (cad.some(c => chaveDe(c) === chaveDe(x))) { console.log('  ja existe: ' + m.nome); return; }
    cad.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      ...x, emUso: 0, emEstoque: 0, obs: m.obs,
      atualizadoPor: 'Junior', atualizadoEm: agora, em: agora, por: 'Junior' });
    n++;
    console.log('  novo: ' + m.nome + ' [' + m.setor + '] -- ' + m.desc + (m.obs ? ' -- ' + m.obs : ''));
  });
  console.log('\n' + n + ' material(is) a cadastrar.');
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-materiais-filme-kraft-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiaisEstCad: JSON.stringify(cad) }), updatedAt);

  const { data: d2 } = await lerBlob(cab);
  const c2 = JSON.parse(d2.materiaisEstCad);
  const ok = MATERIAIS.filter(m => c2.some(c => chaveDe(c) === chaveDe({ nome: m.nome, desc: m.desc, unidade: 'desc', setor: m.setor }))).length;
  console.log('gravado. No estoque de materiais agora: ' + ok + ' de ' + MATERIAIS.length + ' (total ' + c2.length + ')');
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
