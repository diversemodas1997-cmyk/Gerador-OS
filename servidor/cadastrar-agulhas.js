/*
 * CADASTRA OS TIPOS DE AGULHA NO ESTOQUE DE PECAS.
 *
 * Junior, 28/09/2026: "Analise a imagem Agulhas Diverse na pasta download e
 * cadastre todos os tipos de agulhas o Estoque de pecas". A foto tem 12
 * cartelas; duas sao o mesmo tipo (Groz-Beckert 774012, B27 75/11), entao
 * sao 11 tipos. A maquina vem do que esta escrito a mao em cada cartela.
 *
 * So cadastra os TIPOS (Unidade Descalvado), com em uso e em estoque zerados:
 * a foto nao diz quantas agulhas restam em cada cartela. A contagem entra
 * depois pelo "editar" ou pelo "+ entrada" da linha.
 * Tipo que ja existe (mesmo nome + descricao + unidade) fica como esta.
 *
 *   node servidor/cadastrar-agulhas.js            (so mostra)
 *   node servidor/cadastrar-agulhas.js --gravar   (grava)
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

// nome = sistema + numero (como o exemplo do campo: "Agulha DBx1 no 11");
// desc = marca, codigo, Nm, ponta e a maquina escrita na cartela.
const AGULHAS = [
  { nome: 'Agulha DBx1 nº 11', desc: 'Groz-Beckert 786605 (1738) · Nm 75/11 · dur FFG/SES · máquina: Reta' },
  { nome: 'Agulha DBx1 nº 14', desc: 'Groz-Beckert 773395 (1738) · Nm 90/14 · dur FFG/SES · máquina: Reta' },
  { nome: 'Agulha B27 nº 11', desc: 'Groz-Beckert 774012 (81x1 / DCx27) · Nm 75/11 · GEBEDUR FFG/SES · máquina: Overloque' },
  { nome: 'Agulha B27 nº 14', desc: 'Groz-Beckert 773935 (81x1) · Nm 90/14 · dur FFG/SES · máquina: Overloque' },
  { nome: 'Agulha UY128GAS nº 11', desc: 'Groz-Beckert 781705 (UY128GBS / TVx3 · 029) · Nm 75/11 · dur FFG/SES · máquina: Galoneira' },
  { nome: 'Agulha UY128GAS nº 14', desc: 'Groz-Beckert 389.200 BC01 (UY128GBS / 1280 / 149x3 / 149x31 / TVx3 · 036) · Nm 90/14 · FFG/SES · máquina: Galoneira' },
  { nome: 'Agulha 149x7 nº 14', desc: 'Groz-Beckert 720382 (TVx7 / MY1002A) · Nm 90/14 · FFG/SES · máquina: Ombro a ombro' },
  { nome: 'Agulha DPx5 nº 14', desc: 'Groz-Beckert 717675 (134) · Nm 90/14 · dur FFG/SES · máquina: Caseadeira' },
  { nome: 'Agulha DPx5 nº 14 ponta bola', desc: 'Singer 1955-06 (DPx5 L BALL · cabo grosso) · 90/14 · máquina: Pespontadeira' },
  { nome: 'Agulha DOx558 nº 12', desc: 'Groz-Beckert 755002 (558) · Nm 80/12 · RS/SPI · máquina: Caseadeira' },
  { nome: 'Agulha TQx1 nº 12', desc: 'Groz-Beckert 726392 (1985 / 175x1) · Nm 80/12 · RG · máquina: Botoneira' }
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
  let pecas = [];
  try { pecas = typeof data.pecasCad === 'string' ? JSON.parse(data.pecasCad) : (data.pecasCad || []); } catch (e) { pecas = []; }
  if (!Array.isArray(pecas)) pecas = [];

  // A mesma chave do app (salvarEstoqueItem): nome + descricao + unidade.
  const chaveDe = i => normNome(i.nome) + '||' + normNome(i.desc || '') + '||' + (i.unidade === 'sc' ? 'sc' : 'desc');
  const agora = new Date().toISOString();
  const novos = [];
  console.log('Estoque de pecas hoje: ' + pecas.length + ' cadastro(s).\n');
  AGULHAS.forEach((a, i) => {
    const x = { nome: a.nome, desc: a.desc, unidade: 'desc' };
    if (pecas.some(p => chaveDe(p) === chaveDe(x))) { console.log('  ja existe: ' + a.nome); return; }
    novos.push({
      id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      ...x, emUso: 0, emEstoque: 0, obs: '',
      atualizadoPor: 'Junior', atualizadoEm: agora, em: agora, por: 'Junior'
    });
    console.log('  novo: ' + a.nome + '  --  ' + a.desc);
  });
  console.log('\n' + novos.length + ' tipo(s) a cadastrar.');
  if (!novos.length) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-cadastrar-agulhas-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);

  const lista = pecas.concat(novos);
  const novo = Object.assign({}, data, {
    pecasCad: typeof data.pecasCad === 'object' && data.pecasCad !== null ? lista : JSON.stringify(lista)
  });
  await gravarBlob(cab, novo, updatedAt);

  const { data: d2 } = await lerBlob(cab);
  const p2 = typeof d2.pecasCad === 'string' ? JSON.parse(d2.pecasCad) : d2.pecasCad;
  const ok = AGULHAS.filter(a => p2.some(p => chaveDe(p) === chaveDe({ nome: a.nome, desc: a.desc, unidade: 'desc' }))).length;
  console.log('gravado. Agulhas no cadastro agora: ' + ok + ' de ' + AGULHAS.length);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
