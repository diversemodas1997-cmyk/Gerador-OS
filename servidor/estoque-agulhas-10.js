/*
 * POE 10 EM ESTOQUE EM CADA AGULHA CADASTRADA POR cadastrar-agulhas.js.
 *
 * Junior, 28/09/2026: "coloque 10 em estoque em cada agulha". Como o
 * "editar" do app: a quantidade mudada a mao vira um AJUSTE (correcao) no
 * historico (pecasMov), com dEstoque = +10. So mexe na agulha que ainda esta
 * com 0 em estoque -- rodar de novo nao soma outra vez.
 *
 * Abaixo, o cabecalho do script de cadastro (a lista AGULHAS vem dele):
 *
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
 *   node servidor/estoque-agulhas-10.js            (so mostra)
 *   node servidor/estoque-agulhas-10.js --gravar   (grava)
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
const QTD = 10;
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
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const pecas = ler('pecasCad'), movs = ler('pecasMov');
  const chaveDe = i => normNome(i.nome) + '||' + normNome(i.desc || '') + '||' + (i.unidade === 'sc' ? 'sc' : 'desc');
  const agora = new Date().toISOString();
  const hoje = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  let n = 0;
  AGULHAS.forEach((a, i) => {
    const x = pecas.find(p => chaveDe(p) === chaveDe({ nome: a.nome, desc: a.desc, unidade: 'desc' }));
    if (!x) { console.log('  NAO ACHEI: ' + a.nome); return; }
    const antes = Math.round(Number(x.emEstoque) || 0);
    if (antes !== 0) { console.log('  ja tem ' + antes + ' em estoque, fica: ' + a.nome); return; }
    x.emEstoque = QTD; x.atualizadoEm = agora; x.atualizadoPor = 'Junior';
    movs.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      itemId: x.id, nome: x.nome, desc: x.desc || '', unidade: 'desc', data: hoje, obs: '',
      tipo: 'ajuste', motivo: 'correcao', qtd: 0, dUso: 0, dEstoque: QTD, por: 'Junior', em: agora });
    n++;
    console.log('  ' + a.nome + ': 0 -> ' + QTD);
  });
  console.log('\n' + n + ' agulha(s) a acertar.');
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-estoque-agulhas-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { pecasCad: JSON.stringify(pecas), pecasMov: JSON.stringify(movs) }), updatedAt);

  const { data: d2 } = await lerBlob(cab);
  const p2 = JSON.parse(d2.pecasCad), m2 = JSON.parse(d2.pecasMov);
  console.log('gravado. Agulhas com ' + QTD + ' em estoque: ' + p2.filter(p => /^Agulha /.test(p.nome) && p.emEstoque === QTD).length +
    ' · lancamentos no historico: ' + m2.length);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
