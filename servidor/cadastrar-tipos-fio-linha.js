/*
 * CADASTRA OS TIPOS DE FIO E DE LINHA DAS ETIQUETAS DOS CONES.
 *
 * Junior, 28/09/2026, sobre a foto "Etiquetas Fio e linhas": "Essas
 * informacoes devem servir para cadastro do tipo de linha e fio" e "As
 * etiquetas ilegiveis tem as mesmas informacoes das de mesmo tipo".
 *
 *   Fio   TP200  -- 4 etiquetas; cores 1500 e 9900 (as outras duas, rasgadas
 *                   no campo da cor, entram quando o Junior disser o codigo)
 *   Linha BR02C  -- 3 etiquetas; cores 000242, 000105 e 000154
 *
 * As cores entram so com o codigo: o nome vem do Junior depois (editar o
 * tipo na tela, "codigo = nome"). Nenhuma quantidade: isso e lancamento.
 * Tipo que ja existe (mesmo item + codigo) fica como esta.
 *
 *   node servidor/cadastrar-tipos-fio-linha.js            (so mostra)
 *   node servidor/cadastrar-tipos-fio-linha.js --gravar   (grava)
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

const TIPOS = [
  { item: 'Fio', codigo: 'TP200',
    desc: '100% poliéster 150 · Tex 27 · fabricante CNPJ 00.139.737/0006-02',
    cores: ['1500', '9900'] },
  { item: 'Linha', codigo: 'BR02C',
    desc: '100% poliéster · 120 · 1828 m · 29,00 Tex · Kalina Ind. Fios (CNPJ 20.555.875/0002-48)',
    cores: ['000242', '000105', '000154'] }
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
  let lista = [];
  try { lista = typeof data.aviamentoTipos === 'string' ? JSON.parse(data.aviamentoTipos) : (data.aviamentoTipos || []); } catch (e) { lista = []; }
  if (!Array.isArray(lista)) lista = [];
  const agora = new Date().toISOString();
  let n = 0;
  console.log('Tipos de fio e linha hoje: ' + lista.length + '.\n');
  TIPOS.forEach((t, i) => {
    if (lista.some(x => normNome(x.item) === normNome(t.item) && normNome(x.codigo) === normNome(t.codigo))) {
      console.log('  ja existe: ' + t.item + ' ' + t.codigo); return;
    }
    lista.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      item: t.item, codigo: t.codigo, desc: t.desc, cores: t.cores.map(c => ({ codigo: c, nome: '' })),
      atualizadoPor: 'Junior', atualizadoEm: agora });
    n++;
    console.log('  novo: ' + t.item + ' ' + t.codigo + ' -- ' + t.desc + ' -- cores ' + t.cores.join(', '));
  });
  console.log('\n' + n + ' tipo(s) a cadastrar.');
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-tipos-fio-linha-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(lista) }), updatedAt);

  const { data: d2 } = await lerBlob(cab);
  const l2 = JSON.parse(d2.aviamentoTipos);
  console.log('gravado. Tipos no cadastro: ' + l2.map(t => t.item + ' ' + t.codigo + ' (' + t.cores.length + ' cores)').join(', '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
