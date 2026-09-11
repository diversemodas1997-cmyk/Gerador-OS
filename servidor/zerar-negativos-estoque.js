/*
 * ZERA OS SALDOS NEGATIVOS DO ESTOQUE DE TECIDOS.
 *
 * Junior, 11/09/2026: "cadastre a entrada dos tecidos do estoque de tecidos que
 * estao negativos, apenas, para zerar o estoque."
 *
 * Saldo negativo quer dizer que o pano foi consumido por OS antes de a compra
 * dele ter sido lancada — o razao tem a saida e nao tem a entrada. Este script
 * nao apaga saida nenhuma: ele lanca UMA ENTRADA MANUAL por prateleira
 * (tecido + cor) com exatamente o kg que falta, e o saldo daquela linha para em
 * zero. Prateleira com saldo positivo ou zero nao e tocada.
 *
 * A CONTA E A DO APP, RECORTADA
 *
 * Saldo = entradas (manuais + compras por NF da Contabilidade) − reservado − saida,
 * com a cor passada por corCanonicaPorTecido na leitura. Nao ha uma segunda
 * versao dessa conta aqui: as funcoes sao recortadas do app.js na hora de rodar,
 * como em corrigir-cor-cruzada.js. Se o app mudar de ideia sobre o saldo, este
 * script muda junto.
 *
 * As compras por NF entram na conta mas NAO no que e gravado — elas vivem na
 * tabela compras_materiais, nao no blob.
 *
 * COMO RODAR
 *
 *   node servidor/zerar-negativos-estoque.js             so relata, nao grava
 *   node servidor/zerar-negativos-estoque.js --gravar    grava no servidor
 *
 * Antes de gravar salva o blob inteiro em backups/, e a escrita e
 * read-modify-write com conferencia de carimbo: se alguem gravou entre a leitura
 * e a escrita, aborta em vez de passar por cima.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const HOJE = new Date().toISOString().slice(0, 10);
const OBS = 'Ajuste de abertura: entrada para zerar saldo negativo';

/* ---- a conta do saldo vem do app.js, recortada: uma so versao dela ---- */
function contaDoApp(STATE, comprasCache) {
  const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
  const solta = nome => {
    const i = src.indexOf('function ' + nome + '(');
    if (i < 0) throw new Error('nao achei ' + nome + ' no app.js');
    return src.slice(i, src.indexOf('\n}', i) + 2);
  };
  return new Function('STATE', 'comprasCache', [
    solta('_normNome'), solta('_sufixoTecidoNorm'),
    solta('corBaseNome'), solta('corCanonicaPorTecido'),
    solta('comprasComoMovimentos'), solta('movimentacoesEstoque'),
    solta('calcularSaldosEstoque'),
    'return { calcularSaldosEstoque };'
  ].join('\n'))(STATE, comprasCache);
}

/* ---- o servidor da fabrica ---- */
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
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at };
}

// As compras por NF sao ENTRADA no saldo. Se a tabela nao existir, segue sem elas
// — e o mesmo silencio do app (carregarComprasMateriais).
async function lerCompras(cab) {
  try {
    const r = await fetch(`${SUPA}/rest/v1/compras_materiais?select=*`, { headers: cab });
    if (!r.ok) return [];
    const d = await r.json();
    return Array.isArray(d) ? d : [];
  } catch (e) { return []; }
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
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
}

const uid = () => 'aj' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const kg3 = n => Math.round(n * 1000) / 1000;

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const compras = await lerCompras(cab);
  // Cada chave do blob e guardada como TEXTO JSON, nao como objeto.
  const le = k => JSON.parse(data[k] || '[]');
  const estoqueMov = le('estoqueMov');
  const STATE = { cores: le('cores'), tecidos: le('tecidos'), estoqueMov };
  const { calcularSaldosEstoque } = contaDoApp(STATE, compras);

  const { detalhe } = calcularSaldosEstoque();
  // Arredonda antes de comparar: sobra de centesimo de grama nao e falta de pano.
  const negativos = detalhe.filter(d => kg3(d.disponivel) < 0);

  console.log('');
  console.log('SALDOS NEGATIVOS  (' + negativos.length + ' de ' + detalhe.length + ' prateleiras)');
  console.log('');
  if (!negativos.length) { console.log('  Nenhum saldo negativo. Nada a fazer.'); return; }
  let falta = 0;
  negativos.forEach(d => {
    const f = -kg3(d.disponivel);
    falta += f;
    console.log('  ' + f.toFixed(3).padStart(10) + ' kg   ' +
      (d.tecidoNome || '—') + ' · ' + (d.corNome || 'sem cor'));
  });
  console.log('');
  console.log('  total a lancar como entrada: ' + kg3(falta).toFixed(3) + ' kg em ' +
    negativos.length + ' lancamento(s), data ' + HOJE + '.');

  const novos = negativos.map(d => ({
    id: uid(),
    tipo: 'entrada',
    tecidoNome: d.tecidoNome || '',
    corNome: d.corNome || '',
    kg: kg3(-d.disponivel),
    fechados: 0,
    abertos: 0,
    data: HOJE,
    origem: 'manual',
    osId: '',
    osNumero: '',
    obs: OBS
  }));

  // Confere pela propria conta do app: depois dos lancamentos, nenhuma linha
  // negativa e as zeradas param em zero.
  const STATE2 = { ...STATE, estoqueMov: estoqueMov.concat(novos) };
  const conferencia = contaDoApp(STATE2, compras).calcularSaldosEstoque();
  const sobrou = conferencia.detalhe.filter(d => kg3(d.disponivel) < 0);
  if (sobrou.length) {
    throw new Error('conferencia falhou: ainda sobraram ' + sobrou.length + ' linhas negativas');
  }

  if (!GRAVAR) {
    console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.');
    return;
  }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-zerar-negativos-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.estoqueMov = JSON.stringify(STATE2.estoqueMov);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + novos.length + ' entradas manuais.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
