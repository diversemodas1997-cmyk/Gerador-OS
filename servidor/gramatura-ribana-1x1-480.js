/*
 * A GRAMATURA DA RIBANA MALHA ALGODAO PASSA DE 530 PARA 480.
 *
 * Junior, 11/09/2026: "grava 480 agora".
 *
 * DE ONDE VEM O 480
 *
 * A etiqueta da HTK vem com a gramatura zerada, entao o valor sempre foi
 * estimativa. Os 530 vinham de um catalogo SUPOSTO de 265 x 2. A pesquisa de
 * mercado de 11/09/2026 achou a ficha da composicao EXATA — ribana 1x1,
 * 98% algodao / 2% elastano — em 240 g/m². Como o campo conta as duas faces do
 * tubo, 240 x 2 = 480.
 *
 * A convencao do x2 nao e suposicao: quatro fichas tecnicas independentes
 * (Malharia Indaial, Austral, DN7, Emporio Stampe) publicam gramatura de uma
 * face + largura do tubo achatado, e o rendimento que elas mesmas informam so
 * fecha dobrando. Ex.: Austral, 280 g/m² e 0,55 m, informa 3,25 m/kg —
 * 1000/(280 x 0,55 x 2) = 3,25.
 *
 * O QUE ESTE SCRIPT NAO FAZ: NAO REESCREVE O RAZAO
 *
 * As baixas ja gravadas (origem 'os') guardam o kg calculado com 530 e ficam
 * como estao. O saldo do estoque NAO se mexe.
 *
 * Isso e de proposito. A baixa gravada e o que o programa disse na epoca, e
 * mexer nela mudaria o saldo de hoje por causa de uma conta de ontem — logo
 * depois de as prateleiras negativas terem sido zeradas. Quem quiser a OS antiga
 * recalculada e so salva-la de novo: `aplicarBaixaEstoqueOS` apaga e refaz os
 * movimentos daquela OS com a gramatura corrente.
 *
 * De hoje em diante, toda OS nova desse pano calcula com 480 — cerca de 9,4% a
 * menos de kg por enfesto.
 *
 * COMO RODAR
 *
 *   node servidor/gramatura-ribana-1x1-480.js             so relata
 *   node servidor/gramatura-ribana-1x1-480.js --gravar    grava no servidor
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const DE = 530, PARA = 480;

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

const OBS_NOVA =
  'A etiqueta da HTK vem com GRAMATURA E LARGURA ZERADAS (0 gr/m², 0 cm) — o sistema da ' +
  'fiação não preenche esses campos; os zeros não são medida. ' +
  'Gramatura 480 no campo = 240 g/m² da ficha de mercado × 2 faces do tubo. ' +
  'O 240 é o padrão para ribana 1x1 98% CO / 2% EL (pesquisa de 11/09/2026, ficha da ' +
  'Malharia Indaial na composição exata, com rendimento 2,42 m/kg que confirma o ×2). ' +
  'Valor anterior: 530, que vinha de um catálogo suposto de 265 — trocado em 11/09/2026. ' +
  'AINDA É FICHA DE MERCADO, NÃO MEDIÇÃO DESTE PANO. Para fechar sem estimativa: medir a ' +
  'metragem de um rolo. Um rolo de 9,37 kg dá 35 a 37 m se a gramatura for 240, e 31 a 33 m ' +
  'se for 265 (tubo de 53 a 56 cm). ' +
  'Largura fica em branco pelo mesmo motivo: a etiqueta não a informa. ' +
  'Fabricante do rolo: Fiação e Tinturaria Irmãos Assini Ltda.';

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const le = k => JSON.parse(data[k] || '[]');
  const tecidos = le('tecidos');

  const tec = tecidos.find(t => String(t.nome || '').trim().toLowerCase() === 'ribana malha algodão');
  if (!tec) throw new Error('nao achei o tecido "Ribana Malha Algodão"');

  console.log('');
  console.log('GRAMATURA DA RIBANA MALHA ALGODAO');
  console.log('');
  if (tec.peso === PARA) {
    console.log('  Ja esta em ' + PARA + ' g/m². Nada a fazer.');
    return;
  }
  if (tec.peso !== DE) {
    throw new Error('esperava encontrar ' + DE + ' g/m², achei ' + tec.peso + ' — confira antes de gravar');
  }

  tec.peso = PARA;
  tec.obs = OBS_NOVA;

  // Quanto isso muda dali para a frente, medido no que ja foi consumido.
  const mov = le('estoqueMov').filter(m => m.tecidoNome === tec.nome);
  const saida = mov.filter(m => m.tipo !== 'entrada').reduce((a, m) => a + (parseFloat(m.kg) || 0), 0);

  console.log('  ' + DE + ' g/m²  ->  ' + PARA + ' g/m²   (−' + ((1 - PARA / DE) * 100).toFixed(1) + '%)');
  console.log('');
  console.log('  O RAZAO NAO E REESCRITO. As ' + mov.length + ' movimentacoes ja gravadas');
  console.log('  guardam o kg calculado com ' + DE + ' (' + saida.toFixed(1) + ' kg de saida) e o saldo');
  console.log('  do estoque fica como esta. OS antiga salva de novo recalcula sozinha.');
  console.log('');
  console.log('  Para efeito de comparacao, se tudo fosse recalculado com ' + PARA + ':');
  console.log('    saida passaria de ' + saida.toFixed(1) + ' kg para ' + (saida * PARA / DE).toFixed(1) +
    ' kg  (' + (saida - saida * PARA / DE).toFixed(1) + ' kg a menos)');

  if (!GRAVAR) {
    console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.');
    return;
  }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-gramatura-480-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.tecidos = JSON.stringify(tecidos);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: gramatura ' + PARA + ' g/m².');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
