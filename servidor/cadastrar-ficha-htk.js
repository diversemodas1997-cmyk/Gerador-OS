/*
 * CADASTRA A FICHA DA ETIQUETA DA HTK (ribana 1x1, Ribana Malha Algodao).
 *
 * Junior, 11/09/2026: "determine as informacoes para completar o cadastro de
 * ribana malha algodao" — a partir da foto da etiqueta.
 *
 * O QUE A ETIQUETA DIZ
 *
 *   FIACAO E TINTURARIA IRMAOS ASSINI LTDA   (quem produziu o rolo)
 *   1992 - HTK COMERCIO DE TECIDOS LTDA      (o cliente 1992 da fiacao = quem vende pra casa)
 *   Cor 50827 01 - MARINHO                   (codigo antigo AZ0801)
 *   3009 - RIBANA 1X1 98% ALGODAO 2% ELASTANO
 *   NF/Romaneio 210 · O.S. 382449 · Peca 1/1 · Peso 9,37
 *   Largura 0cm · Gramatura 0gr/m²
 *
 * O QUE ESTE SCRIPT NAO FAZ
 *
 * NAO GRAVA A LARGURA NEM A GRAMATURA. Os zeros da etiqueta nao sao resposta:
 * sao campo em branco do sistema da fiacao. Gravar 0 na gramatura faria o kg de
 * TODO enfesto desse pano ir a zero.
 *
 * A gramatura fica nos 530 que ja estavam la — que sao ESTIMATIVA, nao ficha.
 * A pesquisa de mercado (11/09/2026) aponta 240 g/m² na etiqueta para ribana 1x1
 * 98% CO / 2% EL, o que daria 480 no campo (o campo conta as duas faces do tubo).
 * Trocar 530 por 480 move ~96 kg no razao acumulado, entao e decisao do Junior —
 * e a medicao da metragem de um rolo fecha a conta sem estimativa nenhuma.
 *
 * NAO GRAVA A COR DO ROLO. "MARINHO" nao existe no cadastro e pode ser a "Azul
 * Ribana Malha Algodao" (28 movimentos) com outro nome; inventar uma prateleira
 * nova e o erro que ja partiu 21 prateleiras em agosto. A linha entra sem cor, e
 * o rastro da NF, da partida e do peso fica registrado do mesmo jeito.
 *
 * NAO GRAVA A DATA DA ENTRADA. O carimbo 15/07/2026 e da IMPRESSAO da etiqueta,
 * nao da nota.
 *
 * COMO RODAR
 *
 *   node servidor/cadastrar-ficha-htk.js             so relata, nao grava
 *   node servidor/cadastrar-ficha-htk.js --gravar    grava no servidor
 *
 * Rodar duas vezes nao duplica nada.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const uid = () => 'id_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
const norm = s => String(s || '').trim().toLowerCase();

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

const OBS_TECIDO =
  'A etiqueta da HTK vem com GRAMATURA E LARGURA ZERADAS (0 gr/m², 0 cm) — o sistema da ' +
  'fiação não preenche esses campos; os zeros não são medida. ' +
  'A gramatura 530 no campo é ESTIMATIVA, não ficha técnica. ' +
  'Padrão de mercado (pesquisa de 11/09/2026) para ribana 1x1 98% CO / 2% EL: 240 g/m² na ' +
  'etiqueta, que dariam 480 no campo — o campo conta as duas faces do tubo. ' +
  'PARA FECHAR SEM ESTIMATIVA: medir a metragem de um rolo. Um rolo de 9,37 kg dá 35 a 37 m ' +
  'se a gramatura for 240, e 31 a 33 m se for 265 (tubo de 53 a 56 cm). ' +
  'Largura fica em branco pelo mesmo motivo: a etiqueta não a informa. ' +
  'Fabricante do rolo: Fiação e Tinturaria Irmãos Assini Ltda.';

const OBS_FORNECEDOR =
  'Não indenizamos peças talhadas. ' +
  'Código do cliente na fiação: 1992. ' +
  'Os rolos vêm produzidos pela Fiação e Tinturaria Irmãos Assini Ltda. ' +
  'As etiquetas deste fornecedor vêm com gramatura e largura zeradas.';

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const le = k => JSON.parse(data[k] || '[]');
  const fornecedores = le('fornecedores');
  const tecidos = le('tecidos');
  const feito = [];

  /* ---- 1. fornecedor HTK ---- */
  let htk = fornecedores.find(f => norm(f.nome).startsWith('htk'));
  if (!htk) {
    htk = {
      id: uid(),
      nome: 'HTK COMERCIO DE TECIDOS LTDA',
      cnpj: '',                // a etiqueta nao traz
      cidadeUf: 'São Paulo / SP',
      telefone: '',
      obs: OBS_FORNECEDOR
    };
    fornecedores.push(htk);
    feito.push('fornecedor HTK CRIADO');
  } else {
    feito.push('fornecedor HTK já existia');
  }

  /* ---- 2. tecido: so o que a etiqueta AFIRMA ---- */
  const tec = tecidos.find(t => norm(t.nome) === 'ribana malha algodão');
  if (!tec) throw new Error('nao achei o tecido "Ribana Malha Algodão"');
  const antes = JSON.stringify(tec);
  const pesoAntes = tec.peso;
  tec.estrutura = 'ribana-1x1';
  tec.desc = '98% CO 2% ELAS';
  tec.codigoFornecedor = '3009';
  tec.fornecedorId = htk.id;
  tec.tubular = 'tubular';
  tec.categoria = 'ribana';
  tec.obs = OBS_TECIDO;
  // peso e largura NAO sao tocados de proposito — ver o cabecalho.
  if (tec.peso !== pesoAntes) throw new Error('a gramatura nao devia ter mudado');
  feito.push(JSON.stringify(tec) === antes
    ? 'tecido Ribana Malha Algodão já estava completo'
    : 'tecido Ribana Malha Algodão COMPLETADO (artigo 3009, ribana 1x1)');

  /* ---- 3. entrada de rolo da NF 210 ---- */
  if (!Array.isArray(tec.entradasRolo)) tec.entradasRolo = [];
  const mesmo = r => r.nf === '210' && r.partida === '382449';
  if (!tec.entradasRolo.some(mesmo)) {
    tec.entradasRolo.push({
      id: uid(),
      nf: '210',              // NF/Romaneio
      data: '',               // o carimbo da etiqueta e da impressao, nao da nota
      lote: '',               // "Lote Cliente" veio vazio
      partida: '382449',      // a O.S. da fiacao
      rolo: '1 de 1',         // Peca 1/1
      peso: '9,37',
      corId: '',              // MARINHO ainda nao resolvido no cadastro
      corNome: ''
    });
    feito.push('entrada de rolo NF 210 / partida 382449 CRIADA');
  } else {
    feito.push('entrada de rolo NF 210 / partida 382449 já existia');
  }

  console.log('');
  console.log('FICHA DA ETIQUETA DA HTK');
  console.log('');
  feito.forEach(f => console.log('  · ' + f));
  console.log('');
  console.log('  tecido    : ' + tec.nome + '  [' + tec.codigoFornecedor + ' · ribana 1x1]');
  console.log('              ' + tec.desc + ' · tubular · categoria ' + tec.categoria);
  console.log('  gramatura : ' + tec.peso + ' g/m²  (INALTERADA — estimativa; a etiqueta traz 0)');
  console.log('  largura   : ' + (tec.largura ? tec.largura + ' cm' : 'em branco (a etiqueta traz 0)'));
  console.log('  rolos     : ' + tec.entradasRolo.length);

  if (!GRAVAR) {
    console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.');
    return;
  }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-ficha-htk-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.fornecedores = JSON.stringify(fornecedores);
  data.tecidos = JSON.stringify(tecidos);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
