/*
 * CADASTRA A FICHA DA ETIQUETA DA PLUMA (ribana 2x1 PA, preto enxofre).
 *
 * Junior, 11/09/2026: "insira os registros abaixo como dados iniciais, sem
 * duplicar se ja existirem."
 *
 * O QUE ELE FAZ — E O QUE ELE NAO DUPLICA
 *
 * 1. FORNECEDOR Pluma. Nao existia nenhum fornecedor cadastrado; nasce aqui.
 *
 * 2. TECIDO. O pano da etiqueta — Ribana 2x1 PA, 48% CO / 48% PES / 4% ELAS,
 *    tubo de 60 cm, da Pluma — JA ESTA cadastrado, com o nome de casa
 *    "Ribana Moletom". Este script COMPLETA a ficha dele em vez de abrir um
 *    segundo tecido: 47 OSs, as grades e o razao do estoque penduram no nome
 *    atual, e um segundo cadastro partiria a prateleira em duas.
 *
 *    A GRAMATURA FICA EM 660, nao nos 330 da etiqueta. A largura cadastrada nas
 *    fases e a do TUBO ACHATADO, entao a gramatura tem de contar as duas faces:
 *    330 x 2. Gravar 330 faria o kg de todo enfesto desse pano sair pela metade
 *    — foi exatamente o que aconteceu com a malha em agosto. Ver
 *    project_gramatura_tecido.
 *
 *    A CATEGORIA FICA EM "ribana", nao em "malha algodao". O limite pedido (80
 *    folhas) e o mesmo nos dois (LIMITE_CAMADAS.ribana === 80), e "ribana" e o
 *    que liga o resto: Unidades da grade, bobinas por fase, a conta de camadas.
 *    Alem disso `categoriaEfetivaTecido` forca "ribana" para qualquer nome que
 *    contenha "ribana" — trocar o campo mudaria a ficha e nao mudaria a conta.
 *
 * 3. COR. "PRETO ENXOFRE 020104" e como a PLUMA chama o preto dessa ribana; o
 *    nome de casa "Preto Ribana Moletom" ja existe e e usado por 47 OSs. O
 *    script preenche o codigo e o nome DO FORNECEDOR na cor que ja existe, sem
 *    tocar no codigo Linx (001) nem no nome de casa — que e quem amarra a cor ao
 *    tecido em corCanonicaPorTecido.
 *
 * 4. ENTRADA DE ROLO da NF 658. PESO LIQUIDO EM BRANCO: a etiqueta veio com 0, e
 *    vazio e pergunta, zero e resposta.
 *
 * COMO RODAR
 *
 *   node servidor/cadastrar-ficha-pluma.js             so relata, nao grava
 *   node servidor/cadastrar-ficha-pluma.js --gravar    grava no servidor
 *
 * Rodar duas vezes nao duplica nada: cada registro e procurado antes.
 * Antes de gravar salva o blob inteiro em backups/, e a escrita e
 * read-modify-write com conferencia de carimbo.
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
const soDigitos = s => String(s || '').replace(/\D/g, '');

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

const OBS_TECIDO = 'Padrão de mercado para ribana 2x1 PA: 280 a 330 g/m². ' +
  'Tolerância aceitável de ±5% na gramatura (313 a 347 g/m²) e ±2% na largura. ' +
  'Uso: punho, gola, barra e cós. ' +
  'Gramatura no campo = 660 porque o pano é tubular: 330 da etiqueta × 2 faces ' +
  '(a largura das fases é a do tubo achatado).';

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const le = k => JSON.parse(data[k] || '[]');
  const fornecedores = le('fornecedores');
  const tecidos = le('tecidos');
  const cores = le('cores');
  const feito = [];

  /* ---- 1. fornecedor ---- */
  let pluma = fornecedores.find(f =>
    soDigitos(f.cnpj) === '06906323000125' || norm(f.nome).startsWith('pluma'));
  if (!pluma) {
    pluma = {
      id: uid(),
      nome: 'PLUMA INDUSTRIA E COMERCIO DE TECIDOS',
      cnpj: '06.906.323/0001-25',
      cidadeUf: '',
      telefone: '',
      obs: 'Só indeniza malha em rolo, mediante apresentação da etiqueta.'
    };
    fornecedores.push(pluma);
    feito.push('fornecedor Pluma CRIADO');
  } else {
    feito.push('fornecedor Pluma já existia (' + pluma.nome + ')');
  }

  /* ---- 2. tecido: completa a ficha do que ja existe ---- */
  const tec = tecidos.find(t => norm(t.nome) === 'ribana moletom');
  if (!tec) throw new Error('nao achei o tecido "Ribana Moletom" — confira o cadastro antes de rodar');
  const antesTec = JSON.stringify(tec);
  tec.estrutura = 'ribana-2x1';
  tec.desc = '48% CO 48% PES 4% ELAS';
  tec.largura = 60;
  tec.rendimento = 2.59;         // da ficha do fornecedor; o calculo daria 2,53
  tec.tubular = 'tubular';
  tec.peso = 660;                // 330 da etiqueta x 2 faces — NAO baixar para 330
  tec.categoria = 'ribana';      // limite 80 camadas, igual ao pedido
  tec.fornecedorId = pluma.id;
  tec.codigoFornecedor = 'RIBANA 2X1 PA';
  tec.obs = OBS_TECIDO;
  feito.push(JSON.stringify(tec) === antesTec
    ? 'tecido Ribana Moletom já estava completo'
    : 'tecido Ribana Moletom COMPLETADO (artigo RIBANA 2X1 PA)');

  /* ---- 3. cor: o codigo e o nome DO FORNECEDOR na cor que ja existe ---- */
  const cor = cores.find(c => norm(c.nome) === 'preto ribana moletom');
  if (!cor) throw new Error('nao achei a cor "Preto Ribana Moletom" — confira o cadastro antes de rodar');
  const antesCor = JSON.stringify(cor);
  cor.fornecedorId = pluma.id;
  cor.tecidoId = tec.id;
  cor.codigoFornecedor = '020104';
  cor.nomeFornecedor = 'PRETO ENXOFRE';
  feito.push(JSON.stringify(cor) === antesCor
    ? 'cor Preto Ribana Moletom já estava completa'
    : 'cor Preto Ribana Moletom COMPLETADA (020104 · PRETO ENXOFRE)');

  /* ---- 4. entrada de rolo da NF 658 ---- */
  if (!Array.isArray(tec.entradasRolo)) tec.entradasRolo = [];
  const mesmoRolo = r => r.nf === '658' && r.lote === '1795' && r.partida === '04286' && r.rolo === '1 de 4';
  if (!tec.entradasRolo.some(mesmoRolo)) {
    tec.entradasRolo.push({
      id: uid(),
      nf: '658',
      data: '2026-05-14',
      lote: '1795',
      partida: '04286',
      rolo: '1 de 4',
      peso: '',                 // a etiqueta veio com 0 — fica em branco de propósito
      corId: cor.id,
      corNome: cor.nome
    });
    feito.push('entrada de rolo NF 658 / lote 1795 CRIADA');
  } else {
    feito.push('entrada de rolo NF 658 / lote 1795 já existia');
  }

  console.log('');
  console.log('FICHA DA ETIQUETA DA PLUMA');
  console.log('');
  feito.forEach(f => console.log('  · ' + f));
  console.log('');
  console.log('  tecido : ' + tec.nome + '  [' + tec.codigoFornecedor + ']');
  console.log('           ' + tec.desc + ' · ' + tec.peso + ' g/m² · tubo ' + tec.largura +
    ' cm · ' + tec.rendimento + ' m/kg · categoria ' + tec.categoria + ' (máx 80 camadas)');
  console.log('  cor    : ' + cor.nome + '  [Linx ' + (cor.codigo || '—') +
    ' · Pluma ' + cor.codigoFornecedor + ' ' + cor.nomeFornecedor + ']');
  console.log('  rolos  : ' + tec.entradasRolo.length);

  if (!GRAVAR) {
    console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.');
    return;
  }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-ficha-pluma-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.fornecedores = JSON.stringify(fornecedores);
  data.tecidos = JSON.stringify(tecidos);
  data.cores = JSON.stringify(cores);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
