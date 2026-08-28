/*
 * Leva a entrada de ABERTURA DE ESTOQUE para a prateleira onde o consumo dela
 * foi parar — e a redimensiona para casar com ele.
 *
 * POR QUE ISTO EXISTE
 *
 * Em 27/08/2026 a abertura do estoque criou uma entrada manual por prateleira
 * encontrada no razao, cada uma do tamanho exato do consumo que ja estava
 * lancado. A frase que ela deixou gravada diz o proposito: "Saldo inicial —
 * abertura do estoque (o consumo das OS ja lancadas nao tinha entrada
 * correspondente)". Ela existe para casar com consumo antigo, e nada mais.
 *
 * Em 28/08 as OS 0483 e 0484 mudaram de pano — estavam em `Texturizado | Rugao`
 * por causa de uma grade copiada, e passaram para `Texturizado | Jacguar`, que
 * e o que os desenhos 0032/0033 sempre disseram. O consumo mudou de prateleira;
 * a abertura ficou onde estava. Resultado: 20 kg de Rugao numa prateleira que
 * nao existe (Rugao numa cor Jacguar, Rugao numa cor Prime) e o Jaguar devendo
 * 18,5 kg.
 *
 * Junior, 28/08/2026: "move a abertura junto, tudo zera."
 *
 * COMO ELE DECIDE
 *
 * ORFA e a entrada de abertura cuja prateleira nao tem consumo NENHUM de OS —
 * o pano que ela cobria foi embora. DESTINO e a prateleira que tem consumo de
 * OS e nenhuma abertura — o pano que chegou sem cobertura. O par se faz pela
 * COR BASE (a primeira palavra: "Preto Jacguar" e "Preto Rugao" sao os dois
 * "preto"), e so quando ha exatamente UM destino para aquela cor. Mais de um, ou
 * nenhum, e RELATADO e pulado: mover no escuro poe quilo na prateleira errada,
 * que e pior do que a linha estranha visivel.
 *
 * O KG NAO E O DA ORFA, E O DO DESTINO. O pano mudou de gramatura junto com o
 * tecido (Jaguar 250 g/m2, Rugao 270), entao repetir os 8,902 kg deixaria a
 * prateleira nova com sobra. A abertura existe para ZERAR o consumo; entao ela
 * vale o consumo.
 *
 * A entrada e reescrita NO LUGAR: mesmo id, mesma data, mesma origem. Apagar e
 * criar outra faria o historico mostrar uma entrada nova hoje, e nao houve
 * compra nenhuma hoje — o pano e o mesmo de sempre, so mudou de nome.
 *
 * O QUE ELE NAO FAZ
 *
 * Nao mexe em prateleira que apenas nao esta zerada. Depois de 27/08 a fabrica
 * continuou produzindo: sobra e falta sao o estoque de verdade se mexendo, e
 * zera-las seria apagar o que aconteceu. So a ORFA — abertura sem consumo
 * algum — e tratada aqui.
 *
 * COMO RODAR
 *
 *   node servidor/mover-abertura-orfa.js             so relata
 *   node servidor/mover-abertura-orfa.js --gravar    grava no servidor
 *
 * Antes de gravar salva o blob em backups/, e a escrita confere o carimbo.
 *
 * ATENCAO: recarregue a pagina (F5) depois de rodar.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

// A frase que a abertura de 27/08 deixou gravada em cada entrada dela.
const ABERTURA = /Saldo inicial . abertura do estoque/;
const norm = s => String(s || '').trim().toLowerCase();
// "Preto Jacguar" e "Preto Rugao" sao os dois "preto": e a cor da peca, sem o
// pano. A primeira palavra basta porque o cadastro nomeia sempre "Base Tecido".
const corBase = s => norm(s).split(' ')[0];

async function lerBlob() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at, cab };
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

(async () => {
  const { data, updatedAt, cab } = await lerBlob();
  const estoqueMov = JSON.parse(data.estoqueMov || '[]');

  // Por prateleira: quanto de abertura, quanto de consumo de OS.
  const chave = m => (m.tecidoNome || '') + '||' + (m.corNome || '');
  const prat = new Map();
  estoqueMov.forEach(m => {
    const k = chave(m);
    const v = prat.get(k) || { tecidoNome: m.tecidoNome || '', corNome: m.corNome || '', ab: 0, os: 0 };
    const kg = parseFloat(m.kg) || 0;
    if (m.origem === 'manual' && ABERTURA.test(m.obs || '')) v.ab += kg;
    else if (m.origem === 'os') v.os += kg;
    prat.set(k, v);
  });

  const orfas = [...prat.values()].filter(v => v.ab > 0 && v.os === 0);
  const destinos = [...prat.values()].filter(v => v.os > 0 && v.ab === 0);

  if (!orfas.length) { console.log('Nenhuma entrada de abertura orfa.'); return; }

  let movidas = 0;
  orfas.forEach(o => {
    const casam = destinos.filter(d => corBase(d.corNome) === corBase(o.corNome));
    if (casam.length !== 1) {
      console.log('  PULADA: ' + o.tecidoNome + ' :: ' + o.corNome
        + '  (' + casam.length + ' destinos com a cor "' + corBase(o.corNome) + '")');
      return;
    }
    const d = casam[0];
    const linhas = estoqueMov.filter(m => m.origem === 'manual' && ABERTURA.test(m.obs || '')
      && (m.tecidoNome || '') === o.tecidoNome && (m.corNome || '') === o.corNome);
    if (linhas.length !== 1) {
      console.log('  PULADA: ' + o.tecidoNome + ' :: ' + o.corNome
        + '  (' + linhas.length + ' entradas de abertura nessa prateleira)');
      return;
    }
    const m = linhas[0];
    console.log('  ' + o.tecidoNome + ' :: ' + o.corNome + '   ' + m.kg + ' kg');
    console.log('     ->  ' + d.tecidoNome + ' :: ' + d.corNome + '   '
      + (Math.round(d.os * 1000) / 1000) + ' kg   (o tamanho do consumo que ela cobre)');
    m.tecidoNome = d.tecidoNome;
    m.corNome = d.corNome;
    m.kg = Math.round(d.os * 1000) / 1000;
    if (!/movida junto com o consumo/.test(m.obs || '')) {
      m.obs = (m.obs || '') + ' · movida junto com o consumo, que trocou de pano em 28/08/2026';
    }
    movidas++;
    // Um destino so recebe uma vez.
    destinos.splice(destinos.indexOf(d), 1);
  });

  console.log('');
  console.log('entradas movidas: ' + movidas);
  if (!movidas) { console.log('Nada a gravar.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-abertura-orfa-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('copia de seguranca: ' + path.relative(RAIZ, arq));

  data.estoqueMov = JSON.stringify(estoqueMov);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
