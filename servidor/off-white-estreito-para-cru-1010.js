/*
 * OFF-WHITE ESTREITO DA MALHA ALGODAO VIRA CRU.
 *
 * Junior, 10/10/2026: "Altere o nome da cor dos tecidos em estoque de tecidos,
 * de off-white para Cru, todas as medidas, exceto off-white 120 cm".
 *
 * Sao as entradas manuais de Malha Algodao, cor "Off-White Malha Algodao", com
 * largura diferente de 120 (80, 115, 117 e 119 cm). Passam para a cor do
 * cadastro "Cru malha algodao" (0054) -- largura, bobinas, kg, data e id ficam.
 * So a Malha Algodao: e o unico tecido com Cru no cadastro e com Off-White em
 * mais de uma largura.
 *
 * SO AS ENTRADAS, por escolha do Junior. O consumo das OS de grade 117 cm
 * (sem largura gravada) continua "Off-White Malha Algodao": sem a largura 117
 * lancada no Off-White, o quadro passa a contá-lo na linha de 120 cm.
 *
 *   node servidor/off-white-estreito-para-cru-1010.js            (so mostra)
 *   node servidor/off-white-estreito-para-cru-1010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');
const https = require('https');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'https://193.168.0.200';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const DE = 'Off-White Malha Algodão', PARA = 'Cru malha algodão';
const agente = new https.Agent({ rejectUnauthorized: false });

function req(method, caminho, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request(BASE + caminho, {
      method, agent: agente,
      headers: {
        apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json',
        Prefer: 'return=representation',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, (resp) => {
      let d = ''; resp.on('data', c => d += c);
      resp.on('end', () => resp.statusCode < 300 ? res(d) : rej(new Error('HTTP ' + resp.statusCode + ': ' + d.slice(0, 300))));
    });
    r.on('error', rej); if (data) r.write(data); r.end();
  });
}

const arr = s => (typeof s === 'string' ? JSON.parse(s || '[]') : (s || []));
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const ehAlvo = m => m.tipo === 'entrada' && m.origem !== 'os'
  && norm(m.tecidoNome) === 'malha algodao' && norm(m.corNome) === norm(DE)
  && Number(m.largura) > 0 && Number(m.largura) !== 120;
const linha = m => `    ${m.data}  ${String(m.fechados || 0).padStart(2)} bob  ${String(m.largura).padStart(3)} cm  ${String(m.kg).padStart(8)} kg   ${m.corNome}`;

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  if (!arr(data.cores).some(c => c.nome === PARA)) throw new Error('a cor "' + PARA + '" nao existe no cadastro');
  const movs = arr(data.estoqueMov);
  const alvo = movs.filter(ehAlvo);
  if (!alvo.length) { console.log('  Nenhuma entrada Off-White fora de 120 cm -- nada a fazer.'); return; }
  const kg = l => Math.round(l.reduce((a, m) => a + (+m.kg || 0), 0) * 1000) / 1000;
  console.log(`ENTRADAS (${alvo.length}) -> "${PARA}":\n` + alvo.map(linha).join('\n') + `\n    total: ${kg(alvo)} kg`);

  const saida = movs.map(m => !ehAlvo(m) ? m : Object.assign({}, m, {
    corNome: PARA,
    obs: [m.obs, 'Era "' + m.corNome + '" -- passou para Cru em 10/10/2026'].filter(Boolean).join(' · ')
  }));
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-off-white-para-cru-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));

  const novo = Object.assign({}, data, {
    estoqueMov: typeof data.estoqueMov === 'string' ? JSON.stringify(saida) : saida,
    _device: 'off-white-para-cru-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');

  const m2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.estoqueMov);
  const cru = m2.filter(m => m.tipo === 'entrada' && m.origem !== 'os' && m.corNome === PARA);
  console.log(`gravado. Agora: ${m2.filter(ehAlvo).length} entradas Off-White fora de 120 cm; `
    + `${cru.length} entradas Cru (${kg(cru)} kg); ${m2.length} movimentos no blob.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
