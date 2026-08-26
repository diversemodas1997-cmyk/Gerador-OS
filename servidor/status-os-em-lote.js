/* Carimba o STATUS de uma FAIXA de OS de uma vez.

   POR QUE ISTO EXISTE
   O status da OS (não iniciado / em andamento / parado / finalizado) nasceu em
   26/08/2026 já com centenas de OS antigas na lista — todas em "não iniciado",
   que é a ausência do campo. Marcar 300 OS uma a uma no seletor da tela não é
   trabalho de ninguém; e a faixa antiga é justamente a que já acabou.

   Roda NO SERVIDOR da fábrica. Lê a chave de serviço do .env do Supabase local,
   lê o shared_data, carimba o status nas OS da faixa e grava de volta. É a mesma
   conta read-modify-write da restauração credenciada e do reatribuir-notas.

   O NÚMERO DA OS é lido como em numeroOSordenacao (app.js): só os dígitos, então
   "0186" e "186" são a mesma OS. OS sem número fica de fora.

   Antes de gravar, o blob inteiro é copiado para backups/ — é a volta atrás se
   a faixa sair errada.

   USO (confere primeiro, grava depois):
     node servidor/status-os-em-lote.js --de 186 --ate 485 --status finalizado
     node servidor/status-os-em-lote.js --de 186 --ate 485 --status finalizado --gravar
*/
const fs = require('fs');
const path = require('path');
const https = require('https');

// Os quatro estados são os do app.js (STATUS_OS). "nao-iniciado" aqui APAGA os
// campos, igual ao que o seletor da tela faz — é assim que uma OS que ninguém
// tocou não pesa no blob que desce inteiro a cada abertura.
const ESTADOS = ['nao-iniciado', 'andamento', 'parado', 'finalizado'];

const args = process.argv.slice(2);
const opt = (n) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : null; };
const DE = parseInt(opt('--de'), 10);
const ATE = parseInt(opt('--ate'), 10);
const STATUS = (opt('--status') || '').trim();
const POR = (opt('--por') || 'admin@diverse.local').trim().toLowerCase();
const GRAVAR = args.includes('--gravar');
const BASE = opt('--url') || 'https://193.168.0.200';
const ENV = opt('--env') || 'C:\\supabase\\docker\\.env';

if (isNaN(DE) || isNaN(ATE)) { console.error('faltou --de <numero> e/ou --ate <numero>'); process.exit(1); }
if (DE > ATE) { console.error('--de é maior que --ate'); process.exit(1); }
if (!ESTADOS.includes(STATUS)) {
  console.error('--status precisa ser um de: ' + ESTADOS.join(', '));
  process.exit(1);
}

const envTxt = fs.readFileSync(ENV, 'utf8');
const m = envTxt.match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m);
if (!m) { console.error('não achei SERVICE_ROLE_KEY em ' + ENV); process.exit(1); }
const KEY = m[1].trim();

// O certificado é o da própria fábrica; este script fala com o servidor local.
const agente = new https.Agent({ rejectUnauthorized: false });

function req(method, caminho, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const u = new URL(BASE + caminho);
    const r = https.request({
      hostname: u.hostname, port: u.port || 443, path: u.pathname + u.search, method, agent: agente,
      headers: {
        apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json',
        Prefer: method === 'PATCH' ? 'return=minimal' : '',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, (resp) => {
      let d = ''; resp.on('data', c => d += c);
      resp.on('end', () => resp.statusCode < 300 ? res({ status: resp.statusCode, body: d })
                                                 : rej(new Error('HTTP ' + resp.statusCode + ': ' + d.slice(0, 300))));
    });
    r.on('error', rej); if (data) r.write(data); r.end();
  });
}

function pk(v) { return typeof v === 'string' ? JSON.parse(v) : v; }
const numeroOS = (o) => {
  const n = parseInt(String((o && o.os) || '').replace(/\D/g, ''), 10);
  return Number.isNaN(n) ? null : n;
};

(async () => {
  const rows = JSON.parse((await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at')).body);
  if (!rows.length) { console.error('shared_data id=main não encontrado'); process.exit(1); }
  const data = rows[0].data || {};
  const ordensEraString = typeof data.ordens === 'string';
  const ordens = pk(data.ordens) || [];

  const naFaixa = ordens.filter(o => { const n = numeroOS(o); return n !== null && n >= DE && n <= ATE; });
  const jaEstavam = naFaixa.filter(o => String(o.statusOS || 'nao-iniciado') === STATUS).length;
  const outroStatus = naFaixa.filter(o => o.statusOS && o.statusOS !== STATUS);
  const quando = new Date().toISOString();

  let mudadas = 0;
  naFaixa.forEach(o => {
    if (String(o.statusOS || 'nao-iniciado') === STATUS) return;
    if (STATUS === 'nao-iniciado') { delete o.statusOS; delete o.statusOSPor; delete o.statusOSEm; }
    else { o.statusOS = STATUS; o.statusOSPor = POR; o.statusOSEm = quando; }
    // A data de finalização segue a mesma regra da tela (mudarStatusOS): entra
    // ao marcar "Finalizado", sai ao tirar. Aqui ela é o dia da RODADA — o dia
    // em que a faixa foi carimbada —, que é a única data que o script sabe.
    if (STATUS === 'finalizado') o.finalizadaEm = quando; else delete o.finalizadaEm;
    mudadas++;
  });

  const nums = naFaixa.map(numeroOS).sort((a, b) => a - b);
  console.log(`faixa: OS ${DE} a ${ATE}  ·  status: ${STATUS}  ·  assinado por: ${POR}`);
  console.log(`OS na faixa: ${naFaixa.length}` + (nums.length ? ` (da ${nums[0]} à ${nums[nums.length - 1]})` : ''));
  console.log(`já estavam em "${STATUS}": ${jaEstavam}  ·  a mudar: ${mudadas}`);
  // As que tinham OUTRO status são as únicas em que se apaga um carimbo de
  // alguém. Vale ver os números antes de gravar.
  if (outroStatus.length) {
    console.log(`atenção: ${outroStatus.length} OS tinham outro status e serão sobrescritas: `
      + outroStatus.slice(0, 20).map(o => `${o.os}(${o.statusOS})`).join(', ')
      + (outroStatus.length > 20 ? ' …' : ''));
  }
  // Buraco na faixa é normal (OS excluída), mas dito em voz alta evita a
  // surpresa de "pedi 300 e mexeu em 280".
  const faltando = [];
  for (let n = DE; n <= ATE; n++) if (!nums.includes(n)) faltando.push(n);
  if (faltando.length) {
    console.log(`números da faixa sem OS cadastrada: ${faltando.length}`
      + (faltando.length <= 30 ? ' (' + faltando.join(', ') + ')' : ''));
  }

  if (!GRAVAR) { console.log('\n(conferência — nada gravado. Rode de novo com --gravar para aplicar.)'); return; }
  if (!mudadas) { console.log('nada a gravar.'); return; }

  // Cópia do blob ANTES de mexer: é a volta atrás se a faixa sair errada.
  const dir = path.join(__dirname, '..', 'backups');
  fs.mkdirSync(dir, { recursive: true });
  const arq = path.join(dir, 'shared_data-antes-status-' + quando.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(rows[0]), 'utf8');
  console.log('cópia de segurança: ' + arq);

  data.ordens = ordensEraString ? JSON.stringify(ordens) : ordens;
  data._device = 'status-em-lote-' + Date.now();   // sentinela: todos os clientes recarregam
  await req('PATCH', '/rest/v1/shared_data?id=eq.main',
            { data, updated_at: new Date().toISOString() });
  console.log(`\n✅ ${mudadas} OS gravadas como "${STATUS}". Os clientes conectados recarregam sozinhos.`);
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
