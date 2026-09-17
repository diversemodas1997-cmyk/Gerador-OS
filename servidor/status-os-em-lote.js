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
     node servidor/status-os-em-lote.js --de 186 --ate 485 --status estoque
     node servidor/status-os-em-lote.js --de 186 --ate 485 --status estoque --gravar

   Ou por LISTA, quando as OS não formam uma faixa, com um motivo junto:
     node servidor/status-os-em-lote.js --os 292,293,296 --status cancelado \
       --nota "lote desistido"

   Os status válidos são os da tabela STATUS_OS do app.js — lidos de lá a cada
   rodada, não copiados para cá.
*/
const fs = require('fs');
const path = require('path');
const https = require('https');

/* OS ESTADOS SAIEM DO app.js, e não de uma cópia aqui (17/09/2026).

   Estavam escritos à mão — `['nao-iniciado','andamento','parado','finalizado']`
   — e ficaram três anos-luz atrás da tela: em 15/09 o status virou a ETAPA
   (Enfestando, Cortando, Costurando | São Carlos…), "andamento" e "finalizado"
   deixaram de existir, e este script continuava oferecendo os quatro antigos e
   recusando os de verdade. Uma lista paralela é uma lista que sai do lugar.

   Lê da tabela STATUS_OS do app.js, do início da declaração até o primeiro
   ponto e vírgula — a mesma técnica dos testes, e o mesmo motivo de não haver
   ponto e vírgula nos comentários daquela tabela. */
const ESTADOS = (() => {
  const appjs = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const m = appjs.match(/^const STATUS_OS = [^;]+;/m);
  if (!m) { console.error('não achei STATUS_OS no app.js'); process.exit(1); }
  const ks = [...m[0].matchAll(/\{\s*k:\s*'([a-z-]+)'/g)].map(x => x[1]);
  if (!ks.length) { console.error('STATUS_OS sem chaves legíveis'); process.exit(1); }
  return ks;
})();

const args = process.argv.slice(2);
const opt = (n) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : null; };
const DE = parseInt(opt('--de'), 10);
const ATE = parseInt(opt('--ate'), 10);
/* Uma LISTA de números, em vez de uma faixa: "--os 292,293,296". A faixa serve
   para carimbar o passado inteiro de uma vez; a lista, para as poucas OS que
   têm alguma coisa em comum que a numeração não expressa. */
const LISTA = (opt('--os') || '').split(',').map(x => parseInt(x, 10)).filter(n => n > 0);
const STATUS = (opt('--status') || '').trim();
const POR = (opt('--por') || 'admin@diverse.local').trim().toLowerCase();
/* Uma observação junto do carimbo. Mudança de status em lote quase sempre tem
   um MOTIVO que não cabe no status — e o motivo é o que faz sentido para quem
   abrir a OS daqui a um ano. Vai assinada por --por, como qualquer nota. */
const NOTA = (opt('--nota') || '').trim();
const GRAVAR = args.includes('--gravar');
const BASE = opt('--url') || 'https://193.168.0.200';
const ENV = opt('--env') || 'C:\\supabase\\docker\\.env';

if (!LISTA.length && (isNaN(DE) || isNaN(ATE))) {
  console.error('faltou --de <numero> e --ate <numero>, ou --os <n,n,n>');
  process.exit(1);
}
if (!LISTA.length && DE > ATE) { console.error('--de é maior que --ate'); process.exit(1); }
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

  const naFaixa = LISTA.length
    ? ordens.filter(o => LISTA.indexOf(numeroOS(o)) >= 0)
    : ordens.filter(o => { const n = numeroOS(o); return n !== null && n >= DE && n <= ATE; });
  const jaEstavam = naFaixa.filter(o => String(o.statusOS || 'nao-iniciado') === STATUS).length;
  /* O status ANTIGO tem de ser copiado agora, e não lido na hora de imprimir o
     aviso: o laço abaixo muda `o.statusOS` no PRÓPRIO objeto, e o aviso saía
     dizendo "0293(cancelado) será sobrescrita" — o valor novo, no lugar do que
     se vai perder. É justamente a linha que a pessoa lê para decidir se grava. */
  const outroStatus = naFaixa.filter(o => o.statusOS && o.statusOS !== STATUS)
    .map(o => ({ os: o.os, antes: o.statusOS }));
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
    if (NOTA) {
      if (!Array.isArray(o.obsNotas)) o.obsNotas = [];
      o.obsNotas.push({ login: POR, texto: NOTA, em: quando });
    }
    mudadas++;
  });

  const nums = naFaixa.map(numeroOS).sort((a, b) => a - b);
  console.log((LISTA.length ? `lista: OS ${LISTA.join(', ')}` : `faixa: OS ${DE} a ${ATE}`)
    + `  ·  status: ${STATUS}  ·  assinado por: ${POR}`);
  if (NOTA) console.log(`observação: ${NOTA}`);
  console.log(`OS encontradas: ${naFaixa.length}` + (nums.length ? ` (${nums.join(', ')})` : ''));
  console.log(`já estavam em "${STATUS}": ${jaEstavam}  ·  a mudar: ${mudadas}`);
  // As que tinham OUTRO status são as únicas em que se apaga um carimbo de
  // alguém. Vale ver os números antes de gravar.
  if (outroStatus.length) {
    console.log(`atenção: ${outroStatus.length} OS tinham outro status e serão sobrescritas: `
      + outroStatus.slice(0, 20).map(o => `${o.os}(${o.antes})`).join(', ')
      + (outroStatus.length > 20 ? ' …' : ''));
  }
  // Buraco na faixa é normal (OS excluída), mas dito em voz alta evita a
  // surpresa de "pedi 300 e mexeu em 280".
  const faltando = [];
  if (LISTA.length) LISTA.forEach(n => { if (!nums.includes(n)) faltando.push(n); });
  else for (let n = DE; n <= ATE; n++) if (!nums.includes(n)) faltando.push(n);
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
