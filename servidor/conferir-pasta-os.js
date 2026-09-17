/* CONFERE A PASTA DE PDFs CONTRA A LISTA DE OS DO SERVIDOR. Somente leitura.
 *
 * POR QUE ISTO EXISTE (17/09/2026)
 * A pasta e a TESTEMUNHA. Quando uma OS some do programa, o PDF dela continua
 * la — o arquivo e escrito no disco da maquina, e nao no blob que se perdeu.
 * Foi assim que a 0509, a 0536 e a 0537 foram descobertas, e e assim que a
 * proxima vai ser: PDF na pasta sem OS no programa e perda, com prova.
 *
 * O app tem a mesma conferencia embutida (ver abrirNumerosFaltandoOS), mas ela
 * so funciona na maquina que tem a pasta conectada. Esta aqui roda no servidor,
 * de fora, e serve para rodar de vez em quando ou junto do backup.
 *
 * USO:
 *   node servidor/conferir-pasta-os.js
 *   node servidor/conferir-pasta-os.js --pasta "D:\outro\caminho"
 */
const fs = require('fs');
const https = require('https');
const iP = process.argv.indexOf('--pasta');
const PASTA = iP > 0 ? process.argv[iP + 1] : 'J:/Meu Drive/Ordens de Serviço/Ordens de Serviço_Setor Corte';
const KEY = (fs.readFileSync('C:\\supabase\\docker\\.env', 'utf8').match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m) || [])[1].trim();
const agente = new https.Agent({ rejectUnauthorized: false });

function get(caminho) {
  return new Promise((res, rej) => {
    const u = new URL('https://193.168.0.200' + caminho);
    https.request({ hostname: u.hostname, port: 443, path: u.pathname + u.search, method: 'GET', agent: agente,
      headers: { apikey: KEY, Authorization: 'Bearer ' + KEY } }, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => r.statusCode < 300 ? res(d) : rej(new Error('HTTP ' + r.statusCode)));
    }).on('error', rej).end();
  });
}

(async () => {
  const rows = JSON.parse(await get('/rest/v1/shared_data?id=eq.main&select=data,updated_at'));
  const d = rows[0].data;
  const ordens = typeof d.ordens === 'string' ? JSON.parse(d.ordens) : d.ordens;
  const num = o => { const n = parseInt(String(o.os || '').replace(/\D/g, ''), 10); return Number.isNaN(n) ? null : n; };
  const noPrograma = new Set(ordens.map(num).filter(n => n > 0));

  const naPasta = new Map();
  for (const nome of fs.readdirSync(PASTA)) {
    const m = /^OS-0*(\d{1,4})(?:-\d{2}-\d{2}-\d{4})?\.pdf$/i.exec(nome);
    if (!m) continue;
    const n = parseInt(m[1], 10);
    const st = fs.statSync(PASTA + '/' + nome);
    if (!naPasta.has(n) || st.mtime > naPasta.get(n).quando) naPasta.set(n, { nome, quando: st.mtime });
  }

  const orfaos = [...naPasta.keys()].filter(n => !noPrograma.has(n)).sort((a, b) => a - b);
  const semPdf = [...noPrograma].filter(n => !naPasta.has(n)).sort((a, b) => a - b);

  console.log(`programa: ${noPrograma.size} OS   ·   pasta: ${naPasta.size} folhas em PDF\n`);
  console.log(`PDF NA PASTA SEM OS NO PROGRAMA (perdidas): ${orfaos.length}`);
  orfaos.forEach(n => {
    const f = naPasta.get(n);
    console.log(`   ${f.nome}   escrito em ${f.quando.toISOString().slice(0, 16).replace('T', ' ')}`);
  });
  console.log(`\nOS no programa sem PDF na pasta: ${semPdf.length}`);
  if (semPdf.length) {
    const faixas = [];
    semPdf.forEach(n => { const u = faixas[faixas.length - 1]; if (u && n === u.ate + 1) u.ate = n; else faixas.push({ de: n, ate: n }); });
    console.log('   ' + faixas.map(f => f.de === f.ate ? String(f.de).padStart(4, '0')
      : String(f.de).padStart(4, '0') + '-' + String(f.ate).padStart(4, '0')).join(', '));
  }
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
