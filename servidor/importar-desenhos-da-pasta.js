/* Carrega as imagens dos desenhos técnicos a partir de uma PASTA no disco,
   em vez de copiá-las da nuvem.

   POR QUE EXISTE
   As 25 imagens estão só no Storage da nuvem — nenhum backup local as tem
   (todos guardam apenas o endereço). Com a nuvem restrita, esta é a saída: você
   aponta a pasta onde as imagens originais estão e elas sobem direto para o
   servidor da fábrica.

   COMO ELE SABE QUAL IMAGEM É DE QUAL DESENHO
   Pelo SKU do desenho, casado com a pasta e o nome do arquivo — ver
   parear-desenhos.js, que é onde a lógica mora e onde o formato está explicado.
   Na dúvida ele NÃO pareia: um desenho sem imagem é um incômodo, o desenho
   errado numa OS de corte é tecido perdido.

   SEMPRE RODE PRIMEIRO COM --so-listar. Ele mostra o pareamento sem enviar
   nada, e é aí que se percebe um arquivo casado com o desenho errado.

   Uso:
     node servidor\importar-desenhos-da-pasta.js ^
       --pasta "Desenhos técnicos" ^
       --url http://localhost:8000 --key <SERVICE_ROLE_KEY> ^
       [--so-listar] [--mapa mapa.json]

   O --mapa resolve na mão o que o SKU não resolve: {"0013":"CM.REC/VERDE.png"}.
*/
const fs = require('fs');
const path = require('path');
const { parear } = require('./parear-desenhos');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const PASTA = arg('pasta');
const URL_ = (arg('url') || '').replace(/\/+$/, '');
const KEY = arg('key');
const MAPA = arg('mapa');
const SO_LISTAR = process.argv.includes('--so-listar');
if (!PASTA || !URL_ || !KEY) {
  console.error('Faltou --pasta, --url ou --key. Veja o cabeçalho deste arquivo.');
  process.exit(1);
}
if (!fs.existsSync(PASTA)) { console.error('Pasta não encontrada: ' + PASTA); process.exit(1); }

const EXT = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.webp': 'image/webp' };
const cab = { apikey: KEY, Authorization: 'Bearer ' + KEY };

// { 'CM.LISA': ['PRETO.png', ...] } — uma pasta por prefixo de SKU.
function lerPastas(raiz) {
  const porPasta = {};
  for (const nome of fs.readdirSync(raiz)) {
    const dir = path.join(raiz, nome);
    if (!fs.statSync(dir).isDirectory()) continue;
    const imgs = fs.readdirSync(dir).filter(f => EXT[path.extname(f).toLowerCase()]);
    if (imgs.length) porPasta[nome] = imgs;
  }
  return porPasta;
}

async function principal() {
  const linhas = await fetch(`${URL_}/rest/v1/shared_data?id=eq.main&select=data`, { headers: cab })
    .then(r => r.json());
  const blob = (linhas && linhas[0] && linhas[0].data) || null;
  if (!blob || !blob.desenhos) {
    console.error('shared_data ainda não foi migrado. Rode migrar-do-backup.js antes.');
    process.exit(1);
  }
  const desenhos = JSON.parse(blob.desenhos);
  const porPasta = lerPastas(PASTA);
  const mapaManual = MAPA ? JSON.parse(fs.readFileSync(MAPA, 'utf8')) : null;

  const { pares, pendencias, sobrando } = parear(desenhos, porPasta, mapaManual);
  const nArq = Object.values(porPasta).reduce((a, v) => a + v.length, 0);
  console.log(`Desenhos cadastrados: ${desenhos.length} | imagens na pasta: ${nArq}`);

  console.log(`\nPAREADOS (${pares.size})`);
  for (const d of desenhos) {
    const p = pares.get(String(d.codigo || '').trim());
    if (!p) continue;
    const marca = p.origem === 'manual' ? '  (mapa)' : '';
    console.log(`  ${String(d.codigo).padEnd(6)} ${String(d.skuLinha || '—').padEnd(18)} -> ${p.arquivo}${marca}`);
  }

  if (pendencias.length) {
    console.log(`\nPENDENTES (${pendencias.length}) — ficam como estão, nada é chutado`);
    pendencias.forEach(p =>
      console.log(`  ${String(p.cod).padEnd(6)} ${p.desc.slice(0, 40).padEnd(40)} ${p.motivo}`));
  }
  if (sobrando.length) {
    console.log(`\nARQUIVOS SEM DESENHO (${sobrando.length}) — modelos ainda não cadastrados`);
    sobrando.forEach(a => console.log('  ' + a));
  }

  if (SO_LISTAR) {
    console.log('\n--so-listar: nada foi enviado. Confira o pareamento e rode de novo sem ele.\n');
    return;
  }
  if (!pares.size) { console.log('\nNada a enviar.\n'); return; }

  console.log('\nEnviando…');
  let enviadas = 0;
  for (const d of desenhos) {
    const p = pares.get(String(d.codigo || '').trim());
    if (!p) continue;
    const arq = path.join(PASTA, p.arquivo.split('/').join(path.sep));
    const ext = path.extname(arq).toLowerCase();
    // Nome novo, no padrão do app. Não reaproveita o nome antigo de propósito:
    // se um dia a nuvem voltar, os dois lados ficam com arquivos distintos em
    // vez de um sobrescrever o outro pela metade.
    const nome = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}${ext}`;
    const r = await fetch(`${URL_}/storage/v1/object/desenhos/${nome}`, {
      method: 'POST',
      headers: Object.assign({}, cab, { 'Content-Type': EXT[ext], 'x-upsert': 'true' }),
      body: fs.readFileSync(arq)
    });
    if (!r.ok) { console.log(`  FALHOU ${d.codigo}: ${r.status} ${await r.text()}`); continue; }
    d.img = nome;                 // só o nome: o app monta o endereço (urlDesenho)
    enviadas++;
    process.stdout.write(`\r  ${enviadas}/${pares.size}`);
  }
  console.log('');
  if (!enviadas) { console.log('Nenhuma imagem subiu; os dados não foram alterados.\n'); return; }

  const r = await fetch(`${URL_}/rest/v1/shared_data?id=eq.main`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json' }),
    body: JSON.stringify({ data: Object.assign({}, blob, { desenhos: JSON.stringify(desenhos) }) })
  });
  if (!r.ok) {
    console.error('As imagens subiram, mas o vínculo NÃO foi gravado: ' + r.status + ' ' + await r.text());
    process.exit(1);
  }

  console.log(`\n✅ ${enviadas} desenho(s) com imagem no servidor da fábrica.`);
  if (pendencias.length) console.log(`   ${pendencias.length} pendente(s) — ver a lista acima.`);
  console.log('');
}

principal().catch(e => { console.error(e); process.exit(1); });
