/*
 * Tira o multiplicador de dezena do NOME dos PDFs de risco.
 *
 * POR QUE
 *
 * O encaixe da ribana leva dez grades no mesmo pano, e alguns arquivos ficaram
 * com esse dez no nome: "CM.LISA - RIBANA 10M-10G-10GG-10G1.pdf",
 * "RIBANA 10X P M G GG G1 G2 G3.pdf". Junior, em 20/08/2026: o nome tem de ser
 * o da GRADE — "se a grade e 2P-2M-...-2G3, a gola deve ter o mesmo nome,
 * conferindo que as quantidades sao equivalentes na base de 10".
 *
 * O nome do arquivo nao e enfeite: e por ele que a importacao adivinha a FASE
 * ("CORPO 2" -> Corpo Parte 2), e e ele que aparece na coluna Riscos, onde
 * alguem procura o risco da grade que tem na mao. Um "10M-10G" ali manda a
 * pessoa procurar uma grade que nao existe.
 *
 * O QUE ELE CONFERE ANTES DE RENOMEAR
 *
 * A EQUIVALENCIA NA BASE DE DEZ, lendo o proprio relatorio: a tabela de
 * tamanhos do PDF, dividida pela maior potencia de dez que couber, tem de dar
 * exatamente a distribuicao da PASTA em que ele esta. Nao batendo, o arquivo e
 * DEIXADO COMO ESTA e relatado — nome errado atrapalha, mas nome trocado por
 * adivinhacao aponta para a grade errada, que e pior.
 *
 * E ELE ARRUMA AS PONTAS SOLTAS
 *
 * `fase.risco` guarda o CAMINHO do PDF. Renomear sem mexer nele quebraria a
 * prova de origem da medida (o "✓ usado" some e a OS volta a avisar), entao o
 * script reaponta as fases que citam o arquivo antigo. E o indice
 * `dados/riscos-pdf.json` precisa ser regerado depois — ele avisa.
 *
 * COMO RODAR
 *
 *   node servidor/renomear-riscos-10x.js             so relata
 *   node servidor/renomear-riscos-10x.js --gravar    renomeia e reaponta
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const PASTA = path.join(RAIZ, 'Desenhos técnicos -grades de corte');
const VENDOR = path.join(RAIZ, 'vendor');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

// O nome que uma distribuicao pede, na convencao da casa — recortado do app.js.
const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
const solta = n => { const i = src.indexOf('function ' + n + '('); return src.slice(i, src.indexOf('\n}', i) + 2); };
const { nomeTam, chaveTam, formas } = new Function(`
  ${solta('_normNome')}
  ${solta('_riscoNomeTamanhos')}
  ${solta('_riscoFormasDoNome')}
  ${src.slice(src.indexOf('const _chaveTam ='), src.indexOf('\n', src.indexOf('const _chaveTam =')))}
  return { nomeTam: _riscoNomeTamanhos, chaveTam: _chaveTam, formas: _riscoFormasDoNome };
`)();

function pdfjs() {
  global.window = global;
  global.navigator = { userAgent: 'node' };
  const lib = require(path.join(VENDOR, 'pdf-3.11.174.min.js'));
  lib.GlobalWorkerOptions.workerSrc = path.join(VENDOR, 'pdf.worker-3.11.174.min.js');
  return lib;
}

// A tabela de tamanhos do relatorio (coluna "Completos").
async function tamanhosDoPdf(lib, arq) {
  const doc = await lib.getDocument({ data: new Uint8Array(fs.readFileSync(arq)) }).promise;
  const tc = await (await doc.getPage(1)).getTextContent();
  const itens = tc.items.filter(i => (i.str || '').trim())
    .map(i => ({ t: i.str.trim(), x: i.transform[4], y: i.transform[5] }));
  const filas = [];
  itens.slice().sort((a, b) => b.y - a.y).forEach(i => {
    const f = filas[filas.length - 1];
    if (f && Math.abs(f.y - i.y) <= 3) { f.itens.push(i); return; }
    filas.push({ y: i.y, itens: [i] });
  });
  filas.forEach(f => f.itens.sort((a, b) => a.x - b.x));
  const txt = f => f.itens.map(i => i.t).join(' ');
  const iCab = filas.findIndex(f => /(^|\s)Tamanho(\s|$)/i.test(txt(f)) && /Completos/i.test(txt(f)));
  const tam = {};
  if (iCab < 0) return tam;
  const iFim = filas.findIndex((f, k) => k > iCab && /^Encaixe$/i.test(txt(f).trim()));
  const ate = iFim > iCab ? iFim : filas.length;
  for (let k = iCab + 1; k < ate; k++) {
    const toks = filas[k].itens;
    const iT = toks.findIndex(t => /^(P|M|G|GG|G1|G2|G3)$/i.test(t.t));
    if (iT < 0) continue;
    const ns = toks.slice(iT + 1).filter(t => /^\d+$/.test(t.t));
    const n = ns[0] ? parseInt(ns[0].t, 10) : 0;
    if (n > 0) tam[toks[iT].t.toLowerCase()] = n;
  }
  return tam;
}

const KEYS = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3'];
// A maior potencia de dez que divide todas as quantidades — a mesma regra que a
// importacao usa (_riscoDivisorDoGrupo no app.js).
function divisorBaseDez(tam) {
  const qs = KEYS.map(k => parseInt(tam[k], 10) || 0).filter(n => n > 0);
  if (!qs.length) return 0;
  let d = 1;
  while (qs.every(n => n % (d * 10) === 0)) d *= 10;
  return d;
}
const dividir = (tam, d) => {
  const o = {};
  KEYS.forEach(k => { const n = parseInt(tam[k], 10) || 0; if (n > 0) o[k] = n / d; });
  return o;
};

function todos(dir, acc = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) todos(p, acc);
    else if (/\.pdf$/i.test(e.name)) acc.push(p);
  }
  return acc;
}

(async () => {
  const lib = pdfjs();
  const arqs = todos(PASTA);
  // Candidato: o nome do arquivo carrega uma dezena antes de um tamanho.
  const candidatos = arqs.filter(a => /(^|[\/ \-])(\d0)\s*X?\s*[PMG]/i.test(path.basename(a)));
  console.log(`${arqs.length} PDFs na pasta · ${candidatos.length} com dezena no nome`);
  console.log('');

  const trocas = [];
  for (const arq of candidatos) {
    const rel = path.relative(PASTA, arq).split(path.sep).join('/');
    const base = path.basename(arq);
    const daPasta = rel.split('/')[1] || '';            // <LINHA>/<TAMANHOS>/...
    const tam = await tamanhosDoPdf(lib, arq);
    const d = divisorBaseDez(tam);
    const lido = KEYS.filter(k => tam[k]).map(k => k.toUpperCase() + '=' + tam[k]).join(' ');
    if (d < 10) {
      console.log(`  = ${base}\n      relatorio ${lido || '(sem tabela)'} — nao e multiplo de dez, DEIXADO COMO ESTA`);
      continue;
    }
    const daGrade = dividir(tam, d);
    const pedido = nomeTam(daGrade);
    // A CONFERENCIA: o que o relatorio diz, dividido, tem de ser a pasta. Vale
    // pelas DUAS grafias — "P ao G3" e "P-M-G-GG-G1-G2-G3" sao a mesma
    // distribuicao, e a casa usa as duas (ver _riscoFormasDoNome).
    if (!formas(daGrade).concat([pedido]).some(f => chaveTam(f) === chaveTam(daPasta))) {
      console.log(`  ! ${base}\n      relatorio ${lido} ÷ ${d} = ${pedido}, mas a pasta e ${daPasta} — DEIXADO COMO ESTA`);
      continue;
    }
    // O nome novo leva o NOME DA PASTA, que e o da grade — foi o pedido:
    // "se a grade e 2P-2M-...-2G3, a gola deve ter o mesmo nome". O resto do
    // nome fica de pe (o prefixo da linha, a palavra RIBANA, o "-v1"): quem
    // batizou sabia o que escrevia, e o que esta errado e so o multiplicador.
    const novoBase = base
      .replace(/(\d0)\s*X\s*[PMG][^.]*?(?=\s*(-v\d)?\.pdf$)/i, daPasta)
      .replace(/(\d0[PMG][A-Z0-9]*)(-\d0[PMG][A-Z0-9]*)*/i, daPasta);
    if (novoBase === base) {
      console.log(`  ? ${base}\n      nao consegui reescrever o nome — DEIXADO COMO ESTA`);
      continue;
    }
    // Nome novo ja ocupado: parar. Dois arquivos com o mesmo nome na mesma
    // pasta e o comeco de uma confusao que ninguem desfaz depois.
    if (fs.existsSync(path.join(path.dirname(arq), novoBase))) {
      console.log(`  ! ${base}
      ja existe um "${novoBase}" nesta pasta — DEIXADO COMO ESTA`);
      continue;
    }
    trocas.push({ arq, rel, base, novoBase, relNovo: rel.replace(/[^/]+$/, novoBase), pedido, lido, d });
    console.log(`  → ${base}\n      ${lido} ÷ ${d} = ${pedido} (= pasta ${daPasta})\n      vira: ${novoBase}`);
  }

  console.log('');
  console.log(`a renomear: ${trocas.length}`);
  if (!trocas.length) return;

  // As fases que apontam para os arquivos antigos.
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
  const data = linhas[0].data, updatedAt = linhas[0].updated_at;
  const grades = JSON.parse(data.grades || '[]');
  const reaponta = [];
  grades.forEach(g => (g.fases || []).forEach(f => {
    const t = trocas.find(x => String(f.risco || '') === x.rel);
    if (t) reaponta.push({ g, f, t });
  }));
  console.log(`fases que apontam para eles e serao reapontadas: ${reaponta.length}`);
  reaponta.forEach(r => console.log(`   ${r.g.nome} · ${r.f.nome}`));

  if (!GRAVAR) { console.log('\n(nada foi feito — rode com --gravar para valer)'); return; }

  // Primeiro o cadastro: se a gravacao falhar, os arquivos continuam onde as
  // fases dizem que estao. Renomeando antes, uma falha aqui deixaria o cadastro
  // apontando para arquivo que nao existe mais.
  if (reaponta.length) {
    reaponta.forEach(r => { r.f.risco = r.t.relNovo; });
    data.grades = JSON.stringify(grades);
    const r = await fetch(
      `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
        method: 'PATCH',
        headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
        body: JSON.stringify({ data, updated_at: new Date().toISOString() })
      });
    const volta = await r.json();
    if (!r.ok || !Array.isArray(volta) || !volta.length) {
      throw new Error('nao consegui reapontar as fases no servidor — NADA foi renomeado. '
        + (r.ok ? 'Alguem gravou no meio do caminho; rode de novo.' : JSON.stringify(volta).slice(0, 200)));
    }
    console.log('\ncadastro reapontado.');
  }
  trocas.forEach(t => {
    fs.renameSync(t.arq, path.join(path.dirname(t.arq), t.novoBase));
  });
  console.log(`${trocas.length} arquivo(s) renomeado(s).`);
  console.log('\nAgora rode:  node servidor/indexar-riscos.js');
})().catch(e => { console.error('ERRO: ' + (e && e.message || e)); process.exit(1); });
