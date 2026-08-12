/* Aplica a REGRA DO EXCEDENTE DE ENFESTO em todas as fases de todas as grades,
   direto no banco do servidor da fábrica.

   POR QUE ISTO EXISTE
   O caminho normal é o botão "Aplicar regra do excedente", em Configurações —
   ele mostra a prévia e grava o snapshot. Este script é a mesma coisa pela
   linha de comando, para quando o botão não puder ser usado.

   A REGRA NÃO É COPIADA: as funções são RECORTADAS do app.js de verdade
   (excedenteRegraDaFase e as que ela usa). Reescrever a regra aqui seria criar
   uma segunda verdade, que um dia diverge da primeira sem ninguém perceber —
   e o resultado de uma reescrita em massa é justamente o que ninguém confere
   linha a linha depois.

   Uso:
     node servidor\aplicar-excedente-regra.js --entrada <grades.json> \
          [--meta <meta.json>] --saida <grades-novo.json> [--aplicar]

   Sem --aplicar ele só relata o que MUDARIA (e não escreve a saída). É assim
   que se confere antes.

   --meta é a chave `meta` do blob, e é ONDE MORA A REGRA CADASTRADA
   (`meta.excedenteCfg`, editável em Configurações). Sem ela o script usa o
   padrão de fábrica — o que dá silenciosamente um resultado DIFERENTE do que
   o botão faria, se alguém tiver mudado as faixas na tela. Para pegá-la:

     docker exec supabase-db psql -U postgres -d postgres -At \
       -c "SELECT data->>'meta' FROM shared_data WHERE id='main';" > meta.json
*/
const fs = require('fs');
const path = require('path');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const ENTRADA = arg('entrada');
const SAIDA = arg('saida');
const META = arg('meta');
const APLICAR = process.argv.includes('--aplicar');
if (!ENTRADA) { console.error('Faltou --entrada <arquivo com o array de grades>.'); process.exit(1); }

// A regra CADASTRADA, se veio. O `meta` do blob é um objeto (às vezes gravado
// como string JSON, como o resto das chaves) — aceita-se as duas formas.
let meta = {};
if (META) {
  try {
    let m = JSON.parse(fsBoot().readFileSync(META, 'utf8').replace(/^﻿/, ''));
    if (typeof m === 'string') m = JSON.parse(m);
    meta = (m && typeof m === 'object') ? m : {};
  } catch (e) {
    console.error('nao consegui ler --meta: ' + ((e && e.message) || e));
    process.exit(1);
  }
}
function fsBoot() { return require('fs'); }

// ---------------------------------------------------------------- a regra
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + de + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + de); process.exit(1); }
  return src.slice(i, j + 2);
}
function achaConst(re, oQue) {
  const m = re.exec(src);
  if (!m) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return m[0];
}
const regra = new Function('STATE', [
  achaConst(/const EXCEDENTE_ENFESTO_PADRAO_CM = \d+;/, 'o padrao da casa'),
  achaConst(/const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/, 'a tabela de faixas'),
  achaConst(/const EXCEDENTE_GOLA_CM = \d+;/, 'o excedente da gola'),
  achaConst(/const EXCEDENTE_VIES_CM = \d+;/, 'o excedente do vies'),
  achaConst(/const _EXC_LIGACAO = new Set\(\[[^\]]*\]\);/, 'as palavras de ligacao'),
  achaConst(/const _PAL_VIES = new Set\(\[[^\]]*\]\);/, 'as palavras do vies'),
  achaConst(/const _PAL_GOLA = new Set\(\[[^\]]*\]\);/, 'as palavras da gola'),
  recorte('function _normNome'),
  recorte('function _normFaseNome'),
  recorte('function _faseSoDe'),
  recorte('function excedenteCfg'),
  recorte('function excedentePorComprimento'),
  recorte('function excedenteRegraDaFase'),
  'return { excedenteRegraDaFase, _faseSoDe, _PAL_VIES, _PAL_GOLA, excedenteCfg };'
].join('\n'))({ meta });

// Diz em voz alta qual regra vai valer. Rodar isto achando que usa a regra da
// tela, e usar o padrão de fábrica, é o erro que este aviso existe para evitar.
const cfgUsada = regra.excedenteCfg();
console.log('REGRA EM USO' + (meta.excedenteCfg ? ' (cadastrada em Configurações)' : ' (padrão de fábrica)') + ':');
cfgUsada.faixas.forEach((f, i) => {
  const de = i === 0 ? '' : `de ${String(cfgUsada.faixas[i - 1].ate).replace('.', ',')} m `;
  console.log(`   ${de}até ${String(f.ate).replace('.', ',')} m  ->  ${f.cm} cm`);
});
console.log(`   gola ${cfgUsada.gola} cm · viés ${cfgUsada.vies} cm\n`);

// ---------------------------------------------------------------- aplicar
const grades = JSON.parse(fs.readFileSync(ENTRADA, 'utf8').replace(/^﻿/, ''));
if (!Array.isArray(grades) || !grades.length) {
  console.error('A entrada nao e um array de grades com conteudo — abortando.');
  process.exit(1);
}

let mudam = 0, jaCertas = 0, semRegra = 0, excecoes = 0, total = 0, sobrescritas = 0;
const dist = {}, exemplos = [];
grades.forEach(g => (g.fases || []).forEach(f => {
  total++;
  const novo = regra.excedenteRegraDaFase(f);
  if (novo == null) { semRegra++; return; }
  if (regra._faseSoDe(f.nome, regra._PAL_VIES) || regra._faseSoDe(f.nome, regra._PAL_GOLA)) excecoes++;
  const cru = f.excedente;
  const proprio = !(cru === '' || cru == null);
  const atual = proprio ? Math.round(parseFloat(cru)) : null;
  if (atual === novo) { jaCertas++; return; }
  if (proprio) {
    sobrescritas++;
    if (exemplos.length < 20) exemplos.push(`${g.nome} · ${f.nome}: ${atual} -> ${novo} (comp ${f.comp})`);
  }
  const k = (proprio ? atual : 'vazio') + ' -> ' + novo;
  dist[k] = (dist[k] || 0) + 1;
  mudam++;
  if (APLICAR) f.excedente = novo;
}));

console.log(`grades: ${grades.length}   fases: ${total}`);
console.log(`MUDAM            : ${mudam}`);
console.log(`ja estavam certas: ${jaCertas}`);
console.log(`sem regra (nao toca): ${semRegra}`);
console.log(`por excecao (gola/vies): ${excecoes}`);
console.log('distribuicao:');
Object.entries(dist).sort((a, b) => b[1] - a[1]).forEach(([k, n]) => console.log(`   ${k} : ${n}`));
if (sobrescritas) {
  console.log(`\nATENCAO: ${sobrescritas} fase(s) TINHAM valor proprio e serao trocadas:`);
  exemplos.forEach(e => console.log('   ' + e));
} else {
  console.log('\nNenhuma fase com valor proprio sera trocada.');
}

if (!APLICAR) { console.log('\n(prévia — nada foi escrito. Use --aplicar para valer.)'); process.exit(0); }
if (!SAIDA) { console.error('\nFaltou --saida para gravar o resultado.'); process.exit(1); }
fs.writeFileSync(SAIDA, JSON.stringify(grades));
console.log(`\nGravado: ${SAIDA}  (${fs.statSync(SAIDA).size} bytes)`);
