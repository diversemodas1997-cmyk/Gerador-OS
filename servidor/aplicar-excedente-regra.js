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
          --saida <grades-novo.json> [--aplicar]

   Sem --aplicar ele só relata o que MUDARIA (e não escreve a saída). É assim
   que se confere antes.
*/
const fs = require('fs');
const path = require('path');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const ENTRADA = arg('entrada');
const SAIDA = arg('saida');
const APLICAR = process.argv.includes('--aplicar');
if (!ENTRADA) { console.error('Faltou --entrada <arquivo com o array de grades>.'); process.exit(1); }

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
const regra = new Function([
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
  recorte('function excedentePorComprimento'),
  recorte('function excedenteRegraDaFase'),
  'return { excedenteRegraDaFase, _faseSoDe, _PAL_VIES, _PAL_GOLA };'
].join('\n'))();

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
