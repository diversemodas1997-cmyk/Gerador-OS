/* Rode com:  node testes/riscos-cot.js

   A COLUNA RISCOS ACHA O PDF EM QUALQUER FORMA DE PASTA.

   O caminho do PDF É a informação: linha, tamanhos e largura estão nele. Só que
   o acervo tem SEIS formas de caminho — 261 PDFs medidos em 24/08/2026:

     LINHA/TAM/CM ............ 237   BM.LISA/2G-G1/182 cm/…
     LINHA/TAM/CM/Versão ...... 11   BM.TRI/2P-…/175 cm - MAPAS IMPRESSOS/Versão 1/…
     LINHA/TAM ................. 9   BM.LISA/3P/…
     LINHA/TAM/TAM/CM .......... 2   PM.LISA/Malha Piquet/Piquet Dry/175 cm/…
     LINHA/CM .................. 1   COT.JAC/157 cm/…
     LINHA/CM/TAM .............. 1   COT.PRI/152 cm/P-M-G-GG-G1/…

   A leitura antiga era por POSIÇÃO (p[1] = tamanhos, p[2] = largura) e só
   acertava as três primeiras. Nas outras o "157 cm" caía no lugar dos tamanhos,
   e o PDF não aparecia na coluna Riscos de grade nenhuma — foi o que o Junior
   viu nas três linhas COT, que têm um PDF cada, cada um numa forma diferente.

   Agora a leitura é por CONTEÚDO: o segmento que diz "N cm" é a largura, os
   outros são candidatos a tamanhos, e o NOME DO ARQUIVO entra como candidato
   quando a pasta não tem nível de tamanhos.

   Este teste roda contra o índice REAL (dados/riscos-pdf.json): é ele que muda
   quando alguém acrescenta uma pasta em forma nova. */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
const idx = JSON.parse(fs.readFileSync(path.join(RAIZ, 'dados', 'riscos-pdf.json'), 'utf8'));

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const linhaConst = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = .*$', 'm'));
  if (!m) { console.error('nao achei a const ' + nome); process.exit(1); }
  return m[0];
};
const blocoConst = (nome) => {
  const i = src.indexOf('const ' + nome + ' = ');
  const j = src.indexOf('\n};', i);
  if (i < 0 || j < 0) { console.error('nao achei a const ' + nome); process.exit(1); }
  return src.slice(i, j + 3);
};

const motor = [
  corta('function _normNome'),
  blocoConst('_riscoCmDoTexto'),
  linhaConst('_RISCO_TAM_RE'),
  corta('function _riscoTamsDoTexto'),
  corta('function _riscoItemDoCaminho'),
  corta('function _riscoNomeTamanhos'),
  corta('function _riscoFormasDoNome'),
  linhaConst('_chaveTam'),
  corta('function _gradeNomePartes'),
  corta('function _riscosDaGrade')
].join('\n');

// Roda a coluna Riscos de uma grade contra o índice inteiro.
function riscosDe(grade, arquivos) {
  const fn = new Function('grade', 'arquivos', `
    ${motor}
    const _riscosIdx = { pasta: '', gerado: '', itens: arquivos.map(_riscoItemDoCaminho) };
    return _riscosDaGrade(grade);
  `);
  return fn(grade, arquivos);
}
// O item lido de um caminho, isolado.
function itemDe(rel) {
  return new Function('rel', `${motor}\nreturn _riscoItemDoCaminho(rel);`)(rel);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

const grade = (nome, tamanhos) => ({ id: 'g1', nome, tamanhos });
const P_G1 = { p: 1, m: 1, g: 1, gg: 1, g1: 1 };
const acha = (g, trecho) => {
  const r = riscosDe(g, idx.arquivos);
  return (r.itens || []).some(i => i.rel.includes(trecho));
};

/* ---------- 1. as três linhas COT, uma por forma de pasta ---------- */

ok('COT.JAC (LINHA/CM — os tamanhos só no nome do arquivo)',
   acha(grade('P-M-G-GG-G1 | COT.JAC | 157cm', P_G1), 'COT.JAC/157 cm/'));
ok('COT.PRI (LINHA/CM/TAM — largura e tamanhos invertidos)',
   acha(grade('P-M-G-GG-G1 | COT.PRI | 152cm', P_G1), 'COT.PRI/152 cm/'));
ok('COT.RUG (LINHA/TAM/CM — a forma comum, que já funcionava)',
   acha(grade('P-M-G-GG-G1 | COT.RUG | 157cm', P_G1), 'COT.RUG/P-M-G-GG-G1/'));

/* ---------- 2. a largura passa a ser lida nessas formas ---------- */

eq('COT.JAC: a largura sai do segmento certo', itemDe('COT.JAC/157 cm/COT.JAC - CORPO -P-M-G-GG-G1.pdf').cm, '157');
eq('COT.PRI: idem, com a ordem invertida', itemDe('COT.PRI/152 cm/P-M-G-GG-G1/COT.PRI - CORPO - P-M-G-GG-G1.pdf').cm, '152');
// Largura lida é largura conferida: a grade de OUTRA largura recebe o aviso.
ok('a grade de 200cm é avisada de que o PDF é de outra largura',
   /OUTRA largura/.test(riscosDe(grade('P-M-G-GG-G1 | COT.JAC | 200cm', P_G1), idx.arquivos).aviso || ''),
   riscosDe(grade('P-M-G-GG-G1 | COT.JAC | 200cm', P_G1), idx.arquivos).aviso);

/* ---------- 3. PM.LISA: dois níveis de tecido antes da largura ---------- */

ok('PM.LISA (LINHA/TAM/TAM/CM — "Malha Piquet/Piquet Dry")',
   acha(grade('M-G-GG-G1-G3 | PM.LISA | 175cm', { m: 1, g: 1, gg: 1, g1: 1, g3: 1 }),
        'PM.LISA/Malha Piquet/Piquet Dry/'));

/* ---------- 4. o que já funcionava continua funcionando ---------- */

ok('BM.LISA/2G-G1/182 cm (forma comum)',
   acha(grade('2G-G1 | BM.LISA | 182cm', { g: 2, g1: 1 }), 'BM.LISA/2G-G1/182 cm/'));
ok('BM.LISA/3P (pasta sem largura)',
   acha(grade('3P | BM.LISA', { p: 3 }), 'BM.LISA/3P/'));
// A pasta sem largura entra em qualquer grade da linha: "não diz" não é "diz outra".
ok('a pasta sem largura entra mesmo com largura no nome da grade',
   acha(grade('3P | BM.LISA | 180cm', { p: 3 }), 'BM.LISA/3P/'));

/* ---------- 5. o nome do arquivo não inventa tamanho ---------- */

// "CORPO" tem um P e "PM.LISA" começa com PM: nenhum dos dois é tamanho.
eq('o P de CORPO não vira tamanho', itemDe('X/Y cm/COT - CORPO.pdf').tams.length, 1);
ok('a linha nunca entra como candidato de tamanho',
   !itemDe('PM.LISA/175 cm/PM.LISA - CORPO.pdf').tams.some(t => /PM/i.test(t)),
   itemDe('PM.LISA/175 cm/PM.LISA - CORPO.pdf').tams);

/* ---------- 6. o acervo inteiro: quantos caminhos ninguém consegue ler ---------- */

// Cara de lista de tamanhos: é o que a coluna Riscos consegue casar com uma grade.
const cara = t => /^(?:\d*\s*[xX]?\s*(?:GG|G[123]|[PMG]))(?:\s*-\s*\d*\s*[xX]?\s*(?:GG|G[123]|[PMG]))*$/i.test(String(t).trim());
const orfaos = idx.arquivos.filter(rel => !itemDe(rel).tams.some(cara));
// Os 4 que sobram são de "BM.TRI/2PP-2G2" — pasta e arquivos falam de um
// tamanho "PP" que não existe no cadastro. É dado sujo, não leitura: consertar
// aqui seria inventar uma grade. Se este número subir, chegou pasta em forma
// nova e a leitura precisa aprendê-la.
eq('caminhos que nenhuma grade consegue casar', orfaos.length, 4);
ok('e todos eles são o "2PP" da BM.TRI', orfaos.every(r => /2PP/i.test(r)), orfaos);

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
