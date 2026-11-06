/* Rode com:  node testes/etiqueta-moletom-nao-blusa.js

   ETIQUETA DA CJM NO PADRÃO DA BLUSA (06/11/2026, OS 0676). O moletom que não
   é blusa sai por grupo de até três peças, como a BM.LISA/BM.TRI:
     · as peças de uma mesma roupa ficam juntas (calça × blusa do conjunto);
     · o que todas as peças do grupo repetem sai do nome e vira prefixo;
     · a quantidade por peça vem do componente da OS.

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(de, ate) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + de + ' no app.js'); process.exit(1); }
  return src.slice(i, j + ate.length);
}
const f = new Function(`
  ${corta('function _normNome', '\n}')}
  ${corta('function _pecasEtiquetaPelosComponentes', '\n}')}
  return _pecasEtiquetaPelosComponentes;
`)();

let falhas = 0;
const ok = (cond, msg) => { if (cond) console.log('ok   ' + msg); else { falhas++; console.log('FALHA ' + msg); } };
const c = (nome, q) => ({ nome, qtdPorPeca: q });

// OS 0676
const g676 = f({ componentes: [c('Frente D', 1), c('Frente E', 1), c('Costa D', 1), c('Costa E', 1), c('Bolsos Frente', 4), c('Bolso costa', 1),
  c('Frente Blusa Moletom Básica', 1), c('Mangas Blusa Moletom Básica', 2), c('Costas Blusa Moletom Básica', 1)] });
ok(g676.length === 3, 'nove peças em três etiquetas');
ok(g676.map(g => g.nome).join(' | ') === 'Frente D/Frente E/Costa D | Costa E/Bolsos Frente/Bolso costa | Blusa: Frente/Mangas/Costas', 'nomes dos grupos: ' + g676.map(g => g.nome).join(' | '));
ok(JSON.stringify(g676[2].porBlusa) === JSON.stringify([['Frente', 1], ['Mangas', 2], ['Costas', 1]]), 'quantidade por peça vem do componente, nome curto');
ok(g676[1].porBlusa[1][1] === 4, 'bolsos da frente: 4 por peça');
ok(g676[2].partes[1].test('mangas blusa moletom basica'), 'a parte acha o componente pelo nome inteiro (para a cor)');

// OS 0589: a calça e a blusa não se misturam
const g589 = f({ componentes: [c('Frente', 1), c('Costas', 1), c('Frente Blusa Moletom Básica', 1), c('Mangas Blusa Moletom Básica', 2), c('Costas Blusa Moletom Básica', 1)] });
ok(g589.map(g => g.nome).join(' | ') === 'Frente/Costas | Blusa: Frente/Mangas/Costas', 'calça e blusa em grupos separados: ' + g589.map(g => g.nome).join(' | '));

// Sem componentes, nada (a etiqueta cai no formato antigo, sem inventar lista)
ok(f({ componentes: [] }) === null, 'sem componentes: null');
ok(f({ componentes: [c('Frente', 0)] }) === null, 'componente sem quantidade não conta');

console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
