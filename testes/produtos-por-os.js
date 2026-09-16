/* Rode com:  node testes/produtos-por-os.js

   PRODUTOS, E NÃO PEÇAS (16/09/2026, Junior): "se um produto é composto por 4
   peças, o programa deve informar que foram produzidos x produtos e não 4x
   peças".

   O que este teste protege:
     · a OS conta PRODUTO: 100 camisetas de frente, costas e 2 mangas são 100,
       e não 400;
     · o Total geral da folha manda quando existe (é o número impresso);
     · sem folha, a conta sai dos componentes, dividindo pelo "por peça";
     · por tecido + cor, cada linha diz quantos produtos têm peça ali — a
       ribana da gola também é 100, e não entra somada com a malha.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const corta = (nome) => {
  const i = src.indexOf(nome);
  if (i < 0) { console.error('nao achei ' + nome); process.exit(1); }
  return src.slice(i, src.indexOf('\n}', i) + 2);
};

const monta = (totalGeral) => new Function(`
  const totaisPorTamanhoTomOS = () => ({ totalGeral: ${Number(totalGeral) || 0} });
  const corCanonicaPorTecido = (cor) => cor || '';
  const _normNome = s => String(s || '').trim().toLowerCase();
  ${corta('function _produtosDosComponentes')}
  ${corta('function produtosOS')}
  ${corta('function produtosPorTecidoCorOS')}
  return { produtosOS, produtosPorTecidoCorOS };
`)();

let falhas = 0;
const ok = (nome, cond, obtido) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + obtido));
  if (!cond) falhas++;
};

const tam = (n) => ({ p: n, m: n, g: n, gg: 0 });
const camiseta = {
  componentes: [
    { nome: 'Frente', materialNome: 'Malha', corNome: 'Preto', qtdPorPeca: 1, qtdPorTamanho: tam(100), qtdTotal: 300 },
    { nome: 'Costas', materialNome: 'Malha', corNome: 'Preto', qtdPorPeca: 1, qtdPorTamanho: tam(100), qtdTotal: 300 },
    { nome: 'Mangas', materialNome: 'Malha', corNome: 'Preto', qtdPorPeca: 2, qtdPorTamanho: tam(200), qtdTotal: 600 },
    { nome: 'Gola', materialNome: 'Ribana', corNome: 'Preto', qtdPorPeca: 1, qtdPorTamanho: tam(100), qtdTotal: 300 }
  ]
};

console.log('-- sem folha: a conta sai dos componentes --');
let f = monta(0);
ok('1. 1.500 peças (5 por camiseta: frente, costas, 2 mangas, gola) são 300 camisetas', f.produtosOS(camiseta) === 300, f.produtosOS(camiseta));
const linhas = f.produtosPorTecidoCorOS(camiseta);
const malha = linhas.find(l => l.tecidoNome === 'Malha');
const ribana = linhas.find(l => l.tecidoNome === 'Ribana');
ok('2. a linha da malha tem 300 produtos, não 1.200 peças', malha && malha.qtd === 300, JSON.stringify(malha));
ok('3. a linha da ribana também tem 300', ribana && ribana.qtd === 300, JSON.stringify(ribana));
ok('4. componente sem qtdPorTamanho divide pelo por peça',
   f.produtosOS({ componentes: [{ qtdPorPeca: 2, qtdTotal: 80 }] }) === 40,
   f.produtosOS({ componentes: [{ qtdPorPeca: 2, qtdTotal: 80 }] }));
ok('5. OS sem componentes e sem folha: zero', f.produtosOS({}) === 0, f.produtosOS({}));

console.log('');
console.log('-- com folha: o Total geral manda --');
f = monta(250);
ok('6. o número é o da folha, mesmo com componentes antigos', f.produtosOS(camiseta) === 250, f.produtosOS(camiseta));
const m2 = f.produtosPorTecidoCorOS(camiseta).find(l => l.tecidoNome === 'Malha');
ok('7. e as linhas por tecido acompanham a folha', m2 && m2.qtd === 250, JSON.stringify(m2));

console.log('');
const html = src;
ok('8. o Início fala em produtos', /em produtos \(unidades completas\) e em número de OS/.test(html), '');
ok('9. o Ranking usa produtosOS, e não enfesto.totalPecas',
   /const totalPecas = produtosOS\(o\);/.test(src), '');

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
