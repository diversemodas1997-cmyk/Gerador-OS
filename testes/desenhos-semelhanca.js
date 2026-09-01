/* Rode com:  node testes/desenhos-semelhanca.js

   OS DESENHOS SEMELHANTES FICAM JUNTOS NA HORA DE ESCOLHER.

   A lista do formulário de OS era a ordem de cadastro: as nove camisetas
   básicas, as oito blusas de moletom e as seis texturizadas apareciam
   embaralhadas, e achar "a básica verde" era ler código por código. Quem
   escolhe o desenho não procura um código — procura uma FAMÍLIA e, dentro dela,
   uma cor.

   A chave é o SKU do produto (o `skuLinha` sem o sufixo de cor), medido com os
   33 desenhos da fábrica: 8 grupos, nenhum com um só. Ganha do modelo por
   separar as três Camisetas Oversized Texturizadas, que são panos diferentes.

   O que este teste protege:

     · a chave é o SKU SEM a cor — senão cada cor viraria um grupo de um só,
       que é o erro que reprovou a pasta por faixa de tamanhos nas grades;
     · a família COT não volta a ser um grupo só;
     · dentro do grupo, ordem por código com leitura numérica (001 antes de
       0010, e o 0024 antes do 0025 — no cadastro eles estão ao contrário);
     · o cabeçalho só afirma o modelo quando ele vale para o grupo INTEIRO;
     · desenho sem SKU não some nem se mistura: vai para o fim, em grupo próprio;
     · a lista suspensa sai em <optgroup>, e o desenho já escolhido continua
       marcado depois de agrupada. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}
const A = new Function(`
  ${recorte('function _rotuloDesenhoOS', 'o rotulo do desenho')}
  ${recorte('function _desenhoSkuProduto', 'o SKU do produto')}
  ${recorte('function _desenhoModelo', 'o modelo do desenho')}
  ${recorte('function _desenhoVariacao', 'a variacao do desenho')}
  ${recorte('function _desenhosAgrupados', 'o agrupamento')}
  ${recorte('function _rotuloGrupoDesenho', 'o rotulo do grupo')}
  ${recorte('function _rotuloDesenhoNoGrupo', 'o rotulo dentro do grupo')}
  return { _desenhoSkuProduto, _desenhosAgrupados, _rotuloGrupoDesenho, _rotuloDesenhoNoGrupo };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// O cadastro da fábrica, reduzido ao que decide o agrupamento.
const d = (codigo, desc, sku) => ({ id: 'id' + codigo, codigo, desc, skuLinha: sku });
const CADASTRO = [
  d('001', 'Camiseta Básica | Preto', 'CM.LISA-PRE'),
  d('002', 'Camiseta Básica | Branco', 'CM.LISA-BRA'),
  d('009', 'Camiseta Básica | Grafite', 'CM.LISA-GRAF'),
  d('0010', 'Camiseta Recortada | Preto/Branco', 'CM.REC.LISA-PRE'),
  d('0013', 'Camiseta Recortada | Verde/Branco', 'CM.REC.LISA-VERDE'),
  d('0018', 'Blusa Moletom Básica | Preto', 'BM.LISA-PRE'),
  d('0027', 'Blusa Moletom Básica | Verde', 'BM.LISA-VERDE'),
  d('0025', 'Blusa Moletom Tricolor | Preto/Mostarda/Off-White', 'BM.TRI-PRE'),
  d('0024', 'Blusa Moletom Tricolor | Verde/Preto/Bege', 'BM.TRI-VERDE'),
  d('0028', 'Camiseta Oversized Texturizada | Rugão Preto', 'COT.RUG-PRE'),
  d('0029', 'Camiseta Oversized Texturizada | Rugão Off-white', 'COT.RUG-OFF'),
  d('0030', 'Camiseta Oversized Texturizada | Prime Preto', 'COT.PRI-PRE'),
  d('0031', 'Camiseta Oversized Texturizada | Prime Off-white', 'COT.PRI-OFF'),
  d('0032', 'Camiseta Oversized Texturizada | Jacguar Preto', 'COT.JAC-PRE'),
  d('0033', 'Camiseta Oversized Texturizada | Jacguar Off-white', 'COT.JAC-OFF')
];

console.log('-- a chave --');
ok('1. o SKU do produto e o skuLinha SEM a cor',
  A._desenhoSkuProduto({ skuLinha: 'CM.LISA-PRE' }) === 'CM.LISA',
  A._desenhoSkuProduto({ skuLinha: 'CM.LISA-PRE' }));
ok('2. e aguenta o SKU de tres partes',
  A._desenhoSkuProduto({ skuLinha: 'CM.REC.LISA-PRE' }) === 'CM.REC.LISA',
  A._desenhoSkuProduto({ skuLinha: 'CM.REC.LISA-PRE' }));
ok('3. sem SKU, nao inventa um', A._desenhoSkuProduto({}) === '', A._desenhoSkuProduto({}));

console.log('');
console.log('-- os grupos --');
const g = A._desenhosAgrupados(CADASTRO);
const cmlisaDireto = A._desenhosAgrupados(CADASTRO.filter(x => x.skuLinha.indexOf('CM.LISA-') === 0));
ok('4. as familias saem separadas e em ordem',
  g.map(x => x.sku).join(', ') === 'BM.LISA, BM.TRI, CM.LISA, CM.REC.LISA, COT.JAC, COT.PRI, COT.RUG',
  g.map(x => x.sku));
ok('5. a cor NAO entra na chave: as tres CM.LISA caem no MESMO grupo',
  cmlisaDireto.length === 1 && cmlisaDireto[0].itens.length === 3,
  cmlisaDireto.map(x => x.sku + ':' + x.itens.length));
ok('5b. e nenhum grupo fica com um so — foi isso que reprovou a pasta por tamanhos',
  g.every(x => x.itens.length >= 2), g.map(x => x.sku + ':' + x.itens.length));
ok('6. a familia COT nao vira um grupo so — sao panos diferentes',
  g.filter(x => x.sku.indexOf('COT.') === 0).length === 3,
  g.filter(x => x.sku.indexOf('COT.') === 0).map(x => x.sku));

console.log('');
console.log('-- a ordem dentro do grupo --');
const bmtri = g.find(x => x.sku === 'BM.TRI');
ok('7. por codigo, e o 0024 vem antes do 0025 (no cadastro estao ao contrario)',
  bmtri.itens.map(x => x.codigo).join(',') === '0024,0025', bmtri.itens.map(x => x.codigo));
const cmlisa = g.find(x => x.sku === 'CM.LISA');
ok('8. leitura numerica: 001 antes de 009, e nao ordem de texto',
  cmlisa.itens.map(x => x.codigo).join(',') === '001,002,009', cmlisa.itens.map(x => x.codigo));

console.log('');
console.log('-- o que cada linha diz --');
ok('9. o cabecalho junta SKU e modelo',
  A._rotuloGrupoDesenho(cmlisa) === 'CM.LISA · Camiseta Básica', A._rotuloGrupoDesenho(cmlisa));
ok('10. e a linha mostra so o codigo e a cor, sem repetir o modelo',
  A._rotuloDesenhoNoGrupo(cmlisa.itens[0], cmlisa) === '001 · Preto',
  A._rotuloDesenhoNoGrupo(cmlisa.itens[0], cmlisa));

const misturado = A._desenhosAgrupados([
  d('100', 'Camiseta Básica | Preto', 'XX.MIX-PRE'),
  d('101', 'Blusa Moletom Básica | Preto', 'XX.MIX-BRA')
])[0];
ok('11. modelo diferente dentro do mesmo SKU: o cabecalho nao afirma modelo nenhum',
  A._rotuloGrupoDesenho(misturado) === 'XX.MIX', A._rotuloGrupoDesenho(misturado));
ok('12. e ai a linha volta a carregar a descricao inteira',
  A._rotuloDesenhoNoGrupo(misturado.itens[0], misturado) === '100 · Camiseta Básica | Preto',
  A._rotuloDesenhoNoGrupo(misturado.itens[0], misturado));

console.log('');
console.log('-- as bordas --');
const comOrfao = A._desenhosAgrupados(CADASTRO.concat([d('999', 'Peça nova | Preto', '')]));
const ultimo = comOrfao[comOrfao.length - 1];
ok('13. desenho sem SKU vai para o FIM, e nao some',
  ultimo.sku === '' && ultimo.itens.length === 1, comOrfao.map(x => x.sku));
ok('14. e o grupo dele se anuncia, em vez de aparecer sem nome',
  A._rotuloGrupoDesenho(ultimo) === 'Sem SKU', A._rotuloGrupoDesenho(ultimo));
ok('15. lista vazia nao quebra', A._desenhosAgrupados([]).length === 0);

console.log('');
console.log('-- a lista suspensa --');
const render = src.slice(src.indexOf('function filtrarDesenhosOS'), src.indexOf('function filtrarDesenhosOS') + 2000);
ok('16. sai em <optgroup>, que e o que poe o nome do grupo na lista',
  render.indexOf('<optgroup label=') > 0 && render.indexOf('_rotuloGrupoDesenho(g)') > 0);
ok('17. e o desenho ja escolhido continua marcado depois de agrupada',
  render.indexOf("d.id === escolhido ? ' selected' : ''") > 0);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
