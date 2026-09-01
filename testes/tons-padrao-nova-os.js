/* Rode com:  node testes/tons-padrao-nova-os.js

   A OS NOVA JÁ NASCE COM TRÊS LINHAS DE TOM.

   A folha nascia com UMA linha, e quem enfesta em três tonalidades — que é o
   normal da casa — tinha de abrir as outras duas no + antes de escrever
   qualquer coisa. Pior no papel: a folha impressa saía com uma linha só, e as
   tonalidades 2 e 3 iam para a margem, à caneta, fora do lugar onde o programa
   saberia lê-las depois.

   Três é o padrão de quem enfesta, não um teto: o + continua abrindo até
   MAX_TONS e o − recolhe o que não for usado.

   O que este teste protege:

     · OS nova nasce com 3 tons marcados, contíguos a partir do 1;
     · OS JÁ GRAVADA não é tocada ao ser reeditada — mexer nela mudaria o volume
       da expedição de lote já planejado (cada tonalidade é ensacada separada:
       pacotes = tamanhos × tonalidades + 1);
     · nem o VALOR já lançado nos tons se perde na reedição;
     · uma OS nova que já chegue com tom definido não é sobrescrita;
     · o padrão nasce SEM valor nenhum: a linha vem vazia esperando as camadas,
       porque quem determina os números continua sendo a fase principal;
     · o padrão respeita o teto MAX_TONS. */
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
const monta = (STATE) => new Function('STATE', `
  ${recorte('const MAX_TONS', 'o teto de tons')}
  ${recorte('function _aplicarTonsPadrao', 'o padrao de tons')}
  ${recorte('function tonsEfetivos', 'os tons efetivos')}
  ${recorte('function nLinhasTomOS', 'as linhas de tom')}
  ${recorte('function _mesclarComOSExistente', 'a mescla com a OS existente')}
  return { _mesclarComOSExistente, nLinhasTomOS, _aplicarTonsPadrao,
           MAX_TONS, TONS_PADRAO_NOVA_OS };
`)(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// Uma OS de março, já gravada, com o tom lançado.
const osVelha = () => ({
  id: 'velha', os: '0400', criadoEm: '2026-03-01T10:00:00.000Z',
  progresso: { totalTamanhoTons: { 1: true }, totalTamanhoTomValor: { 1: 50 } }
});

console.log('-- a OS nova --');
let A = monta({ ordens: [osVelha()] });
const nova = A._mesclarComOSExistente({ id: 'nova', os: '0509' });
ok('1. nasce com 3 linhas de tom', A.nLinhasTomOS(nova) === 3, A.nLinhasTomOS(nova));
ok('2. e os tons sao contiguos a partir do 1',
  JSON.stringify(nova.progresso.totalTamanhoTons) === JSON.stringify({ 1: true, 2: true, 3: true }),
  nova.progresso.totalTamanhoTons);
ok('3. sem VALOR nenhum: a linha vem vazia, esperando as camadas',
  !nova.progresso.totalTamanhoTomValor, nova.progresso.totalTamanhoTomValor);

console.log('');
console.log('-- a OS que ja existe nao e tocada --');
A = monta({ ordens: [osVelha()] });
const edit = A._mesclarComOSExistente({ id: 'velha', os: '0400' });
ok('4. reeditar nao abre tom nenhum — o volume da expedicao dela nao muda',
  A.nLinhasTomOS(edit) === 1, A.nLinhasTomOS(edit));
ok('5. e o valor ja lancado sobrevive',
  edit.progresso.totalTamanhoTomValor && edit.progresso.totalTamanhoTomValor[1] === 50,
  edit.progresso.totalTamanhoTomValor);

console.log('');
console.log('-- as bordas --');
A = monta({ ordens: [] });
const jaComTom = A._mesclarComOSExistente({
  id: 'x', os: '0510', progresso: { totalTamanhoTons: { 1: true, 2: true } }
});
ok('6. OS nova que ja chega com tom definido nao e sobrescrita',
  A.nLinhasTomOS(jaComTom) === 2, A.nLinhasTomOS(jaComTom));
const semProgresso = A._aplicarTonsPadrao({ id: 'y' });
ok('7. OS sem `progresso` nenhum nao quebra',
  A.nLinhasTomOS(semProgresso) === 3, A.nLinhasTomOS(semProgresso));
ok('8. o padrao respeita o teto de tons',
  A.TONS_PADRAO_NOVA_OS <= A.MAX_TONS,
  { padrao: A.TONS_PADRAO_NOVA_OS, teto: A.MAX_TONS });
ok('9. nada quebra com null', A._aplicarTonsPadrao(null) === null);

console.log('');
console.log('-- o + e o − continuam mandando --');
ok('10. o + abre ate MAX_TONS', /if \(n >= MAX_TONS\)/.test(src));
ok('11. o − recolhe, e o Tom 1 nunca sai',
  /if \(n <= 1\) \{ toast\('O Tom 1 nao pode ser retirado\.'|if \(n <= 1\) \{ toast\('O Tom 1 não pode ser retirado\.'/.test(src));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
