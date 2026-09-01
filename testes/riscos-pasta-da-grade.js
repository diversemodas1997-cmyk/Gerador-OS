/* Rode com:  node testes/riscos-pasta-da-grade.js

   A JANELA DE RISCOS MOSTRA SÓ A PASTA DAQUELA GRADE.

   O caminho do PDF é a informação: `LINHA/TAMANHOS/LARGURA cm/arquivo.pdf`. Até
   01/09/2026 o NOME DO ARQUIVO entrava sempre como mais um candidato de
   tamanhos, com o argumento de que "candidato a mais não estraga". Estraga:
   candidato a mais faz o PDF aparecer na janela de uma grade que não é a dele.

   Foi o que o Junior viu na CM.TRI. Dentro de `P-M-2G-GG-G1-G2-G3` havia duas
   ribanas chamadas `... - P-M-G-GG-G1-G2-G3-v1/v2.pdf` — cópias que guardaram o
   nome da grade vizinha —, e por causa do NOME a pasta inteira aparecia na
   janela da grade "P ao G3 | CM.TRI", que tem a própria ribana com esses mesmos
   dois nomes. Duas grades diferentes, uma pasta trocada.

   A regra agora: a PASTA manda. O nome do arquivo continua valendo onde ele é a
   ÚNICA pista — o acervo em que a pasta vai da linha direto para a largura, ou
   nomeia o pano em vez do tamanho ("PM.LISA/malha piquet/…").

   O que este teste protege:

     · pasta que diz tamanho GANHA do nome do arquivo que diz outro;
     · o caso real da CM.TRI, com os nomes de arquivo de verdade;
     · pasta sem faixa de tamanhos continua usando o nome (PM.LISA, COT);
     · pasta e nome que CONCORDAM seguem funcionando (215 dos 279 PDFs);
     · a largura continua sendo lida de qualquer posição do caminho. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const fim = src.indexOf('\n}', i), fimC = src.indexOf('\n};', i);
  const j = (fimC >= 0 && fimC < fim) ? fimC + 3 : fim + 2;
  if (fim < 0 && fimC < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j) + '\n';
}
const A = new Function(
  recorte('function _normNome', 'a normalizacao de nome')
  + recorte('const _riscoCmDoTexto', 'a leitura da largura')
  + recorte('const _RISCO_TAM_RE', 'a expressao dos tamanhos')
  + recorte('function _riscoTamsDoTexto', 'os tamanhos de um texto')
  + recorte('function _riscoItemDoCaminho', 'a leitura do caminho')
  + recorte('const _chaveTam', 'a chave de tamanho')
  + 'return { _riscoItemDoCaminho, _chaveTam };')();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const chaves = rel => A._riscoItemDoCaminho(rel).tams.map(A._chaveTam);

console.log('-- o caso real da CM.TRI --');
// As duas ribanas dentro da pasta 2G, com o nome da grade vizinha.
const ribanaNaPastaErrada = 'CM.TRI/P-M-2G-GG-G1-G2-G3/116.5 cm - MAPAS IMPRESSOS/CM.TRI - RIBANA - P-M-G-GG-G1-G2-G3-v1.pdf';
ok('1. o PDF pertence a pasta em que esta, e nao ao que o nome dele diz',
  chaves(ribanaNaPastaErrada).join(',') === 'p-m-2g-gg-g1-g2-g3', chaves(ribanaNaPastaErrada));
ok('2. e nao carrega mais a faixa da grade vizinha',
  chaves(ribanaNaPastaErrada).indexOf('p-m-g-gg-g1-g2-g3') < 0, chaves(ribanaNaPastaErrada));
const ribanaCerta = 'CM.TRI/P-M-G-GG-G1-G2-G3/116.5 cm/CM.TRI - RIBANA - P-M-G-GG-G1-G2-G3-v1.pdf';
ok('3. a ribana da grade "P ao G3" continua sendo dela',
  chaves(ribanaCerta).join(',') === 'p-m-g-gg-g1-g2-g3', chaves(ribanaCerta));

console.log('');
console.log('-- o outro caso: o nome na forma simples dentro da pasta do dobro --');
const dobro = 'CM.LISA/2M-4G-2GG/117 cm - MAPAS IMPRESSOS/CM.LISA - RIBANA M-2G-GG.pdf';
ok('4. fica com a faixa da PASTA (o dobro), nao com a do nome',
  chaves(dobro).join(',') === '2m-4g-2gg', chaves(dobro));

console.log('');
console.log('-- onde o nome do arquivo AINDA e a unica pista --');
const piquet = 'PM.LISA/malha piquet/piquet dry/PM.LISA - CORPO M-G-GG-G1-G3.pdf';
ok('5. pasta que nomeia o PANO, e nao o tamanho: o nome do arquivo vale',
  chaves(piquet).indexOf('m-g-gg-g1-g3') >= 0, chaves(piquet));
const soLargura = 'COT.JAC/157 cm/COT.JAC - CORPO -P-M-G-GG-G1.pdf';
ok('6. pasta que vai da linha direto para a largura: o nome do arquivo vale',
  chaves(soLargura).indexOf('p-m-g-gg-g1') >= 0, chaves(soLargura));

console.log('');
console.log('-- o que nao pode ter mudado --');
const normal = 'CM.TRI/2P-2G1/116.5 cm/CM.TRI - CORPO 1 - 2P-2G1.pdf';
ok('7. pasta e nome que concordam seguem iguais',
  chaves(normal).join(',') === '2p-2g1', chaves(normal));
const it = A._riscoItemDoCaminho('CM.LISA/M-GG-G3/117 cm/CM.LISA CORPO M-GG-G3.pdf');
ok('8. a largura continua sendo lida', it.cm === '117', it.cm);
ok('9. a linha continua sendo o primeiro pedaco', it.linha === 'CM.LISA', it.linha);
const semLargura = A._riscoItemDoCaminho('CM.LISA/2G-G2/CM.LISA - CORPO - 2G-G2.pdf');
ok('10. pasta sem nivel de largura nao quebra e mantem o tamanho',
  semLargura.cm === '' && A._chaveTam(semLargura.tam) === '2g-g2', semLargura);

console.log('');
console.log('-- a regra, dita no codigo --');
ok('11. o nome so entra quando a pasta nao diz tamanho',
  /const pastaDizTamanho = tams\.some\(seg => _riscoTamsDoTexto\(seg\)\);\s*\n\s*if \(doNome && !pastaDizTamanho\) tams\.push\(doNome\);/.test(src));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
