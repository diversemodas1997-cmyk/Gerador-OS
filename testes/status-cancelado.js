/* Rode com:  node testes/status-cancelado.js

   O STATUS "CANCELADO" (17/09/2026, Junior: "crie um novo status com nome
   Cancelado" e "as OS com status cancelado devem receber uma faixa de texto em
   vermelho cruzando a folha de ponta a ponta escrito CANCELADO").

   Cancelado e o fim que NAO produziu nada. Por isso ele nao e so mais uma linha
   na tabela: ele muda tres contas do programa, e e isso que este teste prende.

     1. O CARIMBO. Nao nasce de etapa nenhuma do checklist — so da mao. Marcar
        Corte numa OS cancelada devolve a OS para o fluxo, como acontece com
        Parado: o carimbo vale ate a proxima etapa ser marcada.

     2. O PANO. Cancelar nao consome: a reserva volta para a prateleira. Mas o
        que ja tinha sido baixado fica baixado, porque cortado nao volta a ser
        rolo. Aqui se prova que `cancelado` NAO esta entre os status que baixam.

     3. A FOLHA. A faixa e desenhada na propria folha (nao numa camada de
        impressao), senao ela nao sairia no PDF — que sai por html2canvas e
        ignora o @media print. E ela e recortada no papel: a diagonal de uma A4
        e mais larga que a folha.

   Recorta as funcoes e o CSS de verdade. */
const fs = require('fs');
const path = require('path');

const raiz = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(raiz, 'app.js'), 'utf8');
const css = fs.readFileSync(path.join(raiz, 'styles.css'), 'utf8');

function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}
function constante(nome) {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
}

const api = new Function(`
  ${constante('STATUS_OS')}
  ${constante('_STATUS_QUE_BAIXAM')}
  ${constante('STATUS_FIM')}
  return { STATUS_OS, _STATUS_QUE_BAIXAM };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const cancelado = api.STATUS_OS.find(s => s.k === 'cancelado');

console.log('-- o status existe, e e um so --');
ok('1. "Cancelado" esta na tabela', !!cancelado && cancelado.rotulo === 'Cancelado',
   JSON.stringify(cancelado));
ok('2. fica junto de Parado, os dois de fora da fila, antes de Estoque',
   api.STATUS_OS.map(s => s.k).join(',').includes('parado,cancelado,estoque'),
   api.STATUS_OS.map(s => s.k).join(','));

console.log('');
console.log('-- nao nasce do checklist: so da mao --');
ok('3. nao tem regra de etapa (`re`)', !cancelado.re, String(cancelado.re));
ok('4. nem e aceso pelo enfesto', !cancelado.enfesto, String(cancelado.enfesto));
/* `ordem` desempata OS antiga sem etapasSeq, e e a ordem REAL da producao.
   Cancelado nao e um degrau do caminho: nao pode disputar esse desempate com
   Cortando ou Ensacado, senao uma OS antiga cancelada seria lida pela ordem. */
ok('5. e nao tem `ordem`: nao e um degrau do caminho', cancelado.ordem === undefined,
   String(cancelado.ordem));

console.log('');
console.log('-- cancelar nao consome pano --');
ok('6. nao esta entre os status que baixam o material',
   api._STATUS_QUE_BAIXAM.indexOf('cancelado') < 0, api._STATUS_QUE_BAIXAM.join(', '));
/* A devolucao da reserva mora em _estoqueSeguirStatusOS, e ela tem de ser
   condicional: so volta o que AINDA e reserva. */
const seguir = recorte('async function _estoqueSeguirStatusOS', 'o estoque que segue o status');
ok('7. a reserva volta para a prateleira ao cancelar',
   /alvo === 'cancelado'/.test(seguir) && /estornarBaixaEstoqueOS/.test(seguir), seguir.slice(0, 200));
ok('8. mas so quando nada foi consumido ainda',
   /consumido/.test(seguir) && seguir.indexOf("m.status === 'consumido'") < seguir.indexOf('estornarBaixaEstoqueOS'),
   seguir.slice(seguir.indexOf('cancelado'), seguir.indexOf('cancelado') + 300));

console.log('');
console.log('-- a cor e unica, como as outras dez --');
const cores = api.STATUS_OS.map(s => s.cor);
ok('9. nenhuma cor se repete na tabela', new Set(cores).size === cores.length,
   api.STATUS_OS.map(s => s.k + '=' + s.cor).join(' '));
ok('10. e a do cancelado nao e a do nao-iniciado',
   cancelado.cor !== api.STATUS_OS.find(s => s.k === 'nao-iniciado').cor, cancelado.cor);

console.log('');
console.log('-- a faixa na folha --');
ok('11. a folha escreve a faixa quando a OS esta cancelada',
   /_statusOS\(o\) === 'cancelado'/.test(src) && /_carimboCanceladoHtml\(\)/.test(src),
   'nao achei o carimbo em renderPrintSheet');
ok('12. e o texto e CANCELADO', /<span>CANCELADO<\/span>/.test(src), 'texto diferente');
/* O PDF sai por html2canvas, que fotografa a TELA. Um carimbo que morasse
   dentro de @media print nao apareceria no PDF que fica na pasta. */
const iRegra = css.indexOf('.sheet-carimbo-cancelado {');
const iPrint = css.indexOf('@media print', css.indexOf('.sheet-carimbo-cancelado span'));
ok('13. a faixa e estilo de TELA, nao de @media print (senao sumiria do PDF)',
   iRegra > 0 && css.lastIndexOf('@media print', iRegra) < css.lastIndexOf('}', iRegra) - 0
   && /position:\s*absolute/.test(css.slice(iRegra, iRegra + 400)),
   'regra em ' + iRegra);
ok('14. cruza na diagonal REAL da A4 (atan 297/210 = 54,7 graus)',
   /rotate\(-54\.7deg\)/.test(css.slice(iRegra, iRegra + 400)),
   css.slice(iRegra, iRegra + 400));
ok('15. passa das duas pontas do papel',
   /width:\s*1[5-9]\d%/.test(css.slice(iRegra, iRegra + 400)),
   css.slice(iRegra, iRegra + 400));
/* QUEM RECORTA E A AREA, E NAO A FOLHA (17/09/2026, Junior: "a folha de OS
   diminui o tamanho quando recebe a faixa cancelado").

   `overflow: hidden` na .sheet escondia o transbordo da faixa, mas nao o tirava
   do scrollWidth/scrollHeight — e e deles que ajustarImpressaoParaA4 tira a
   escala com que a folha entra no papel. Medido no Chrome, a mesma folha:
     sem carimbo              210,1 x 297,1mm  ->  escala 0,9995
     com o carimbo na .sheet  234,2 x 321,2mm  ->  escala 0,8967
     com a area que recorta   210,1 x 297,1mm  ->  escala 0,9995
   A folha nao diminuia: ela era encaixada na A4 como se fosse 24mm maior.

   Entao a .sheet NAO pode levar overflow, e a faixa tem de morar dentro de uma
   area do tamanho exato dela. Estes dois testes sao um par: um exige a area, o
   outro proibe o atalho que parecia funcionar. */
const iArea = css.indexOf('.sheet-carimbo-area {');
ok('16. quem recorta e uma area do tamanho da folha',
   iArea > 0 && /overflow:\s*hidden/.test(css.slice(iArea, iArea + 300))
   && /top:\s*0;\s*right:\s*0;\s*bottom:\s*0;\s*left:\s*0/.test(css.slice(iArea, iArea + 300)),
   css.slice(iArea, iArea + 300));
ok('16b. e a .sheet NAO leva overflow: e isso que a fazia encolher na impressao',
   /\.sheet\.folha-cancelada\s*\{[^}]*position:\s*relative[^}]*\}/.test(css)
   && !/\.sheet\.folha-cancelada\s*\{[^}]*overflow/.test(css),
   (css.match(/\.sheet\.folha-cancelada\s*\{[^}]*\}/) || [''])[0]);
ok('16c. e a faixa vai DENTRO da area, nao solta na folha',
   /sheet-carimbo-area[^<]*>\s*'\s*\+\s*'<div class="sheet-carimbo-cancelado"/.test(src)
   || /<div class="sheet-carimbo-area"[\s\S]{0,120}sheet-carimbo-cancelado/.test(src),
   'a faixa nao esta dentro da area');
ok('17. nao rouba o clique do checklist que esta embaixo',
   /pointer-events:\s*none/.test(css.slice(iArea, iArea + 300)),
   css.slice(iArea, iArea + 300));
ok('18. e sai colorida tambem no papel', iPrint > iRegra
   && /print-color-adjust:\s*exact/.test(css.slice(iPrint, iPrint + 300)),
   'sem print-color-adjust');

console.log('');
console.log('-- a faixa aparece no instante do carimbo, sem redesenhar a folha --');
const rst = recorte('function renderStatusFolhaOS', 'o status no cabecalho da folha');
ok('19. renderStatusFolhaOS chama o carimbo nos DOIS caminhos (com e sem foco)',
   (rst.match(/_carimboCanceladoNaFolha\(o\)/g) || []).length === 2, rst);

console.log('');
console.log('-- o painel do Inicio nao conta lote desistido como producao --');
ok('20. o cancelado esta marcado como fora da producao', cancelado.foraDaProducao === true,
   String(cancelado.foraDaProducao));
const painel = src.slice(src.indexOf('function _dashPorStatusHtml'), src.indexOf('const STATUS_TERMINAL_DASH'));
ok('21. e o total "em producao" o deixa de fora',
   /!x\.st\.foraDaProducao/.test(painel), 'o filtro do emProducao nao olha foraDaProducao');

console.log('');
console.log(falhas ? falhas + ' FALHA(S)' : 'todos os testes passaram');
process.exit(falhas ? 1 : 0);
