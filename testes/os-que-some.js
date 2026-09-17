/* Rode com:  node testes/os-que-some.js

   A OS QUE SOME DO PROGRAMA E FICA SO EM PDF NA PASTA.

   17/09/2026, Junior: "olhei na pasta de OS e descobri que os numeros estao
   sendo ordenados corretamente, mas algumas OS que foram geradas e salvas na
   pasta nao estao no programa" — a 0509 (01/09), a 0536 e a 0537 (10/09).

   O QUE ACONTECIA. Todos os dados vivem num blob so, e cada gravacao mescla o
   que este aparelho mudou com o que esta no servidor (_mergeListaPorRegistro).
   O merge apagava do servidor todo registro que estivesse na BASE e nao na
   lista local, lendo a ausencia como "excluido aqui". So que a base ficava mais
   nova que o STATE: o polling adotava o servidor para dentro do cache e da base
   e, quando o ultimo a gravar tinha sido este proprio aparelho, voltava sem
   refazer o STATE. Dai em diante, a OS criada no OUTRO PC existia na base e nao
   no STATE — e a gravacao seguinte a apagava para todo mundo.

   O QUE ESTE TESTE PROTEGE:

     · o merge NAO apaga registro que este aparelho nunca viu (a 0509);
     · o merge NAO apaga registro que sumiu do STATE sem ninguem ter excluido
       (a regressao exata, com a base "na frente" do local);
     · o merge APAGA o que foi apagado aqui de proposito;
     · saveState anota quem sumiu (_apagadosLocaisRegistrar) e desanota quem
       voltou;
     · e a contagem sabe dizer quais numeros faltam, separando o buraco no meio
       da producao da faixa do arquivo de papel.

   Recorta as funcoes do app.js de verdade. */
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
function constante(nome) {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
}

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  let _apagadosAqui = {};
  ${constante('NUMERO_OS_MAX')}
  ${constante('OS_SALTO_ESTRANHO')}
  ${constante('OS_FAIXA_ARQUIVO')}
  ${recorte('function formatarNumeroOS', 'o numero em quatro digitos')}
  const sanitizeForFilename = (s) => String(s);
  ${recorte('function _numeroOSCanonico', 'o numero canonico')}
  ${recorte('function _numerosOSExistentes', 'os numeros existentes')}
  ${recorte('function _topoDaFilaOS', 'o topo da fila')}
  ${recorte('function _osNumerosFaltando', 'os numeros que faltam')}
  ${recorte('function _mergeListaPorRegistro', 'o merge por registro')}
  ${recorte('function _apagadosLocaisRegistrar', 'o registro das exclusoes locais')}
  ${recorte('function proximoNumeroOS', 'o proximo numero')}
  return { _mergeListaPorRegistro, _apagadosLocaisRegistrar, _osNumerosFaltando,
           _topoDaFilaOS, proximoNumeroOS, _numerosOSExistentes,
           apagados: () => _apagadosAqui };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const os = (id, num) => ({ id, os: num });
const lista = (...v) => JSON.stringify(v);
const numeros = txt => JSON.parse(txt).map(o => o.os).join(',');

const api = monta({ STATE: { ordens: [] } });

console.log('-- o merge nao apaga o que nao foi apagado aqui --');

/* A CENA DE 01/09/2026. Este PC tem ate a 0508. O outro cria a 0509. O
   servidor ja tem as duas; este aparelho nao. */
let m = api._mergeListaPorRegistro(
  lista(os('a', '0507'), os('b', '0508')),                   // base: o que eu vi
  lista(os('a', '0507'), os('b', '0508')),                   // local: o meu STATE
  lista(os('a', '0507'), os('b', '0508'), os('c', '0509')),  // servidor: com a 0509
  null);                                                      // nao apaguei nada
ok('1. a OS criada no outro PC sobrevive ao meu save', numeros(m) === '0507,0508,0509', numeros(m));

/* A REGRESSAO EXATA. A base ficou na frente do STATE — o polling adotou o
   servidor (com a 0509) e nao refez o STATE. Antes do conserto, este merge
   apagava a 0509 do servidor, e era assim que ela sumia. */
m = api._mergeListaPorRegistro(
  lista(os('a', '0507'), os('b', '0508'), os('c', '0509')),  // base: JA com a 0509
  lista(os('a', '0507'), os('b', '0508')),                   // local: STATE velho, sem ela
  lista(os('a', '0507'), os('b', '0508'), os('c', '0509')),  // servidor: com ela
  null);                                                      // ninguem a excluiu
ok('2. base na frente do STATE NAO apaga a OS (a 0509 de 01/09)',
   numeros(m) === '0507,0508,0509', numeros(m));

console.log('');
console.log('-- mas a exclusao de verdade continua valendo --');
m = api._mergeListaPorRegistro(
  lista(os('a', '0507'), os('b', '0508')),
  lista(os('a', '0507')),
  lista(os('a', '0507'), os('b', '0508')),
  new Set(['b']));                       // apaguei a 'b' aqui, de proposito
ok('3. o que foi apagado aqui sai do servidor', numeros(m) === '0507', numeros(m));

/* Excluir aqui e outra pessoa recriar o MESMO id no mesmo intervalo: o registro
   que voltou a existir na minha lista manda, e a exclusao nao o alcanca. */
m = api._mergeListaPorRegistro(
  lista(os('a', '0507'), os('b', '0508')),
  lista(os('a', '0507'), os('b', '0508')),
  lista(os('a', '0507'), os('b', '0508')),
  new Set(['b']));
ok('4. id que voltou a existir na minha lista nao e apagado', numeros(m) === '0507,0508', numeros(m));

console.log('');
console.log('-- saveState anota quem sumiu, e desanota quem voltou --');
api._apagadosLocaisRegistrar('ordens', lista(os('a', '1'), os('b', '2')), lista(os('a', '1')));
ok('5. a exclusao fica anotada', [...(api.apagados().ordens || [])].join(',') === 'b',
   JSON.stringify([...(api.apagados().ordens || [])]));
api._apagadosLocaisRegistrar('ordens', lista(os('a', '1')), lista(os('a', '1'), os('b', '2')));
ok('6. o desfazer traz de volta e a anotacao sai', !api.apagados().ordens,
   JSON.stringify(api.apagados()));
api._apagadosLocaisRegistrar('meta', '{"a":1}', '{"a":2}');
ok('7. chave que nao e lista de registros nao anota nada', !api.apagados().meta,
   JSON.stringify(api.apagados()));

console.log('');
console.log('-- o topo da fila ignora o numero digitado errado --');
const ctx2 = { STATE: { ordens: [], osCounter: 0 } };
const api2 = monta(ctx2);
ctx2.STATE.ordens = [os('a', '0544'), os('b', '0545'), os('c', '0546')];
ok('8. sem engano, o topo e o maior mesmo', api2._topoDaFilaOS() === 546, api2._topoDaFilaOS());
ok('9. e o proximo e o seguinte', api2.proximoNumeroOS() === '0547', api2.proximoNumeroOS());
ctx2.STATE.ordens.push(os('d', '5412'));   // 0542 digitado com um digito a mais
ok('10. um 5412 solto NAO vira o topo da fila', api2._topoDaFilaOS() === 546, api2._topoDaFilaOS());
ok('11. e a numeracao segue de onde parou', api2.proximoNumeroOS() === '0547', api2.proximoNumeroOS());
ctx2.STATE.ordens.push(os('e', '30630123'));   // acima de quatro digitos: nem entra na conta
ok('12. numero acima de 9999 nem conta', api2._topoDaFilaOS() === 546, api2._topoDaFilaOS());

console.log('');
console.log('-- a contagem sabe o que falta --');
const ctx3 = { STATE: { ordens: [] } };
const api3 = monta(ctx3);
/* O retrato de 16/09/2026: o arquivo de papel (faixa larga) e os tres buracos
   do meio da producao. */
const presentes = [];
for (let n = 186; n <= 220; n++) presentes.push(n);
for (let n = 282; n <= 546; n++) if (![292, 293, 296, 509, 536, 537].includes(n)) presentes.push(n);
ctx3.STATE.ordens = presentes.map((n, i) => os('x' + i, formatarNumero(n)));
function formatarNumero(n) { return String(n).padStart(4, '0'); }
const r = api3._osNumerosFaltando();
ok('13. acha os buracos do meio da producao, e so eles',
   r.buracos.join(',') === '292,293,296,509,536,537', r.buracos.join(','));
ok('14. e poe a faixa do arquivo de papel de lado',
   r.arquivo.length === 61 && r.arquivo[0] === 221 && r.arquivo[60] === 281,
   r.arquivo.length + ' de ' + r.arquivo[0] + ' a ' + r.arquivo[r.arquivo.length - 1]);
ok('15. as faixas vem agrupadas, nao numero a numero',
   r.faixas.map(f => f.de + '-' + f.ate).join(' ') === '221-281 292-293 296-296 509-509 536-537',
   r.faixas.map(f => f.de + '-' + f.ate).join(' '));
ok('16. e o total bate com a soma das duas partes',
   r.faltando.length === r.buracos.length + r.arquivo.length, r.faltando.length);

console.log('');
console.log('-- o polling refaz o STATE depois de adotar o servidor --');
/* A outra metade do conserto, provada no proprio arquivo: o loadState nao pode
   ficar DEPOIS do atalho do `_device`, senao o cache anda e o STATE fica. */
const poll = src.slice(src.indexOf('async function verificarServidor'), src.indexOf('function iniciarPolling'));
const iLoad = poll.indexOf('await loadState()');
const iDevice = poll.indexOf("cloudCache._device === DEVICE_ID");
ok('17. o loadState roda ANTES do atalho de "fui eu que gravei"',
   iLoad > 0 && iDevice > 0 && iLoad < iDevice, 'loadState em ' + iLoad + ', atalho em ' + iDevice);

console.log('');
console.log(falhas ? falhas + ' FALHA(S)' : 'todos os testes passaram');
process.exit(falhas ? 1 : 0);
