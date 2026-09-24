/* Rode com:  node testes/estoque-com-fio-sc.js

   ESTOQUE COM FIO NAS DUAS UNIDADES (18/09/2026, Junior: "mude o status Estoque
   com fio para Estoque com fio | Descalvado; depois insira o status Estoque com
   fio | São Carlos"; e, sobre o que o dispara: "o fato de o status mudar para
   Estoque com fio | SC faz com que o volume migre para lá").

   A peça costurada em São Carlos fica lá, com os fios soltos, esperando o
   caminhão de volta. O de Descalvado é o que voltou e espera a mesa de limpeza.
   São dois campos, um em cada ponta da viagem.

   O QUE ESTE TESTE GUARDA:

     · o de São Carlos NÃO tem `re`. Nenhuma caixa do checklist significa "com
       fio em São Carlos" — a da costura de lá já acende o Costurando, e a da
       expedição de volta põe a OS na estrada. Quem o acende é o CARIMBO, como
       em Parado e Cancelado. Pôr um `re` aqui seria roubar o estado de outro;
     · e, justamente por ser carimbo, ele vale ATÉ a próxima etapa ser marcada:
       marcar a expedição de volta tira a OS dali sozinha, que é o caminho da
       peça. Isso não é regra nova — é o _statusOS de sempre —, e é o que faz o
       campo se esvaziar sem ninguém limpar nada à mão;
     · o campo de São Carlos e o Costurando de lá dividem a MESMA etapa de
       entrada (a costura). Quem os separa é o status, do mesmo jeito que separa
       o corte ensacado da costura dentro de uma unidade. Se os dois passassem a
       responder à mesma condição, a OS apareceria nos dois ao mesmo tempo.

   Recorta as funções e as listas do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = new Function(`
  ${cortaArr('const STATUS_OS')}
  ${cortaLinha('const STATUS_FIM')}
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades')}
  ${corta('function osEtapaMarcada')}
  ${corta('function _marcasDoStatus')}
  ${corta('function _statusDoChecklistOS')}
  ${corta('function _ultimaMarcacaoChecklist')}
  ${corta('function _statusOS')}
  ${cortaArr('const FASES_ESTOQUE')}
  return { STATUS_OS, FASES_ESTOQUE, _statusOS };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const st = (k) => motor.STATUS_OS.find(x => x.k === k);
const fase = (id) => motor.FASES_ESTOQUE.find(f => f.id === id);

console.log('-- os dois status --');
ok('1. o de Descalvado passou a dizer a unidade no nome',
   st('estoque-fio').rotulo === 'Estoque com fio | Descalvado', st('estoque-fio'));
ok('2. e a CHAVE dele nao mudou (esta gravada nas OS)', !!st('estoque-fio'));
ok('3. o de Sao Carlos existe', !!st('estoque-fio-sc') && st('estoque-fio-sc').rotulo === 'Estoque com fio | São Carlos',
   st('estoque-fio-sc'));
ok('4. os dois baixam o pano, como os outros da fila',
   st('estoque-fio').baixa === true && st('estoque-fio-sc').baixa === true);
ok('5. o de Sao Carlos NAO nasce do checklist (sem `re`)', !st('estoque-fio-sc').re, String(st('estoque-fio-sc').re));
ok('6. e o de Descalvado continua nascendo da caixa de chegada',
   st('estoque-fio').re.test('Recebido em Descalvado'), String(st('estoque-fio').re));

console.log('');
console.log('-- o carimbo manda, ate a proxima etapa --');
const osSC = (check, seq, statusOS, statusOSEm) => ({
  id: 'a', os: '0700',
  etapas: ['Corte', 'Ensaque', 'Expedição Desc X São Carlos', 'Recebido em São Carlos',
           'Costura CM.LISA | São Carlos', 'Expedição São Carlos X Desc.', 'Recebido em Descalvado'],
  statusOS, statusOSEm,
  progresso: { etapasCheck: check || {}, etapasSeq: seq || {} }
});

let o = osSC({ 'Costura CM.LISA | São Carlos': true }, { 'Costura CM.LISA | São Carlos': 1000 },
             'estoque-fio-sc', new Date(5000).toISOString());
ok('7. carimbado, o status e o de Sao Carlos', motor._statusOS(o) === 'estoque-fio-sc', motor._statusOS(o));

o = osSC({ 'Costura CM.LISA | São Carlos': true, 'Expedição São Carlos X Desc.': true },
         { 'Costura CM.LISA | São Carlos': 1000, 'Expedição São Carlos X Desc.': Date.now() },
         'estoque-fio-sc', new Date(5000).toISOString());
ok('8. marcada a expedicao de volta, a folha ganha do carimbo e a OS sai dali',
   motor._statusOS(o) !== 'estoque-fio-sc', motor._statusOS(o));

o = osSC({ 'Costura CM.LISA | São Carlos': true }, { 'Costura CM.LISA | São Carlos': 1000 });
ok('9. sem carimbo, a costura de la continua acendendo o Costurando | Sao Carlos',
   motor._statusOS(o) === 'costurando-sc', motor._statusOS(o));

console.log('');
console.log('-- o campo --');
const f = fase('estoqueFioSC');
ok('10. o campo existe, com o titulo da unidade', !!f && f.titulo === 'Estoque com fio | São Carlos', f && f.titulo);
ok('11. quem manda nele e o status carimbado',
   !!f && f.cond(osSC({}, {}, 'estoque-fio-sc', new Date().toISOString())) === true
   && f.cond(osSC({}, {}, 'fios', new Date().toISOString())) === false);
ok('12. a etapa de entrada e a costura de Sao Carlos (a mesma do Costurando de la)',
   f.entrada.re.test('Costura CM.LISA | São Carlos') && f.entrada.tipo === 'etapa', String(f.entrada.re));
ok('13. e ele fica ANTES do transito de volta, que e por onde a peca sai',
   motor.FASES_ESTOQUE.findIndex(x => x.id === 'estoqueFioSC')
   < motor.FASES_ESTOQUE.findIndex(x => x.id === 'transitoVolta'));
ok('14. o de Descalvado segue entrando pela caixa "Recebido em Descalvado"',
   fase('estoqueFio').entrada.re.test('Recebido em Descalvado'));
ok('15. e os dois campos nunca respondem a mesma condicao',
   fase('estoqueFio').cond(osSC({}, {}, 'estoque-fio-sc', new Date().toISOString())) === false);

console.log('');
console.log('-- as duas telas --');
ok('16. a pagina de Sao Carlos existe, com o painel dela',
   /data-page="estoque-fio-sc"/.test(html) && /id="estoque-fio-sc-painel"/.test(html));
ok('17. e a rota desenha o campo certo',
   /if \(page === 'estoque-fio-sc'\) renderFasePorId\('estoqueFioSC'\);/.test(src));
ok('18. os dois estao na barra lateral, dentro de Estoques',
   /nav-btn[^>]*data-page="estoque-fio"[^>]*>Estoque com fio \| Descalvado</.test(html)
   && /nav-btn[^>]*data-page="estoque-fio-sc"[^>]*>Estoque com fio \| São Carlos</.test(html));
// O grupo Estoques vai ate o grupo seguinte (Operacoes saiu da barra em 24/09/2026).
const _iniEst = html.indexOf('data-group="estoques"');
const grupoEstoques = html.slice(_iniEst, html.indexOf('data-group=', _iniEst + 1));
ok('19. e e o grupo ESTOQUES mesmo, nao Operacoes',
   grupoEstoques.includes('data-page="estoque-fio-sc"'));
ok('20. o Inicio tem o cartao dos dois', /k: 'estoqueFioSC'/.test(src) && /k: 'estoqueFio'/.test(src));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
