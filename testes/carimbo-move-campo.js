/* Rode com:  node testes/carimbo-move-campo.js

   O CARIMBO DO STATUS MOVE O VOLUME DE CAMPO (18/09/2026, Junior: "modifiquei o
   status da OS 0530, de Estoque em trânsito São Carlos para Ensacado São
   Carlos, e a OS não migrou").

   O status já obedecia a esta regra desde que existe: carimbo à mão vale ATÉ a
   próxima etapa ser marcada (_statusOS). O CAMPO não obedecia — olhava só o
   checklist e escolhia o campo da etapa marcada por último. Dava nisto: a lista
   dizia "Ensacado | São Carlos", carimbado agora, e o volume continuava em
   Em trânsito, porque a última CAIXA marcada era a da expedição de ida.

   O QUE ESTE TESTE GUARDA:

     · o carimbo fresco decide o campo, e o campo antigo perde a OS;
     · marcar qualquer caixa DEPOIS devolve a palavra à folha — é o mesmo
       vencimento do carimbo no status, e é o que impede um carimbo velho de
       segurar a OS num lugar que ela já deixou;
     · status que não é lugar (Parado, Cancelado) não move nada;
     · e o carimbo vale como ENTRADA no campo: sem isso a OS apareceria no campo
       certo pela conta do "onde está" e sumiria da lista dele, que só mostra
       quem entrou.

   Recorta as funções e as listas do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = new Function('STATE', `
  ${cortaArr('const STATUS_OS')}
  ${cortaLinha('const STATUS_FIM')}
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades')}
  ${corta('function osEtapaMarcada')}
  ${corta('function _marcasDoStatus')}
  ${corta('function _statusDoChecklistOS')}
  ${corta('function _ultimaMarcacaoChecklist')}
  ${corta('function _statusOS')}
  ${cortaArr('const FASES_ESTOQUE')}
  ${corta('function _faseCarimbadaOS')}
  ${corta('function _faseEntrouOS')}
  ${corta('function _nomeEtapaDaFase')}
  ${cortaLinha('function _faseIdxPorId')}
  ${cortaLinha('const TERMINAL_ETAPA_RE')}
  ${corta('function faseAtualOS')}
  return { FASES_ESTOQUE, faseAtualOS, _faseEntrouOS, _faseCarimbadaOS, _statusOS };
`)({ ordens: [] });

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const idxDe = (id) => motor.FASES_ESTOQUE.findIndex(f => f.id === id);
const campoDe = (o) => { const i = motor.faseAtualOS(o); return i < 0 ? '(fora do fluxo)' : motor.FASES_ESTOQUE[i].titulo; };

/* A OS 0530 como ela estava no banco: Preparo e Corte marcados, depois a caixa
   da expedição de ida — e nada mais. Sem Ensaque, sem recebimento em SC. */
const T0 = 1789590587590;                       // a marca do Corte
const IDA = 1789751086033;                      // a caixa da expedição de ida
const os0530 = (statusOS, statusOSEm) => ({
  id: 'a', os: '0530',
  etapas: ['Preparo Matéria-prima', 'Corte', 'Ensaque', 'Costura CM.LISA | Descalvado',
           'Expedição Desc X São Carlos', 'Recebido em São Carlos', 'Costura CM.LISA | São Carlos',
           'Expedição São Carlos X Desc.', 'Recebido em Descalvado', 'Retirada de fios', 'Estoque'],
  statusOS, statusOSEm,
  progresso: {
    etapasCheck: { 'Preparo Matéria-prima': true, 'Corte': true, 'Expedição Desc X São Carlos': true },
    etapasSeq: { 'Preparo Matéria-prima': T0 - 5000, 'Corte': T0, 'Expedição Desc X São Carlos': IDA }
  }
});

console.log('-- a OS 0530, como ela estava --');
let o = os0530();
ok('1. sem carimbo, ela esta no transito da ida (a ultima caixa marcada)',
   campoDe(o) === 'Em trânsito · IDA', campoDe(o));

console.log('');
console.log('-- carimbada "Ensacado | Sao Carlos" --');
o = os0530('ensacado-sc', new Date(IDA + 60000).toISOString());
ok('2. o status lido e o carimbado', motor._statusOS(o) === 'ensacado-sc', motor._statusOS(o));
ok('3. e o VOLUME migra para o Estoque corte de Sao Carlos',
   campoDe(o) === 'Estoque corte · Unidade São Carlos', campoDe(o));
ok('4. o campo antigo perde a OS', motor.faseAtualOS(o) !== idxDe('transitoIda'), campoDe(o));
ok('5. o carimbo conta como ENTRADA no campo novo — senao a lista dele nao a mostraria',
   motor._faseEntrouOS(o, motor.FASES_ESTOQUE[idxDe('corteSC')]) === true);

console.log('');
console.log('-- o carimbo vence quando a folha fala de novo --');
o = os0530('ensacado-sc', new Date(IDA + 60000).toISOString());
o.progresso.etapasCheck['Recebido em São Carlos'] = true;
o.progresso.etapasSeq['Recebido em São Carlos'] = IDA + 120000;   // marca DEPOIS do carimbo
ok('6. marcada uma caixa depois, quem manda e a folha',
   motor._faseCarimbadaOS(o) === -1, motor._faseCarimbadaOS(o));
ok('7. e ela continua no campo certo, agora pela caixa de chegada',
   campoDe(o) === 'Estoque corte · Unidade São Carlos', campoDe(o));

o = os0530('ensacado-sc', new Date(T0 - 60000).toISOString());     // carimbo ANTERIOR as marcas
ok('8. carimbo mais velho que a folha nao move nada',
   campoDe(o) === 'Em trânsito · IDA', campoDe(o));

console.log('');
console.log('-- status que nao sao lugar --');
['parado', 'cancelado', 'nao-iniciado'].forEach((k, i) => {
  const x = os0530(k, new Date(IDA + 60000).toISOString());
  ok((9 + i) + '. "' + k + '" nao tem campo, entao nao move o volume',
     motor._faseCarimbadaOS(x) === -1, motor._faseCarimbadaOS(x));
});

console.log('');
console.log('-- os outros carimbos, um a um --');
[['ensacado', 'Estoque corte · Unidade Descalvado'],
 ['costurando', 'Costurando · Unidade Descalvado'],
 ['costurando-sc', 'Costurando · Unidade São Carlos'],
 ['estoque-fio', 'Estoque com fio | Descalvado'],
 ['estoque-fio-sc', 'Estoque com fio | São Carlos'],
 ['transito-ida', 'Em trânsito · IDA'],
 ['transito-volta', 'Em trânsito · VOLTA'],
 ['fios', 'Retirada de fios']].forEach(([k, titulo], i) => {
  const x = os0530(k, new Date(IDA + 60000).toISOString());
  ok((12 + i) + '. "' + k + '" leva para "' + titulo + '"', campoDe(x) === titulo, campoDe(x));
});

console.log('');
console.log('-- e a folha continua mandando quando ninguem carimbou --');
o = os0530();
o.progresso.etapasCheck['Recebido em São Carlos'] = true;
o.progresso.etapasSeq['Recebido em São Carlos'] = IDA + 1000;
ok('20. sem carimbo nenhum, a caixa de chegada leva a OS para Sao Carlos',
   campoDe(o) === 'Estoque corte · Unidade São Carlos', campoDe(o));
o = os0530();
o.progresso.etapasCheck['Estoque'] = true;
o.progresso.etapasSeq['Estoque'] = IDA + 2000;
ok('21. e a etapa terminal continua tirando a OS do fluxo',
   campoDe(o) === '(fora do fluxo)', campoDe(o));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
