/* Rode com:  node testes/merge-por-campo.js

   O MESMO REGISTRO MEXIDO NAS DUAS PONTAS JUNTA CAMPO A CAMPO (30/09/2026).

   A caixa "Expedição Desc X São Carlos" sumiu de 10 OS: a alocação a marcou
   numa máquina, outra máquina (com a OS de antes da alocação) marcou o Ensaque
   e a OS dela substituiu a do servidor inteira. O que o teste guarda:
     · a 0571: Ensaque marcado aqui + caixa marcada lá = as duas marcadas;
     · desmarcar aqui continua desmarcando (a chave some);
     · desmarcado lá, e nós sem mexer naquela caixa, fica desmarcado;
     · conflito real no mesmo campo: vale o daqui (a regra de antes);
     · lista é valor inteiro, não se mistura item a item;
     · sem mudança no servidor, o daqui vai como está (o caminho de sempre);
     · registro novo daqui, e o que só existe lá, continuam como antes.

   Recorta a função real do app.js. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const i = src.indexOf('function _mergeListaPorRegistro');
const j = src.indexOf('\n}', i);
if (i < 0 || j < 0) { console.error('nao achei _mergeListaPorRegistro no app.js'); process.exit(1); }
const merge = new Function(src.slice(i, j + 2) + '\nreturn _mergeListaPorRegistro;')();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const S = v => JSON.stringify(v);
const clone = v => JSON.parse(JSON.stringify(v));
const roda = (base, local, srv, apag) => JSON.parse(merge(S(base), S(local), S(srv), apag));
const IDA = 'Expedição Desc X São Carlos';

const os0 = { id: 'a', os: '0571', etapas: ['Corte', 'Ensaque', IDA],
  progresso: { etapasCheck: { Corte: true }, etapasSeq: { Corte: 1 } } };
const outra = { id: 'b', os: '0572', progresso: {} };

// A 0571 de 28/09.
let local = clone(os0); local.progresso.etapasCheck.Ensaque = true; local.progresso.etapasSeq.Ensaque = 300;
let srv = clone(os0); srv.progresso.etapasCheck[IDA] = true; srv.progresso.etapasSeq[IDA] = 200;
srv.progresso.tarefasCheck = { [IDA]: { 'Carregar': true } };
let r = roda([os0, outra], [local, outra], [srv, outra])[0];
ok('Ensaque daqui e caixa da ida de lá convivem', r.progresso.etapasCheck.Ensaque === true
   && r.progresso.etapasCheck[IDA] === true && r.progresso.etapasCheck.Corte === true, r.progresso.etapasCheck);
ok('e as horas das duas também', r.progresso.etapasSeq.Ensaque === 300 && r.progresso.etapasSeq[IDA] === 200, r.progresso.etapasSeq);
ok('e as tarefas filhas marcadas lá', r.progresso.tarefasCheck && r.progresso.tarefasCheck[IDA].Carregar === true, r.progresso);

// Desmarcar aqui.
local = clone(os0); delete local.progresso.etapasCheck.Corte; delete local.progresso.etapasSeq.Corte;
srv = clone(os0); srv.obs = 'nota de lá';
r = roda([os0], [local], [srv])[0];
ok('desmarcado aqui some, e a nota de lá fica', !('Corte' in r.progresso.etapasCheck) && r.obs === 'nota de lá', r);

// Desmarcado lá, nós mexendo em outra coisa.
local = clone(os0); local.obs = 'nota daqui';
srv = clone(os0); delete srv.progresso.etapasCheck.Corte; delete srv.progresso.etapasSeq.Corte;
r = roda([os0], [local], [srv])[0];
ok('desmarcado lá fica desmarcado', !('Corte' in r.progresso.etapasCheck) && r.obs === 'nota daqui', r);

// Conflito real.
local = clone(os0); local.statusOS = 'parado';
srv = clone(os0); srv.statusOS = 'cortando';
r = roda([os0], [local], [srv])[0];
ok('o mesmo campo mudado nas duas pontas: vale o daqui', r.statusOS === 'parado', r.statusOS);

// Lista é valor inteiro.
local = clone(os0); local.etapas = ['Corte', 'Ensaque', IDA, 'Estoque'];
srv = clone(os0); srv.etapas = ['Corte', IDA];
r = roda([os0], [local], [srv])[0];
ok('lista mudada nas duas pontas não se mistura', S(r.etapas) === S(local.etapas), r.etapas);

// Servidor sem mudança: o daqui vai como está.
local = clone(os0); local.obs = 'x';
r = roda([os0], [local], [os0])[0];
ok('servidor igual à base: vai o registro daqui', S(r) === S(local), r);

// Nada mexido aqui: fica o do servidor.
srv = clone(os0); srv.obs = 'lá';
r = roda([os0], [os0], [srv])[0];
ok('não mexemos: fica o do servidor', S(r) === S(srv), r);

// Novo daqui, e o que só existe lá.
const nova = { id: 'c', os: '0600' }, soLa = { id: 'd', os: '0601' };
const lista = roda([os0], [os0, nova], [os0, soLa]);
ok('registro novo daqui entra e o que só existe lá fica', lista.some(x => x.id === 'c') && lista.some(x => x.id === 'd'), lista.map(x => x.id));

// O diário de status (07/10/2026): as duas pontas anotaram trocas diferentes na
// mesma OS. As anotações das duas ficam, sem repetir, pela hora.
const h0 = { id: 'h', os: '7', statusHist: [{ k: 'cortando', em: 100 }] };
const hAqui = { id: 'h', os: '7', statusHist: [{ k: 'cortando', em: 100 }, { k: 'separando', em: 300, c: 'separando' }] };
const hLa = { id: 'h', os: '7', statusHist: [{ k: 'cortando', em: 100 }, { k: 'enfestando', em: 200 }] };
const hj = roda([h0], [hAqui], [hLa]).find(x => x.id === 'h') || {};
ok('o diário de status junta as anotações das duas pontas, pela hora',
   (hj.statusHist || []).map(x => x.k).join(',') === 'cortando,enfestando,separando', hj.statusHist);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
