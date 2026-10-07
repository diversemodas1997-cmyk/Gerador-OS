/* Rode com:  node testes/expedicao-marca-ida.js

   ALOCAR NUMA IDA MARCA "EXPEDIÇÃO DESC X SÃO CARLOS" (16/09/2026, Junior).

   Decidido com ele: marca na hora da alocação e já no primeiro pacote. O que o
   teste guarda:
     · a etapa marcada é a da OS (achada pela mesma regra do campo Em trânsito ·
       IDA), com a hora do momento;
     · já marcada, a hora antiga fica — alocar de novo não reescreve o passado;
     · OS sem essa etapa no checklist não ganha etapa nenhuma;
     · e a chamada só acontece na perna de IDA.

   Recorta do app.js as funções reais. */
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
// A ida virou um atalho de uma linha para _expMarcarExpedicaoOS(os, 'ida').
const cortaLinha = (nome) => recorte(nome, '\n', nome);

// A marcacao automatica preenche tambem as TAREFAS da etapa, e elas vem do
// cadastro (STATE.etapas) — nao do DOM, que nao existe quando quem marca e o
// programa. Por isso o motor agora recebe um STATE.
const montar = (estado) => new Function('STATE', `
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades')}
  ${cortaArr('const FASES_ESTOQUE')}
  ${corta('function tarefasDaEtapa')}
  ${corta('function _tarefasDaEtapaOS')}
  ${corta('function _expMarcarExpedicaoOS')}
  ${cortaLinha('function _expMarcarExpedicaoIdaOS')}
  return _expMarcarExpedicaoIdaOS;
`)(estado);
// Sem cadastro de etapas: e o cenario dos casos de cima, que so olham o pai.
const marcar = montar({ etapas: [], tarefas: [] });

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const osCom = (check, seq) => ({
  id: 'a', os: '0600',
  etapas: ['Corte', 'Ensaque', 'Costura CM.LISA | Descalvado', 'Expedição Desc X São Carlos',
           'Recebido em São Carlos', 'Expedição São Carlos X Desc.', 'Recebido em Descalvado', 'Estoque'],
  progresso: { etapasCheck: check || {}, etapasSeq: seq || {} }
});

const antes = Date.now();
let o = osCom({ 'Corte': true, 'Ensaque': true }, { 'Corte': 1, 'Ensaque': 2 });
let nome = marcar(o);
ok('marca a etapa "Expedição Desc X São Carlos" da OS', nome === 'Expedição Desc X São Carlos'
   && o.progresso.etapasCheck['Expedição Desc X São Carlos'] === true, o.progresso.etapasCheck);
ok('com a hora do momento da alocação', o.progresso.etapasSeq['Expedição Desc X São Carlos'] >= antes, o.progresso.etapasSeq);
ok('e não confunde com a expedição de volta (São Carlos X Desc.)',
   !o.progresso.etapasCheck['Expedição São Carlos X Desc.'], o.progresso.etapasCheck);
ok('não mexe nas outras etapas', o.progresso.etapasSeq['Ensaque'] === 2, o.progresso.etapasSeq);

o = osCom({ 'Expedição Desc X São Carlos': true }, { 'Expedição Desc X São Carlos': 12345 });
nome = marcar(o);
ok('já marcada: devolve vazio e a hora antiga fica', nome === ''
   && o.progresso.etapasSeq['Expedição Desc X São Carlos'] === 12345, o.progresso.etapasSeq);

o = { id: 'b', os: '0601', etapas: ['Corte', 'Ensaque', 'Estoque'], progresso: {} };
nome = marcar(o);
ok('OS sem essa etapa no checklist fica como está', nome === '' && !Object.keys(o.progresso.etapasCheck || {}).length, o.progresso);

o = { id: 'c', os: '0602', etapas: ['Expedição Desc X São Carlos'] };
nome = marcar(o);
ok('OS sem progresso nenhum ganha o progresso com a marca', nome && o.progresso.etapasCheck['Expedição Desc X São Carlos'] === true, o);
ok('sem OS, não quebra', marcar(null) === '', '');

/* NO DIA E HORA DA CARGA, E NAO NA ALOCACAO (07/10/2026, Junior: "a OS ainda
   precisa receber o status em transito antes de receber o status Ensacado Sao
   Carlos"). Alocar nao marca mais nada: a caixa da perna e marcada pela rotina
   _expMarcarViagensVencidas quando chega a hora da carga, com essa hora. */
const salvar = src.slice(src.indexOf("} else if (ctx.tipo === 'carga') {"), src.indexOf("} else if (ctx.tipo === 'volta') {"));
ok('o salvar da carga NAO marca a caixa na hora: chama a rotina da hora da carga',
   !/_expMarcarExpedicaoOS\(/.test(salvar) && /_expRodarViagensVencidas\(\)/.test(salvar), '');
const salvarVolta = src.slice(src.indexOf("} else if (ctx.tipo === 'volta') {"), src.indexOf("} else if (ctx.tipo === 'config') {"));
ok('a volta tambem: nada marcado na hora, a rotina decide',
   !/_expMarcarExpedicaoOS\(/.test(salvarVolta) && /_expRodarViagensVencidas\(\)/.test(salvarVolta), '');

const montarRotina = (estado) => new Function('STATE', `
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades')}
  ${cortaArr('const FASES_ESTOQUE')}
  ${corta('function tarefasDaEtapa')}
  ${corta('function _tarefasDaEtapaOS')}
  ${corta('function _expMarcarExpedicaoOS')}
  ${corta('function _expCancelSet')}
  ${corta('function _expDataEfetivaCarga')}
  ${corta('function _expInstanteCarga')}
  ${corta('function _expCargasDaPernaOS')}
  ${corta('function _expAcertarCaixaAntiga')}
  ${corta('function _osCanceladaParaExpedicao')}
  ${corta('function _expMarcarViagensVencidas')}
  return { rodar: _expMarcarViagensVencidas, instante: _expInstanteCarga };
`)(estado);
const IDA = 'Expedição Desc X São Carlos';
const H = (d, h, m) => new Date(2026, 9, d, h, m || 0).getTime();   // outubro/2026
const janelas = [{ id: 'j1', horaIda: '14:30', horaVolta: '17:00' }];
const cargaDe = (osId, data, extra) => Object.assign({ id: 'c' + osId + data, osId, janelaId: 'j1', data, perna: 'ida' }, extra || {});

let est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [],
  ordens: [osCom({ 'Ensaque': true }, { 'Ensaque': H(6, 9) })],
  expedicaoCargas: [cargaDe('a', '2026-10-09')] };
let R = montarRotina(est);
ok('a hora da carga e o dia dela com a hora da janela (09/10 14:30)', R.instante(est.expedicaoCargas[0]) === H(9, 14, 30), R.instante(est.expedicaoCargas[0]));
ok('antes da carga, nada e marcado: a OS segue Ensacado | Descalvado',
   R.rodar(H(9, 14, 29)) === 0 && !est.ordens[0].progresso.etapasCheck[IDA], est.ordens[0].progresso);
ok('chegada a hora, a caixa e marcada COM A HORA DA CARGA',
   R.rodar(H(9, 15)) === 1 && est.ordens[0].progresso.etapasSeq[IDA] === H(9, 14, 30), est.ordens[0].progresso.etapasSeq);
ok('rodar de novo nao mexe', R.rodar(H(9, 16)) === 0, '');

est.expedicaoExcecoes = [{ janelaId: 'j1', data: '2026-10-09', tipo: 'remarcada', novaData: '2026-10-12' }];
ok('carga remarcada: vale o dia novo', R.instante(est.expedicaoCargas[0]) === H(12, 14, 30), R.instante(est.expedicaoCargas[0]));
est.expedicaoExcecoes = [];

// A OS da regra antiga: caixa marcada na alocacao (hora = criadaEm da carga).
const alocadaEm = H(7, 10);
est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [],
  ordens: [osCom({ 'Ensaque': true, [IDA]: true }, { 'Ensaque': H(6, 9), [IDA]: alocadaEm })],
  expedicaoCargas: [cargaDe('a', '2026-10-09', { criadaEm: new Date(alocadaEm).toISOString() })] };
R = montarRotina(est);
ok('regra antiga, carga ainda por sair: a caixa marcada na alocacao e desmarcada',
   R.rodar(H(8, 9)) === 1 && !est.ordens[0].progresso.etapasCheck[IDA] && !(IDA in est.ordens[0].progresso.etapasSeq), est.ordens[0].progresso);
ok('e marcada de novo no dia e hora da carga', R.rodar(H(9, 15)) === 1 && est.ordens[0].progresso.etapasSeq[IDA] === H(9, 14, 30), est.ordens[0].progresso.etapasSeq);

est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [],
  ordens: [osCom({ [IDA]: true }, { [IDA]: alocadaEm })],
  expedicaoCargas: [cargaDe('a', '2026-10-09', { criadaEm: new Date(alocadaEm).toISOString() })] };
R = montarRotina(est);
ok('regra antiga, carga que ja saiu: a caixa passa a ter a hora da carga',
   R.rodar(H(10, 9)) === 1 && est.ordens[0].progresso.etapasSeq[IDA] === H(9, 14, 30), est.ordens[0].progresso.etapasSeq);

est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [],
  ordens: [osCom({ [IDA]: true, 'Recebido em São Carlos': true }, { [IDA]: alocadaEm, 'Recebido em São Carlos': H(9, 12) })],
  expedicaoCargas: [cargaDe('a', '2026-10-09', { criadaEm: new Date(alocadaEm).toISOString() })] };
R = montarRotina(est);
R.rodar(H(10, 9));
ok('a caixa nunca fica depois da chegada: Em transito sempre antes do Recebido',
   est.ordens[0].progresso.etapasSeq[IDA] < est.ordens[0].progresso.etapasSeq['Recebido em São Carlos'], est.ordens[0].progresso.etapasSeq);

est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [],
  ordens: [osCom({ [IDA]: true }, { [IDA]: H(5, 8) })],
  expedicaoCargas: [cargaDe('a', '2026-10-09', { criadaEm: new Date(alocadaEm).toISOString() })] };
R = montarRotina(est);
ok('caixa marcada por GENTE (hora que nao e a da alocacao) fica como esta',
   R.rodar(H(8, 9)) === 0 && est.ordens[0].progresso.etapasSeq[IDA] === H(5, 8), est.ordens[0].progresso.etapasSeq);

est = { etapas: [], tarefas: [], expedicaoJanelas: janelas, expedicaoExcecoes: [{ janelaId: 'j1', data: '2026-10-09', tipo: 'cancelada' }],
  ordens: [osCom({}, {})], expedicaoCargas: [cargaDe('a', '2026-10-09')] };
R = montarRotina(est);
ok('carga cancelada nao marca nada', R.rodar(H(10, 9)) === 0 && !est.ordens[0].progresso.etapasCheck[IDA], est.ordens[0].progresso);

// A funcao generica escolhe a fase pela perna, e e dai que sai a caixa certa.
const generica = corta('function _expMarcarExpedicaoOS');
ok('a perna escolhe o campo: volta -> transitoVolta, resto -> transitoIda',
   /perna === 'volta' \? 'transitoVolta' : 'transitoIda'/.test(generica), '');

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
