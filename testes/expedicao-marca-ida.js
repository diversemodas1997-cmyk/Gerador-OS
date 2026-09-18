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

/* A CHAMADA FICA NO SALVAR, E VALE PARA AS DUAS PERNAS (18/09/2026). Ate aqui
   so a ida marcava; com os status de transito (Junior: "esse status deve ser
   alterado automaticamente quando a OS e alocada no plano de expedicao"), a
   volta tambem precisa marcar a caixa dela — senao a OS trazida de Sao Carlos
   continuaria anunciada como se estivesse la. */
const salvar = src.slice(src.indexOf("} else if (ctx.tipo === 'carga') {"), src.indexOf("} else if (ctx.tipo === 'volta') {"));
ok('o salvar da carga marca a caixa da PERNA em que a OS foi alocada',
   /_expMarcarExpedicaoOS\(osAloc, perna === 'volta' \? 'volta' : 'ida'\)/.test(salvar), '');
ok('e grava as OS depois de marcar', /marcou = _expMarcarExpedicaoOS\([\s\S]*saveState\('ordens'\)/.test(salvar), '');

const salvarVolta = src.slice(src.indexOf("} else if (ctx.tipo === 'volta') {"), src.indexOf("} else if (ctx.tipo === 'config') {"));
ok('trazer OS para a volta tambem marca "Expedicao Sao Carlos X Desc."',
   /_expMarcarExpedicaoOS\(osVolta, 'volta'\)/.test(salvarVolta), '');
ok('e so grava as OS quando marcou alguma', /if \(marcadasVolta\) \{[\s\S]*saveState\('ordens'\)/.test(salvarVolta), '');

// A funcao generica escolhe a fase pela perna, e e dai que sai a caixa certa.
const generica = corta('function _expMarcarExpedicaoOS');
ok('a perna escolhe o campo: volta -> transitoVolta, resto -> transitoIda',
   /perna === 'volta' \? 'transitoVolta' : 'transitoIda'/.test(generica), '');

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
