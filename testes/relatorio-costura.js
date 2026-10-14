/* Rode com:  node testes/relatorio-costura.js

   O RELATÓRIO DE COSTURA (14/10/2026): cada lançamento é "Costureira Vanda, em
   OS 659, uniu ombro de 300 unidades, das 14:00 às 15:00". Este teste guarda:

     · a OS se acha por "659", "0659", "OS659" e "OS 0659";
     · as operações oferecidas são as tarefas da etapa de costura DA OS, na
       ordem do cadastro — a OS de CM.LISA não oferece operação de moletom;
     · só quem tem função de costura entra na lista de costureiras;
     · tempo e ritmo saem do horário (14:00 às 15:00 = 1h, 300 un/h);
     · a semana vai de segunda a DOMINGO (sábado de costura não some).

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);
const cortaArrow = (nome) => recorte(nome, '};\n', nome) + '};';

const motor = [
  'const DASH_DIA_MS = 86400000;',
  "const _DASH_DIAS_SEM = ['dom','seg','ter','qua','qui','sex','sáb'];",
  "const _DASH_MESES_LONGOS = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];",
  recorte('const REL_PROD_ESCALAS', 'const REL_PROD_CHAVE', 'escalas'),
  corta('function _relProdPeriodo'),
  corta('function _relCostPeriodo'),
  corta('function _relCostLista'),
  corta('function _relCostPessoas'),
  corta('function _relCostEtapasDaOS'),
  corta('function _relCostAcharOS'),
  cortaArrow('const _relCostMin'),
  corta('function _relCostDuracao'),
  corta('function _relCostFmtMin'),
  cortaLinha('const _relCostRitmo'),
].join('\n');

const STATE = {
  equipe: [
    { id: 'p1', nome: 'Vanda', funcao: 'Costureira' },
    { id: 'p2', nome: 'Roseane', funcao: 'Auxiliar de costura #1' },
    { id: 'p3', nome: 'João', funcao: 'Operador de enfestadeira' }
  ],
  etapas: [
    { nome: 'Costura CM.LISA | Descalvado', tarefas: [
      { nome: 'União de gola com frente', ordem: 3 }, { nome: 'União de ombro', ordem: 2 }, { nome: 'Acabamento de mangas', ordem: 1 }] },
    { nome: 'Costura BM.TRI | Descalvado', tarefas: [{ nome: 'Fixação de bolso', ordem: 1 }] }
  ],
  ordens: [
    { id: 'o659', os: '0659', etapas: ['Corte', 'Costura CM.LISA | Descalvado', 'Estoque'] },
    { id: 'o660', os: '0660', etapas: ['Corte', 'Estoque'] }
  ]
};
const api = new Function('STATE', motor + `
  return { _relCostPeriodo, _relCostPessoas, _relCostEtapasDaOS, _relCostAcharOS,
           _relCostDuracao, _relCostFmtMin, _relCostRitmo };`)(STATE);

let n = 0;
const ok = (c, msg) => { assert.ok(c, msg); n++; };

// A OS por qualquer jeito de escrever o número.
['659', '0659', 'OS659', 'OS 0659', ' os-659 '].forEach(t =>
  ok((api._relCostAcharOS(t) || {}).id === 'o659', `achar OS por "${t}"`));
ok(api._relCostAcharOS('') === null, 'vazio não acha OS');
ok(api._relCostAcharOS('999') === null, 'OS inexistente');

// As operações da etapa da OS, na ordem do cadastro.
const ets = api._relCostEtapasDaOS(api._relCostAcharOS('659'));
ok(ets.length === 1 && ets[0].nome === 'Costura CM.LISA | Descalvado', 'só a etapa de costura da OS');
ok(JSON.stringify(ets[0].ops) === JSON.stringify(['Acabamento de mangas', 'União de ombro', 'União de gola com frente']), 'operações na ordem do cadastro');
ok(!ets[0].ops.includes('Fixação de bolso'), 'operação de outra linha não aparece');
ok(api._relCostEtapasDaOS(api._relCostAcharOS('660')).length === 0, 'OS sem costura não tem etapa');

// Só funções de costura.
const nomes = api._relCostPessoas().map(p => p.nome);
ok(nomes.includes('Vanda') && nomes.includes('Roseane') && !nomes.includes('João'), 'só função de costura');

// Tempo e ritmo do exemplo do Junior.
const l = { qtd: 300, inicio: '14:00', fim: '15:00' };
ok(api._relCostDuracao(l) === 60, '14:00 às 15:00 = 60 min');
ok(api._relCostRitmo(300, 60) === 300, '300 un em 1h = 300 un/h');
ok(api._relCostFmtMin(60) === '1h' && api._relCostFmtMin(90) === '1h30' && api._relCostFmtMin(45) === '45 min', 'formato do tempo');
ok(api._relCostDuracao({ inicio: '15:00', fim: '14:00' }) === 0, 'fim antes do início não conta tempo');

// Semana de segunda a domingo: 14/10/2026 é quarta.
const per = api._relCostPeriodo('semana', '2026-10-14');
const d = t => new Date(t);
ok(d(per.de).getDay() === 1 && d(per.de).getDate() === 12, 'semana começa na segunda 12/10');
ok(d(per.ate).getDay() === 1 && d(per.ate).getDate() === 19, 'e acaba antes da segunda 19/10 (domingo incluído)');
ok(per.nome === '12/10/2026 a 18/10/2026', 'nome da semana: ' + per.nome);
const dia = api._relCostPeriodo('dia', '2026-10-14');
ok(dia.ate - dia.de === 86400000, 'diário = um dia');

console.log(`relatorio-costura: ${n} verificações ok`);
