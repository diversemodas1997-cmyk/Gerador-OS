/* Rode com:  node testes/relatorio-costura.js

   O RELATÓRIO DE COSTURA (14/10/2026): cada lançamento é "Costureira Vanda, em
   OS 659, uniu ombro de 300 unidades, das 14:00 às 15:00". Este teste guarda:

     · a OS se acha por "659", "0659", "OS659" e "OS 0659";
     · as operações oferecidas são as tarefas da etapa de costura DA OS, na
       ordem do cadastro — a OS de CM.LISA não oferece operação de moletom;
     · só quem tem função de costura entra na lista de costureiras;
     · tempo e ritmo saem do horário (14:00 às 15:00 = 1h, 300 un/h);
     · a semana vai de segunda a DOMINGO (sábado de costura não some);
     · cada unidade (aba) tem as suas OS — pelo status Costurando de lá —, as
       suas costureiras (Equipe da unidade) e as operações da etapa de costura
       daquela unidade.

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
  cortaLinha('const COSTURA_SC_RE'),
  cortaLinha('const _equipeUnidadeDe'),
  cortaLinha('const _relCostStatusDaUnidade'),
  cortaLinha('const _relCostEtapaEhDaUnidade'),
  cortaLinha('const _relCostUnidadeDe'),
  "const _DASH_DIAS_SEM = ['dom','seg','ter','qua','qui','sex','sáb'];",
  "const _DASH_MESES_LONGOS = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];",
  recorte('const REL_PROD_ESCALAS', 'const REL_PROD_CHAVE', 'escalas'),
  corta('function _relProdPeriodo'),
  corta('function _relCostPeriodo'),
  corta('function _relCostPorProduto'),
  corta('function _relCostTotalOp'),
  corta('function _relCostLista'),
  corta('function _relCostPessoas'),
  corta('function _relCostEtapasDaOS'),
  "const _normSku = s => String(s || '').toLowerCase().split('.').map(x => x.trim()).join('.').split(' ').join('');",
  "const linhasSkuDaOS = o => (o.sku ? [o.sku] : []);",
  corta('function _relCostAcharOS'),
  corta('function _relCostOSDaUnidade'),
  cortaArrow('const _relCostMin'),
  corta('function _relCostDuracao'),
  corta('function _relCostFmtMin'),
  cortaLinha('const _relCostRitmo'),
].join('\n');

const STATE = {
  equipe: [
    { id: 'p1', nome: 'Vanda', funcao: 'Costureira' },
    { id: 'p2', nome: 'Roseane', funcao: 'Auxiliar de costura #1' },
    { id: 'p3', nome: 'João', funcao: 'Operador de enfestadeira' },
    { id: 'p4', nome: 'Neide', funcao: 'Costureira', unidade: 'sc' }
  ],
  etapas: [
    { nome: 'Costura CM.LISA | Descalvado', tarefas: [
      { nome: 'União de gola com frente', ordem: 3 }, { nome: 'União de ombro', ordem: 2 }, { nome: 'Acabamento de mangas', ordem: 1 }] },
    { nome: 'Costura BM.TRI | Descalvado', tarefas: [{ nome: 'Fixação de bolso', ordem: 1 }] },
    { nome: 'Costura CM.LISA | São Carlos', tarefas: [{ nome: 'Costura ombro a ombro', ordem: 1 }] }
  ],
  ordens: [
    { id: 'o659', os: '0659', st: 'costurando', etapas: ['Corte', 'Costura CM.LISA | Descalvado', 'Costura CM.LISA | São Carlos', 'Estoque'] },
    { id: 'o660', os: '0660', st: 'ensacado', etapas: ['Corte', 'Estoque'] },
    { id: 'o617', os: '0617', st: 'costurando-sc', etapas: ['Costura CM.LISA | Descalvado', 'Costura CM.LISA | São Carlos'] },
    // CM.LISA sem nenhuma etapa de costura gravada na OS: vêm as do cadastro.
    { id: 'o700', os: '0700', st: 'enfestando', sku: 'cm. lisa', etapas: ['Corte', 'Estoque'] }
  ]
};
const _statusOS = o => o.st;
const api = new Function('STATE', '_statusOS', motor + `
  return { _relCostPeriodo, _relCostPessoas, _relCostEtapasDaOS, _relCostAcharOS, _relCostOSDaUnidade, _relCostUnidadeDe,
           _relCostDuracao, _relCostFmtMin, _relCostRitmo, _relCostTotalOp };`)(STATE, _statusOS);

let n = 0;
const ok = (c, msg) => { assert.ok(c, msg); n++; };

// A OS por qualquer jeito de escrever o número.
['659', '0659', 'OS659', 'OS 0659', ' os-659 '].forEach(t =>
  ok((api._relCostAcharOS(t) || {}).id === 'o659', `achar OS por "${t}"`));
ok(api._relCostAcharOS('') === null, 'vazio não acha OS');
ok(api._relCostAcharOS('999') === null, 'OS inexistente');

// As operações da etapa da OS, na ordem do cadastro.
const ets = api._relCostEtapasDaOS(api._relCostAcharOS('659'), 'desc');
ok(ets.length === 1 && ets[0].nome === 'Costura CM.LISA | Descalvado', 'só a etapa de costura da OS em Descalvado');
const etsSC = api._relCostEtapasDaOS(api._relCostAcharOS('659'), 'sc');
ok(etsSC.length === 1 && etsSC[0].ops[0] === 'Costura ombro a ombro', 'em São Carlos, a etapa de São Carlos');
ok(JSON.stringify(ets[0].ops) === JSON.stringify(['Acabamento de mangas', 'União de ombro', 'União de gola com frente']), 'operações na ordem do cadastro');
ok(!ets[0].ops.includes('Fixação de bolso'), 'operação de outra linha não aparece');
ok(api._relCostEtapasDaOS(api._relCostAcharOS('660')).length === 0, 'OS sem costura não tem etapa');
const ets700 = api._relCostEtapasDaOS(api._relCostAcharOS('700')).map(e => e.nome);
ok(ets700.length === 2 && ets700.includes('Costura CM.LISA | Descalvado') && ets700.includes('Costura CM.LISA | São Carlos'),
   'OS CM.LISA traz as etapas de costura de CM.LISA do cadastro');
ok(!ets700.some(n => /BM\.TRI/.test(n)), 'e não as de outra linha');

// Só funções de costura.
const nomes = api._relCostPessoas().map(p => p.nome);
ok(nomes.includes('Vanda') && !nomes.includes('Roseane') && !nomes.includes('João'), 'só a função Costureira (sem a auxiliar de costura)');
// Por unidade: quem não tem unidade é de Descalvado.
const desc = api._relCostPessoas('desc').map(p => p.nome), sc = api._relCostPessoas('sc').map(p => p.nome);
ok(desc.includes('Vanda') && !desc.includes('Neide'), 'Descalvado: Vanda, sem a Neide');
ok(sc.length === 1 && sc[0] === 'Neide', 'São Carlos: só a Neide');

// As OS de cada aba, pelo status.
ok(api._relCostOSDaUnidade('desc').map(o => o.os).join() === '0659', 'aba Descalvado: OS com status Costurando | Descalvado');
ok(api._relCostOSDaUnidade('sc').map(o => o.os).join() === '0617', 'aba São Carlos: OS com status Costurando | São Carlos');
ok(api._relCostUnidadeDe({ unidade: 'sc' }) === 'sc' && api._relCostUnidadeDe({ etapa: 'Costura CM.LISA | São Carlos' }) === 'sc'
   && api._relCostUnidadeDe({ etapa: 'Costura CM.LISA | Descalvado' }) === 'desc', 'unidade do lançamento');

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

// Mangas: duas por produto (15/10/2026). OS de 670 produtos = 1340 mangas.
ok(api._relCostTotalOp(670, 'Acabamento de mangas') === 1340, 'acabamento de mangas = o dobro dos produtos');
ok(api._relCostTotalOp(670, 'União de mangas 1/2/3') === 1340, 'união de mangas = o dobro dos produtos');
ok(api._relCostTotalOp(670, 'Fechamento de barra/punho') === 1340, 'punhos = o dobro dos produtos');
ok(api._relCostTotalOp(670, 'União de ombro') === 670, 'as outras operações = os produtos');
ok(api._relCostTotalOp(0, 'Acabamento de mangas') === 0, 'OS sem produtos = 0');

console.log(`relatorio-costura: ${n} verificações ok`);
