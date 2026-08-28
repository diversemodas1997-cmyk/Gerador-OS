/* Rode com:  node testes/checklist-folha.js

   O CHECKLIST DA FOLHA DE OS: a caixa PAI (etapa) e as caixas FILHO (tarefas).

   Junior, 28/08/2026: "insira funcao no programa para que o click na caixa pai
   do checklist da folha de OS preencha automatico as caixas filho."

   Marcar a etapa quer dizer que ela ACABOU, e etapa que acabou tem todas as
   tarefas dela feitas. Quem estava no chao marcava a etapa e depois clicava uma
   a uma nas tarefas, para o papel nao sair pela metade.

   O que este teste protege:

     · marcar o pai marca todos os filhos DAQUELA etapa, e so dela;
     · desmarcar limpa — etapa desmarcada com as tarefas todas marcadas e uma
       folha que se contradiz, e quem le acredita na parte errada;
     · os filhos saem do DOM, nao de uma segunda derivacao da lista de tarefas:
       assim entram tambem as FASES DO CORTE (que nao vem do cadastro da etapa)
       e as tarefas "fora do cadastro" que a OS carrega de quando aquela etapa
       era outra. Derivar de novo deixaria justamente essas de fora;
     · tudo numa gravacao so — o pai e os filhos nao podem ir em duas viagens;
     · sem permissao de editar a folha, nada e marcado.

   O teste recorta a funcao do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(nome) {
  const i = src.indexOf(nome);
  if (i < 0) { console.error('nao achei ' + nome + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + nome); process.exit(1); }
  return src.slice(i, j + 2);
}

// As caixinhas que a folha desenhou. `sub` e a classe das tarefas; o pai nao
// entra aqui, e e de proposito: se ele entrasse, marcar uma etapa mexeria na
// caixa de outra.
const caixa = (etapa, tarefa, checked = false) => ({ dataset: { etapa, tarefa }, checked });

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const document = {
    querySelectorAll: (sel) => sel === '.os-check.sub' ? ctx.caixas : []
  };
  const saveState = async () => { ctx.salvou++; };
  const exigirEdicaoFolha = () => ctx.podeEditar;
  const sincronizarPlanoExpedicaoDaOS = async () => { ctx.sincronizou++; };
  ${corta('async function togglarChecklistEtapa')}
  ${corta('async function togglarChecklistTarefa')}
  return { togglarChecklistEtapa, togglarChecklistTarefa };
`)(ctx);

const ctxDe = (caixas, progresso = {}, podeEditar = true) => {
  const ctx = {
    podeEditar, salvou: 0, sincronizou: 0, caixas,
    STATE: { ordens: [{ id: 'os1', os: '0500', progresso }] }
  };
  return { ctx, api: monta(ctx), os: () => ctx.STATE.ordens[0] };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};
const tarefas = (os, etapa) => Object.keys(((os.progresso || {}).tarefasCheck || {})[etapa] || {}).sort();

(async () => {
  console.log('-- o pai preenche os filhos --');
  {
    const t = ctxDe([
      caixa('Corte', 'Fase 1 · Corpo'),
      caixa('Corte', 'Fase 2 · Gola'),
      caixa('Corte', 'Conferir molde'),
      caixa('Costura', 'Fechar ombro')
    ]);
    await t.api.togglarChecklistEtapa('os1', 'Corte', true);
    ok('1. marcar a etapa marca as tarefas dela',
       tarefas(t.os(), 'Corte').join(' | ') === 'Conferir molde | Fase 1 · Corpo | Fase 2 · Gola',
       JSON.stringify(tarefas(t.os(), 'Corte')));
    ok('2. e a propria etapa fica marcada, como sempre',
       t.os().progresso.etapasCheck.Corte === true);
    ok('3. a etapa VIZINHA nao e tocada',
       tarefas(t.os(), 'Costura').length === 0, JSON.stringify(tarefas(t.os(), 'Costura')));
    ok('4. as caixas na tela acompanham na hora (a folha nao se redesenha por clique)',
       t.ctx.caixas.slice(0, 3).every(c => c.checked === true)
       && t.ctx.caixas[3].checked === false,
       JSON.stringify(t.ctx.caixas.map(c => c.checked)));
    ok('5. tudo numa gravacao so', t.ctx.salvou === 1, String(t.ctx.salvou));
  }

  console.log('');
  console.log('-- e desmarcar limpa, senao a folha se contradiz --');
  {
    const t = ctxDe(
      [caixa('Corte', 'Fase 1 · Corpo', true), caixa('Corte', 'Conferir molde', true)],
      { etapasCheck: { Corte: true }, etapasSeq: { Corte: 1 },
        tarefasCheck: { Corte: { 'Fase 1 · Corpo': true, 'Conferir molde': true } } });
    await t.api.togglarChecklistEtapa('os1', 'Corte', false);
    ok('6. desmarcar a etapa limpa as tarefas dela',
       tarefas(t.os(), 'Corte').length === 0, JSON.stringify(tarefas(t.os(), 'Corte')));
    ok('7. e a etapa sai do progresso (nao fica false engordando o blob)',
       !('Corte' in t.os().progresso.etapasCheck), JSON.stringify(t.os().progresso.etapasCheck));
    ok('8. as caixas da tela desmarcam junto',
       t.ctx.caixas.every(c => c.checked === false),
       JSON.stringify(t.ctx.caixas.map(c => c.checked)));
  }

  console.log('');
  console.log('-- o que so o DOM sabe --');
  {
    // A fase do Corte nao vem do cadastro da etapa (e derivada do enfesto da
    // OS), e a tarefa "fora do cadastro" veio de quando aquela etapa era outra.
    // As duas so existem no que a folha desenhou.
    const t = ctxDe([
      caixa('Corte', 'Fase 3 · Viés'),
      caixa('Corte', 'Tarefa antiga que saiu do cadastro')
    ]);
    await t.api.togglarChecklistEtapa('os1', 'Corte', true);
    ok('9. fase do corte e tarefa orfa entram na conta',
       tarefas(t.os(), 'Corte').length === 2, JSON.stringify(tarefas(t.os(), 'Corte')));
  }
  {
    // Folha fechada (nenhuma caixa desenhada): a etapa e marcada e mais nada.
    // Inventar tarefas sem a folha aberta marcaria coisa que ninguem viu.
    const t = ctxDe([]);
    await t.api.togglarChecklistEtapa('os1', 'Corte', true);
    ok('10. sem caixas na tela, marca so a etapa',
       t.os().progresso.etapasCheck.Corte === true
       && !((t.os().progresso.tarefasCheck || {}).Corte), JSON.stringify(t.os().progresso));
  }

  console.log('');
  console.log('-- a permissao continua mandando --');
  {
    const t = ctxDe([caixa('Corte', 'Fase 1 · Corpo')], {}, false);
    await t.api.togglarChecklistEtapa('os1', 'Corte', true);
    ok('11. sem direito de editar a folha, nada e marcado e nada e gravado',
       !t.os().progresso.etapasCheck && t.ctx.salvou === 0 && t.ctx.caixas[0].checked === false,
       JSON.stringify(t.os().progresso) + ' salvou=' + t.ctx.salvou);
  }

  console.log('');
  console.log('-- marcar UMA tarefa continua sendo so ela --');
  {
    const t = ctxDe([caixa('Corte', 'Fase 1 · Corpo'), caixa('Corte', 'Fase 2 · Gola')]);
    await t.api.togglarChecklistTarefa('os1', 'Corte', 'Fase 1 · Corpo', true);
    ok('12. o filho nao arrasta os irmaos nem marca o pai',
       tarefas(t.os(), 'Corte').join() === 'Fase 1 · Corpo'
       && !(t.os().progresso.etapasCheck || {}).Corte,
       JSON.stringify(t.os().progresso));
  }

  console.log('');
  if (falhas) { console.error(falhas + ' teste(s) falharam'); process.exit(1); }
  console.log('todos os testes passaram');
})();
