/* Rode com:  node testes/desfazer.js

   O DESFAZER DO PROGRAMA INTEIRO e o APAGAR DE UMA PASTA DE GRADES.

   Toda gravação passa por saveState(chave), e é lá que o desfazer se engancha:
   antes de escrever, lê o que estava gravado — que é exatamente o estado
   anterior à mexida. Um gancho só cobre cadastro, OS, expedição, planejamento e
   compra, e nada que grava fica de fora por esquecimento.

   O que este teste protege:

     · a foto do ANTES é a de antes da ação, e não a do meio dela (uma ação que
       grava a mesma chave duas vezes não pode perder a primeira foto);
     · gravações coladas no tempo são UMA ação e voltam JUNTAS — desfazer metade
       de uma exclusão seria pior do que não desfazer;
     · o desfazer não se registra a si mesmo (senão o botão viraria um
       liga-desliga e não daria para andar para trás no histórico);
     · a pilha não cresce sem fim;
     · apagar a pasta apaga as grades DELA e só elas, e o apelido de uma
       subpasta que outra pasta ainda usa não vai junto.

   Recorta as funções do app.js de verdade. */
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
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

// O armazenamento de mentira (o DB do app) e o STATE: e tudo o que o desfazer
// toca. `saveState` aqui e o do app.js de verdade, recortado.
const monta = (ctx) => new Function('ctx', `
  const toast = (m, t) => ctx.toasts.push((t || '') + ':' + m);
  const confirm = () => ctx.confirmar;
  const document = { getElementById: () => null, addEventListener: () => {}, querySelector: () => null };
  const STATE = ctx.STATE;
  const DB = {
    get: async (k) => (k in ctx.db ? { value: ctx.db[k] } : null),
    set: async (k, v) => { ctx.db[k] = v; ctx.gravacoes.push(k); }
  };
  let _opHoraFixaVersao = 0;
  const _CHAVES_CONTAB_SNAPSHOT = [], _CHAVES_OE = [];
  ${constante('DESFAZER_MAX')}
  ${constante('DESFAZER_JANELA_MS')}
  ${constante('DESFAZER_BYTES_MAX')}
  ${constante('DESFAZER_NOMES')}
  let _desfazerPilha = [], _desfazendo = false, _desfazerRotulo = '';
  ${recorte('function _desfazerNome', 'o nome da chave')}
  ${recorte('function _desfazerFrase', 'a frase da acao')}
  ${recorte('function _desfazerTamanho', 'o tamanho da pilha')}
  ${recorte('function _desfazerRegistrar', 'o registro da acao')}
  ${recorte('function desfazerNomearAcao', 'o nome explicito da acao')}
  ${recorte('function _desfazerQuando', 'o ha quanto tempo')}
  function _desfazerAtualizarBotao() {}
  ${recorte('async function saveState', 'a gravacao')}
  ${recorte('async function desfazerUltimaAcao', 'o desfazer')}
  const _recusarPorModoNuvem = () => false;
  const goto = () => {};
  return {
    saveState, desfazerUltimaAcao, desfazerNomearAcao,
    pilha: () => _desfazerPilha,
    frase: _desfazerFrase
  };
`)(ctx);

const ctxDe = (dados) => {
  const ctx = { STATE: JSON.parse(JSON.stringify(dados)), db: {}, gravacoes: [], toasts: [], confirmar: true };
  Object.keys(dados).forEach(k => { ctx.db[k] = JSON.stringify(dados[k]); });
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};
const espera = (ms) => new Promise(r => setTimeout(r, ms));

(async () => {
  console.log('-- a foto do antes --');
  let t = ctxDe({ grades: [{ id: 'g1' }, { id: 'g2' }, { id: 'g3' }], ordens: [{ id: 'o1' }] });
  t.ctx.STATE.grades = t.ctx.STATE.grades.filter(g => g.id !== 'g2');
  await t.api.saveState('grades');
  ok('1. gravar empilha uma acao', t.api.pilha().length === 1, String(t.api.pilha().length));
  ok('2. e a acao guarda o estado ANTERIOR inteiro',
     JSON.parse(t.api.pilha()[0].antes.grades).length === 3, t.api.pilha()[0].antes.grades);
  ok('3. o nome sai do que mudou', t.api.pilha()[0].rotulo === 'exclusão de 1 grade', t.api.pilha()[0].rotulo);

  // Segunda gravacao da MESMA chave dentro da janela: a primeira foto e a que vale.
  t.ctx.STATE.grades = [];
  await t.api.saveState('grades');
  ok('4. a segunda gravacao da mesma acao nao apaga a primeira foto',
     t.api.pilha().length === 1 && JSON.parse(t.api.pilha()[0].antes.grades).length === 3,
     JSON.stringify(t.api.pilha().map(a => Object.keys(a.antes))));

  console.log('');
  console.log('-- uma acao, varias chaves --');
  t = ctxDe({ grades: [{ id: 'g1' }], gradeFolderLabels: { tp: { x: 'Nome' } } });
  t.api.desfazerNomearAcao('exclusão da pasta de grades "Camiseta"');
  t.ctx.STATE.grades = [];
  await t.api.saveState('grades');
  t.ctx.STATE.gradeFolderLabels = { tp: {} };
  await t.api.saveState('gradeFolderLabels');
  ok('5. duas chaves coladas no tempo viram UMA acao', t.api.pilha().length === 1,
     JSON.stringify(t.api.pilha().map(a => Object.keys(a.antes))));
  ok('6. com as duas fotas dentro',
     Object.keys(t.api.pilha()[0].antes).sort().join(',') === 'gradeFolderLabels,grades',
     Object.keys(t.api.pilha()[0].antes).join(','));
  ok('7. e com o nome que a acao se deu',
     /pasta de grades "Camiseta"/.test(t.api.pilha()[0].rotulo), t.api.pilha()[0].rotulo);

  await t.api.desfazerUltimaAcao();
  ok('8. desfazer devolve as DUAS chaves',
     t.ctx.STATE.grades.length === 1 && t.ctx.STATE.gradeFolderLabels.tp.x === 'Nome',
     JSON.stringify(t.ctx.STATE));
  ok('9. e grava o que voltou', JSON.parse(t.ctx.db.grades).length === 1, t.ctx.db.grades);
  ok('10. o desfazer NAO se registra a si mesmo (senao viraria liga-desliga)',
     t.api.pilha().length === 0, String(t.api.pilha().length));
  ok('11. e avisa o que foi desfeito', /Desfeito/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));

  console.log('');
  console.log('-- acoes separadas no tempo --');
  t = ctxDe({ tecidos: [{ id: 't1' }], cores: [{ id: 'c1' }] });
  t.ctx.STATE.tecidos.push({ id: 't2' });
  await t.api.saveState('tecidos');
  await espera(1100);                       // passou da janela: outra acao
  t.ctx.STATE.cores.push({ id: 'c2' });
  await t.api.saveState('cores');
  ok('12. gravacoes distantes no tempo sao acoes diferentes', t.api.pilha().length === 2,
     String(t.api.pilha().length));
  ok('13. a inclusao tambem tem nome', t.api.pilha()[1].rotulo === 'inclusão de 1 cor',
     t.api.pilha()[1].rotulo);
  await t.api.desfazerUltimaAcao();
  ok('14. desfazer volta so a ULTIMA acao',
     t.ctx.STATE.cores.length === 1 && t.ctx.STATE.tecidos.length === 2, JSON.stringify(t.ctx.STATE));
  await t.api.desfazerUltimaAcao();
  ok('15. desfazer de novo anda mais um passo para tras',
     t.ctx.STATE.tecidos.length === 1, JSON.stringify(t.ctx.STATE.tecidos));

  console.log('');
  console.log('-- a pilha nao cresce sem fim --');
  t = ctxDe({ tecidos: [] });
  for (let i = 0; i < 20; i++) {
    t.ctx.STATE.tecidos.push({ id: 'x' + i });
    await t.api.saveState('tecidos');
    await espera(2);
  }
  ok('16. a pilha para no teto', t.api.pilha().length <= 12, String(t.api.pilha().length));

  console.log('');
  console.log('-- gravacao que nao muda nada --');
  t = ctxDe({ tecidos: [{ id: 't1' }] });
  await t.api.saveState('tecidos');
  ok('17. salvar sem mudar nada nao entra na pilha', t.api.pilha().length === 0,
     JSON.stringify(t.api.pilha()));

  console.log('');
  console.log('-- a frase de cada tipo de mudanca --');
  const F = t.api.frase;
  ok('18. exclusao de mais de um usa o plural',
     F('ordens', '[1,2,3]', '[1]') === 'exclusão de 2 OS', F('ordens', '[1,2,3]', '[1]'));
  ok('19. mesmo tamanho e "alteracao"',
     F('grades', '[{"a":1}]', '[{"a":2}]') === 'alteração em grades', F('grades', '[{"a":1}]', '[{"a":2}]'));
  ok('20. chave que nao e lista tambem tem frase',
     /configurações/.test(F('meta', '{"a":1}', '{"a":2}')), F('meta', '{"a":1}', '{"a":2}'));

  console.log('');
  console.log('-- apagar a pasta de grades --');
  // As tres funcoes de pasta, com o cadastro de mentira.
  const pastaCtx = {
    STATE: {
      grades: [
        { id: 'g1', nome: 'P ao G3', tipoPeca: 'camiseta', variacao: 'basica' },
        { id: 'g2', nome: 'M-G-GG', tipoPeca: 'camiseta', variacao: 'tricolor' },
        { id: 'g3', nome: '2M-2G', tipoPeca: 'blusa_moletom', variacao: 'tricolor' }
      ],
      ordens: [{ id: 'o1', os: '0190', gradeId: 'g1' }],
      gradeFolderLabels: { tp: { camiseta: 'Camisetas' }, vr: { tricolor: 'Tri' }, tpOrder: ['camiseta'], vrOrder: ['tricolor'] }
    }
  };
  const pastaApi = new Function('ctx', `
    const STATE = ctx.STATE;
    const toast = (m, t) => (ctx.toasts = ctx.toasts || []).push(m);
    const confirm = () => ctx.confirmar;
    const exigirAdmin = () => ctx.admin;
    const saveState = async (k) => { (ctx.gravou = ctx.gravou || []).push(k); };
    const renderGrades = () => {};
    const desfazerNomearAcao = (r) => { ctx.rotulo = r; };
    const _gfl = () => STATE.gradeFolderLabels;
    const labelTp = (k) => (STATE.gradeFolderLabels.tp[k] || k);
    const labelVr = (k) => (STATE.gradeFolderLabels.vr[k] || k);
    const _gradeIdDaOS = (o) => o.gradeId || '';
    const pastasGradeExpandidas = new Set(['tp:camiseta', 'tp:camiseta|var:basica']);
    ${recorte('function _gradesDaPasta', 'as grades da pasta')}
    ${recorte('function _osQueUsamGrades', 'as OS que usam as grades')}
    ${recorte('async function _apagarPastaGrade', 'o apagar da pasta')}
    ${recorte('async function deleteGradeFolder', 'o apagar da pasta de tipo')}
    ${recorte('async function deleteGradeSubfolder', 'o apagar da subpasta')}
    return { deleteGradeFolder, deleteGradeSubfolder, _gradesDaPasta, _osQueUsamGrades, pastasGradeExpandidas };
  `)(pastaCtx);

  pastaCtx.admin = false; pastaCtx.confirmar = true;
  await pastaApi.deleteGradeFolder('camiseta');
  ok('21. quem nao e admin nao apaga pasta', pastaCtx.STATE.grades.length === 3,
     String(pastaCtx.STATE.grades.length));

  pastaCtx.admin = true; pastaCtx.confirmar = false;
  await pastaApi.deleteGradeFolder('camiseta');
  ok('22. e nada acontece sem a confirmacao', pastaCtx.STATE.grades.length === 3,
     String(pastaCtx.STATE.grades.length));

  ok('23. o aviso conta as OS que usam as grades da pasta',
     pastaApi._osQueUsamGrades(['g1']).length === 1, '');

  pastaCtx.confirmar = true;
  await pastaApi.deleteGradeSubfolder('camiseta', 'tricolor');
  ok('24. apagar a subpasta leva so a grade dela',
     pastaCtx.STATE.grades.map(g => g.id).join(',') === 'g1,g3', pastaCtx.STATE.grades.map(g => g.id).join(','));
  ok('25. o apelido "Tri" NAO some: o moletom ainda tem tricolor',
     pastaCtx.STATE.gradeFolderLabels.vr.tricolor === 'Tri',
     JSON.stringify(pastaCtx.STATE.gradeFolderLabels.vr));
  ok('26. a acao se nomeia para o desfazer', /pasta de grades/.test(pastaCtx.rotulo || ''), pastaCtx.rotulo);

  await pastaApi.deleteGradeFolder('camiseta');
  ok('27. apagar a pasta leva as grades dela', pastaCtx.STATE.grades.map(g => g.id).join(',') === 'g3',
     pastaCtx.STATE.grades.map(g => g.id).join(','));
  ok('28. e o apelido e a ordem da pasta vao junto',
     !('camiseta' in pastaCtx.STATE.gradeFolderLabels.tp)
     && !pastaCtx.STATE.gradeFolderLabels.tpOrder.includes('camiseta'),
     JSON.stringify(pastaCtx.STATE.gradeFolderLabels));
  ok('29. as pastas abertas daquele galho fecham',
     !pastaApi.pastasGradeExpandidas.has('tp:camiseta')
     && !pastaApi.pastasGradeExpandidas.has('tp:camiseta|var:basica'),
     JSON.stringify([...pastaApi.pastasGradeExpandidas]));
  ok('30. a OS que usava a grade continua na lista', pastaCtx.STATE.ordens.length === 1, '');
  await pastaApi.deleteGradeFolder('camiseta');
  ok('31. pasta ja vazia avisa e nao faz nada',
     /vazia/.test((pastaCtx.toasts || []).join(' ')), (pastaCtx.toasts || []).join(' | '));

  console.log('');
  console.log('-- na tela --');
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  ok('32. o botao de desfazer fica junto do perfil e nasce escondido',
     /id="btnDesfazer"/.test(html) && /class="auth-icone hidden" id="btnDesfazer"/.test(html), 'botao errado');
  ok('33. Ctrl+Z chama o desfazer', /String\(e\.key\)\.toLowerCase\(\) !== 'z'/.test(src), 'sem atalho');
  ok('34. mas nao dentro de um campo de texto (la o Ctrl+Z e do navegador)',
     /tag === 'INPUT' \|\| tag === 'TEXTAREA'/.test(src), 'atalho roubaria o campo');
  ok('35. a arvore de grades tem o apagar em pasta e subpasta',
     /deleteGradeFolder\(\$\{tpJson\}\)/.test(src) && /deleteGradeSubfolder\(\$\{tpJson\}, \$\{vrJson\}\)/.test(src),
     'faltou o botao na arvore');

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
