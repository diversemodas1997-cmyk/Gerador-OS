/* Rode com:  node testes/observacao-por-autor.js

   A OBSERVAÇÃO DA FOLHA É DE QUEM A ESCREVEU.

   Até 20/08/2026 a folha tinha UMA caixa de observação, sem dono: o texto era
   um campo só da OS (`o.obs`) e quem escrevesse por último apagava o recado do
   anterior. Enquanto só o admin escrevia, isso não aparecia. Abrir a folha para
   a produção transformaria a caixa em cima-e-embaixo — dois turnos, um recado.

   Agora cada pessoa tem A SUA (`o.obsNotas`, uma por login). Junior decidiu que
   NEM O ADMIN edita a dos outros: o que está escrito ali é o que aquela pessoa
   viu acontecer, e recado reescrito por terceiro não vale como registro.

   O que este teste protege:

     · escrever cria a nota DAQUELE login, e não toca em nenhuma outra;
     · reescrever mexe só na própria, inclusive quando o admin é quem reescreve;
     · apagar (texto em branco) tira só a própria da folha;
     · sem login não se escreve — a nota ficaria sem autor, que é o oposto disto;
     · a folha mostra o LOGIN de quem escreveu.

   O teste recorta as funções do app.js de verdade. */
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

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const currentUser = ctx.login ? { email: ctx.login } : null;
  let currentRole = ctx.papel;
  const toast = (m, t) => ctx.toasts.push(t + ': ' + m);
  const saveState = async () => { ctx.gravou = (ctx.gravou || 0) + 1; };
  const document = { getElementById: () => null, querySelector: () => null };
  const exigirEdicaoFolha = () => {
    if (!ctx.servidorNoAr) { toast('nuvem', 'err'); return false; }
    if (currentRole === 'admin' || currentRole === 'usuario') return true;
    toast('sem login', 'err'); return false;
  };
  ${recorte('function _obsNotas', 'a lista de notas')}
  ${recorte('function _obsQuemSou', 'quem sou')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('function _obsQuando', 'a data da nota')}
  ${recorte('async function salvarObsNota', 'a gravacao da nota')}
  return { salvarObsNota, _obsNomeLogin, _obsQuando, _obsNotas };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// Uma OS com o recado do corte já escrito.
const osCom = () => ({
  id: 'os1', os: '0282', obs: '',
  obsNotas: [{ login: 'corte@diverse.com', texto: 'Faltou pano na 3a fase', em: '2026-08-20T14:10:00.000Z' }]
});
const ctxDe = (login, papel, os, servidorNoAr = true) => ({
  login, papel, servidorNoAr, toasts: [], STATE: { ordens: [os] }
});

(async () => {
  console.log('-- escrever a sua --');
  let os = osCom();
  let ctx = ctxDe('costura@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'Ribana veio torta, separei 12 pç');
  ok('1. a nota nova entra com o login de quem escreveu',
    os.obsNotas.length === 2 && os.obsNotas[1].login === 'costura@diverse.com'
    && os.obsNotas[1].texto === 'Ribana veio torta, separei 12 pç', os.obsNotas);
  ok('2. e a do corte fica INTACTA',
    os.obsNotas[0].texto === 'Faltou pano na 3a fase'
    && os.obsNotas[0].em === '2026-08-20T14:10:00.000Z', os.obsNotas[0]);
  ok('3. a nota carimba a hora', !!os.obsNotas[1].em);

  console.log('');
  console.log('-- reescrever --');
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'Corrigindo: foram 50min');
  ok('4. quem escreveu reescreve a propria, sem criar outra',
    os.obsNotas.length === 1 && os.obsNotas[0].texto === 'Corrigindo: foram 50min', os.obsNotas);
  ok('5. e fica marcado que foi editada',
    !!os.obsNotas[0].editadoEm && os.obsNotas[0].em === '2026-08-20T14:10:00.000Z', os.obsNotas[0]);

  console.log('');
  console.log('-- o outro NAO mexe na alheia, nem o admin --');
  os = osCom();
  ctx = ctxDe('costura@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'texto da costura');
  ok('6. escrever do outro lado nao apaga a do corte',
    os.obsNotas.length === 2 && os.obsNotas.some(n => n.texto === 'Faltou pano na 3a fase'), os.obsNotas);
  os = osCom();
  ctx = ctxDe('junior@diverse.com', 'admin', os);
  await monta(ctx).salvarObsNota('os1', 'texto do admin');
  ok('7. o ADMIN tambem so escreve a dele — a do corte continua la',
    os.obsNotas.length === 2
    && os.obsNotas[0].login === 'corte@diverse.com'
    && os.obsNotas[0].texto === 'Faltou pano na 3a fase'
    && os.obsNotas[1].login === 'junior@diverse.com', os.obsNotas);
  os = osCom();
  ctx = ctxDe('junior@diverse.com', 'admin', os);
  await monta(ctx).salvarObsNota('os1', '');
  ok('8. e o admin apagando "a observacao" nao apaga a dos outros',
    os.obsNotas.length === 1 && os.obsNotas[0].login === 'corte@diverse.com', os.obsNotas);

  console.log('');
  console.log('-- apagar a propria --');
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '   ');
  ok('9. texto em branco tira a propria da folha', os.obsNotas.length === 0, os.obsNotas);

  console.log('');
  console.log('-- as bordas --');
  os = osCom();
  ctx = ctxDe('', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'de quem seria?');
  ok('10. sem login nao se escreve — nota sem autor e o oposto da regra',
    os.obsNotas.length === 1 && /quem seria a observação/.test(ctx.toasts.join(' ')), ctx.toasts);
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os, false);
  await monta(ctx).salvarObsNota('os1', 'servidor fora do ar');
  ok('11. servidor da fabrica fora do ar: ninguem escreve',
    os.obsNotas[0].texto === 'Faltou pano na 3a fase', os.obsNotas);
  os = osCom();
  ctx = ctxDe('CORTE@Diverse.com ', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'mesma pessoa, caixa diferente');
  ok('12. maiuscula e espaco nao criam uma segunda pessoa',
    os.obsNotas.length === 1 && os.obsNotas[0].texto === 'mesma pessoa, caixa diferente', os.obsNotas);
  os = { id: 'os1', os: '0282' };
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'primeira da OS');
  ok('13. OS que nunca teve nota nenhuma recebe a primeira',
    Array.isArray(os.obsNotas) && os.obsNotas.length === 1, os.obsNotas);

  console.log('');
  console.log('-- o que a folha mostra --');
  const A = monta(ctxDe('x@y.com', 'usuario', osCom()));
  ok('14. mostra o nome do login', A._obsNomeLogin('costura@diverse.com') === 'costura',
    A._obsNomeLogin('costura@diverse.com'));
  ok('15. login sem arroba aparece inteiro', A._obsNomeLogin('supervisor') === 'supervisor');
  ok('16. sem login, nao inventa nome', A._obsNomeLogin('') === '' && A._obsNomeLogin(null) === '');
  ok('17. a data sai curta, dia e hora',
    /^\d{2}\/\d{2} \d{2}:\d{2}$/.test(A._obsQuando({ em: '2026-08-20T14:10:00' })),
    A._obsQuando({ em: '2026-08-20T14:10:00' }));
  ok('18. e diz quando foi editada',
    / \(editada\)$/.test(A._obsQuando({ em: '2026-08-20T14:10:00', editadoEm: '2026-08-20T16:00:00' })),
    A._obsQuando({ em: '2026-08-20T14:10:00', editadoEm: '2026-08-20T16:00:00' }));

  console.log('');
  console.log('-- o campo cresce em vez de rolar --');
  // A folha não tem barra de rolagem (20/08/2026): o que está rolado para baixo
  // simplesmente não sai no papel, e quem lê não tem como saber que falta. Quem
  // faz o campo crescer é _obsAjustarAltura; aqui se cobra o CSS que tira a
  // rolagem, porque é ele que torna o corte silencioso possível se alguém
  // devolver o overflow sem devolver o crescimento.
  const css = fs.readFileSync(path.join(__dirname, '..', 'styles.css'), 'utf8');
  // A regra BASE (inicio de linha), e nao a `.obs-minha .obs-input` que vem antes.
  const iIn = css.search(/^\.obs-input \{/m);
  const regraInput = css.slice(iIn, css.indexOf('}', iIn));
  ok('19. o campo da observacao nao rola', /overflow:\s*hidden/.test(regraInput), regraInput);
  ok('20. e nao tem alca de arrastar', /resize:\s*none/.test(regraInput), regraInput);
  ok('21. e existe quem faca o campo crescer com o texto',
    /function _obsAjustarAltura/.test(src) && /oninput="_obsAjustarAltura\(this\)"/.test(src));

  console.log('');
  console.log('-- o texto antigo e do admin --');
  // Junior confirmou em 20/08/2026: todas as anotações que já existem, em todas
  // as OS, são do admin — até essa data ninguém mais podia escrever ali.
  // Procura no que a tela DESENHA (o rotulo dentro do span), e nao no arquivo
  // inteiro: o comentario logo acima da funcao cita "sem autor" para explicar
  // por que ele saiu, e isso nao pode derrubar o teste.
  const rotulos = (src.match(/class="obs-quem"[^>]*>([^<]*)</g) || []).join(' | ');
  ok('22. aparece assinado "admin", e nao "sem autor"',
    !/>sem autor</.test(src) && /"obs-quem">admin</.test(src), rotulos);
  ok('23. e continua sendo do admin para editar',
    /exigirEdicao\('editar a observação antiga da OS'\)/.test(src));

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
