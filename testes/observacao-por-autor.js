/* Rode com:  node testes/observacao-por-autor.js

   A OBSERVAÇÃO DA FOLHA É DE QUEM A ESCREVEU.

   Até 20/08/2026 a folha tinha UMA caixa de observação, sem dono: o texto era
   um campo só da OS (`o.obs`) e quem escrevesse por último apagava o recado do
   anterior. Enquanto só o admin escrevia, isso não aparecia. Abrir a folha para
   a produção transformaria a caixa em cima-e-embaixo — dois turnos, um recado.

   Agora cada RECADO é um registro (`o.obsNotas`), com autor e data próprios.
   Junior decidiu que NEM O ADMIN edita o dos outros: o que está escrito ali é o
   que aquela pessoa viu acontecer, e recado reescrito por terceiro não vale
   como registro.

   E, desde 01/09/2026, a mesma pessoa tem VÁRIOS na mesma OS. Antes era um por
   login: a caixa vinha preenchida com o texto anterior e escrever de novo
   passava por cima — o recado de 28/08 das OS 0496 e 0507 virou o de 01/09 e
   sumiu. Junto com isso, EDITAR deixou de mexer na data: `em` é quando o recado
   foi feito e é o que a folha carimba; `editadoEm` só acrescenta "(editada)".

   O que este teste protege:

     · escrever cria a nota DAQUELE login, e não toca em nenhuma outra;
     · escrever DE NOVO acrescenta um recado, não passa por cima do anterior;
     · corrigir um recado mantém a data em que ele foi feito;
     · a folha carimba a data de criação, nunca a da edição;
     · só se corrige recado PRÓPRIO — chave de outro não acha nada;
     · apagar (texto em branco) tira só aquele recado da folha;
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
  ${recorte('function _obsChaveData(', 'a chave de ordenacao')}
  ${recorte('function _obsChave(', 'a identidade da nota')}
  ${recorte('function _obsLogin(', 'o login da nota')}
  ${recorte('function _obsEmOrdem(', 'a ordem das notas')}
  ${recorte('async function salvarObsNota', 'a gravacao da nota')}
  return { salvarObsNota, _obsNomeLogin, _obsQuando, _obsNotas, _obsChave, _obsEmOrdem };
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
  await monta(ctx).salvarObsNota('os1', '', 'Ribana veio torta, separei 12 pc');
  ok('1. a nota nova entra com o login de quem escreveu',
    os.obsNotas.length === 2 && os.obsNotas[1].login === 'costura@diverse.com'
    && os.obsNotas[1].texto === 'Ribana veio torta, separei 12 pc', os.obsNotas);
  ok('2. e a do corte fica INTACTA',
    os.obsNotas[0].texto === 'Faltou pano na 3a fase'
    && os.obsNotas[0].em === '2026-08-20T14:10:00.000Z', os.obsNotas[0]);
  ok('3. a nota carimba a hora', !!os.obsNotas[1].em);

  console.log('');
  console.log('-- escrever DE NOVO acrescenta, nao passa por cima --');
  // O caso das OS 0496 e 0507: o recado de 28/08 virava o de 01/09.
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '', 'Segundo recado, outro dia');
  ok('4. a mesma pessoa passa a ter DOIS recados na mesma OS',
    os.obsNotas.length === 2, os.obsNotas);
  ok('5. o recado antigo continua inteiro, com a data dele',
    os.obsNotas[0].texto === 'Faltou pano na 3a fase'
    && os.obsNotas[0].em === '2026-08-20T14:10:00.000Z'
    && !os.obsNotas[0].editadoEm, os.obsNotas[0]);
  ok('6. e o novo tem data propria, diferente da do antigo',
    !!os.obsNotas[1].em && os.obsNotas[1].em !== os.obsNotas[0].em, os.obsNotas[1]);

  console.log('');
  console.log('-- corrigir mantem a data em que o recado foi feito --');
  os = osCom();
  let A = monta(ctxDe('corte@diverse.com', 'usuario', os));
  await A.salvarObsNota('os1', A._obsChave(os.obsNotas[0]), 'Corrigindo: foram 50min');
  ok('7. corrigir mexe no proprio recado, sem criar outro',
    os.obsNotas.length === 1 && os.obsNotas[0].texto === 'Corrigindo: foram 50min', os.obsNotas);
  ok('8. a data de quando foi feito NAO muda',
    os.obsNotas[0].em === '2026-08-20T14:10:00.000Z', os.obsNotas[0]);
  ok('9. e fica marcado que foi editada', !!os.obsNotas[0].editadoEm, os.obsNotas[0]);
  ok('10. a folha carimba a data de CRIACAO, nao a da edicao',
    A._obsQuando(os.obsNotas[0]).indexOf('20/08') === 0, A._obsQuando(os.obsNotas[0]));

  console.log('');
  console.log('-- o outro NAO mexe na alheia, nem o admin --');
  os = osCom();
  ctx = ctxDe('costura@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '', 'texto da costura');
  ok('11. escrever do outro lado nao apaga a do corte',
    os.obsNotas.length === 2 && os.obsNotas.some(n => n.texto === 'Faltou pano na 3a fase'), os.obsNotas);
  os = osCom();
  A = monta(ctxDe('junior@diverse.com', 'admin', os));
  await A.salvarObsNota('os1', A._obsChave(os.obsNotas[0]), 'o admin reescrevendo a do corte');
  ok('12. o ADMIN com a chave da nota do CORTE nao a reescreve — vira recado dele',
    os.obsNotas.length === 2
    && os.obsNotas[0].login === 'corte@diverse.com'
    && os.obsNotas[0].texto === 'Faltou pano na 3a fase'
    && os.obsNotas[1].login === 'junior@diverse.com', os.obsNotas);
  os = osCom();
  A = monta(ctxDe('junior@diverse.com', 'admin', os));
  await A.salvarObsNota('os1', A._obsChave(os.obsNotas[0]), '');
  ok('13. e o admin apagando a do corte nao apaga nada',
    os.obsNotas.length === 1 && os.obsNotas[0].login === 'corte@diverse.com', os.obsNotas);

  console.log('');
  console.log('-- apagar o proprio --');
  os = osCom();
  A = monta(ctxDe('corte@diverse.com', 'usuario', os));
  await A.salvarObsNota('os1', A._obsChave(os.obsNotas[0]), '   ');
  ok('14. texto em branco tira aquele recado da folha', os.obsNotas.length === 0, os.obsNotas);
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '', '   ');
  ok('15. a caixa em branco do fim, deixada em branco, nao grava nada',
    os.obsNotas.length === 1 && ctx.gravou === undefined, os.obsNotas);

  console.log('');
  console.log('-- a folha e a linha do tempo --');
  os = osCom();
  os.obsNotas.push({ login: 'corte@diverse.com', texto: 'de 25/08', em: '2026-08-25T09:00:00.000Z' });
  os.obsNotas.push({ login: 'costura@diverse.com', texto: 'de 22/08', em: '2026-08-22T09:00:00.000Z' });
  A = monta(ctxDe('corte@diverse.com', 'usuario', os));
  ok('16. sai da mais antiga para a mais recente',
    A._obsEmOrdem(os).map(n => n.texto).join(' > ')
      === 'Faltou pano na 3a fase > de 22/08 > de 25/08',
    A._obsEmOrdem(os).map(n => n.texto));
  os.obsNotas[0].editadoEm = '2026-09-01T10:00:00.000Z';
  ok('17. corrigir hoje NAO joga o recado de 20/08 para o fim da fila',
    A._obsEmOrdem(os)[0].texto === 'Faltou pano na 3a fase',
    A._obsEmOrdem(os).map(n => n.texto));

  console.log('');
  console.log('-- as bordas --');
  os = osCom();
  ctx = ctxDe('', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '', 'de quem seria?');
  ok('18. sem login nao se escreve — nota sem autor e o oposto da regra',
    os.obsNotas.length === 1 && /quem seria a observação/.test(ctx.toasts.join(' ')), ctx.toasts);
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os, false);
  await monta(ctx).salvarObsNota('os1', '', 'servidor fora do ar');
  ok('19. servidor da fabrica fora do ar: ninguem escreve',
    os.obsNotas.length === 1 && os.obsNotas[0].texto === 'Faltou pano na 3a fase', os.obsNotas);
  os = osCom();
  A = monta(ctxDe('CORTE@Diverse.com ', 'usuario', os));
  await A.salvarObsNota('os1', 'corte@diverse.com|2026-08-20T14:10:00.000Z', 'mesma pessoa, caixa diferente');
  ok('20. maiuscula e espaco nao criam uma segunda pessoa',
    os.obsNotas.length === 1 && os.obsNotas[0].texto === 'mesma pessoa, caixa diferente', os.obsNotas);
  os = { id: 'os1', os: '0282' };
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', '', 'primeira da OS');
  ok('21. OS que nunca teve nota nenhuma recebe a primeira',
    Array.isArray(os.obsNotas) && os.obsNotas.length === 1, os.obsNotas);
  os = osCom();
  ctx = ctxDe('corte@diverse.com', 'usuario', os);
  await monta(ctx).salvarObsNota('os1', 'corte@diverse.com|2026-01-01T00:00:00.000Z', 'o recado sumiu de outra maquina');
  ok('22. chave que nao acha nada vira recado NOVO, em vez de perder o texto',
    os.obsNotas.length === 2 && os.obsNotas[1].texto === 'o recado sumiu de outra maquina', os.obsNotas);

  console.log('');
  console.log('-- o que a folha mostra --');
  A = monta(ctxDe('x@y.com', 'usuario', osCom()));
  ok('23. mostra o nome do login', A._obsNomeLogin('costura@diverse.com') === 'costura',
    A._obsNomeLogin('costura@diverse.com'));
  ok('24. login sem arroba aparece inteiro', A._obsNomeLogin('supervisor') === 'supervisor');
  ok('25. sem login, nao inventa nome', A._obsNomeLogin('') === '' && A._obsNomeLogin(null) === '');
  ok('26. a data sai curta, dia e hora',
    /^\d{2}\/\d{2} \d{2}:\d{2}$/.test(A._obsQuando({ em: '2026-08-20T14:10:00' })),
    A._obsQuando({ em: '2026-08-20T14:10:00' }));
  ok('27. e diz quando foi editada',
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
  ok('28. o campo da observacao nao rola', /overflow:\s*hidden/.test(regraInput), regraInput);
  ok('29. e nao tem alca de arrastar', /resize:\s*none/.test(regraInput), regraInput);
  ok('30. e existe quem faca o campo crescer com o texto',
    /function _obsAjustarAltura/.test(src) && /oninput="_obsAjustarAltura\(this\)"/.test(src));

  console.log('');
  console.log('-- o texto antigo e do admin --');
  // Junior confirmou em 20/08/2026: todas as anotações que já existem, em todas
  // as OS, são do admin — até essa data ninguém mais podia escrever ali.
  // Procura no que a tela DESENHA (o rotulo dentro do span), e nao no arquivo
  // inteiro: o comentario logo acima da funcao cita "sem autor" para explicar
  // por que ele saiu, e isso nao pode derrubar o teste.
  const rotulos = (src.match(/class="obs-quem"[^>]*>([^<]*)</g) || []).join(' | ');
  ok('31. aparece assinado "admin", e nao "sem autor"',
    !/>sem autor</.test(src) && /"obs-quem">admin</.test(src), rotulos);
  ok('32. e continua sendo do admin para editar',
    /exigirEdicao\('editar a observação antiga da OS'\)/.test(src));

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
