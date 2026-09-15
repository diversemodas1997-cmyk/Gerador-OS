/* Rode com:  node testes/status-os.js

   O STATUS DA OS na coluna AÇÕES da lista de OS Salvas.

   Quatro estados — não iniciado, em andamento, parado, finalizado — e uma regra
   de quem mexe: o ADMIN e o login do ENFESTO/CORTE. É o corte que sabe se a OS
   começou, travou ou terminou; e é a lista de OS Salvas, e nenhuma outra tela,
   que mostra isso.

   O que este teste protege:

     · quem pode mudar (admin e enfesto.corte) e quem só olha — inclusive o
       usuário comum, que registra a folha mas não carimba o status do lote;
     · o nome do login entra de qualquer jeito que tenha sido criado
       ("Enfesto.corte", "enfesto corte", "Enfesto-Corte") e sai na mesma conta;
     · o modo nuvem (servidor da fábrica fora do ar) continua só leitura;
     · "não iniciado" é a AUSÊNCIA dos campos — OS que ninguém tocou não engorda
       o blob, que desce inteiro a cada abertura;
     · quem não pode mudar recebe ETIQUETA, não seletor — e nada é gravado.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
// O cabecalho da folha e marcacao, nao codigo: o lugar do status na janela
// so da para conferir no index.html.
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}
// As listas saem do app.js tambem: o teste nao pode ter a sua propria ideia de
// quais sao os quatro estados nem de qual login manda no status.
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

const monta = (ctx) => new Function('ctx', `
  const toast = (m, t) => ctx.toasts.push(t + ': ' + m);
  const podeGravar = () => ctx.servidorNoAr;
  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const currentUser = ctx.login ? { email: ctx.login } : null;
  let currentRole = ctx.papel;
  const STATE = ctx.STATE;
  const saveState = async () => { ctx.salvou++; };
  const renderListaOS = () => { ctx.redesenhou++; };
  // O status tambem mora no cabecalho da folha de OS desde 02/09/2026, e
  // mudarStatusOS redesenha os dois. Aqui e so um contador: quem prova a caixa
  // da folha e o bloco "a folha mostra o status" la embaixo.
  const renderStatusFolhaOS = () => { ctx.redesenhouFolha = (ctx.redesenhouFolha || 0) + 1; };
  ${constante('STATUS_OS')}
  ${constante('LOGINS_STATUS_OS')}
  ${constante('AREAS_ACESSO')}
  ${constante('ACESSO_PADRAO')}
  ${constante('LOGINS_ESTOQUE_TECIDOS')}
  ${recorte('function _acessosTabela', 'a tabela de acessos')}
  ${recorte('function _acessoChaveConta', 'a chave da conta')}
  ${recorte('function contaTemAcesso', 'o acesso de uma conta')}
  ${recorte('function temAcesso', 'o acesso de quem esta logado')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('function _obsQuando', 'a data da nota')}
  ${recorte('function _chaveLogin', 'a chave do login')}
  ${recorte('function podeMudarStatusOS', 'quem muda o status')}
  ${recorte('function exigirStatusOS', 'a recusa do status')}
  // No celular ninguem grava, nem o admin: aqui o aparelho e um controle
  // do teste, como o papel e o servidor no ar.
  const ehCelular = () => !!ctx.celular;
  ${recorte('function _recusarSomenteLeitura', 'a recusa de quem so le')}
  ${recorte('function _statusOS', 'a leitura do status')}
  ${recorte('function _statusCelulaOS', 'a celula do status')}
  ${recorte('function formatDate', 'a data em dd/mm/aaaa')}
  ${recorte('function _dataFinalizacaoOS', 'a data de finalizacao')}
  ${recorte('function _dataHoraFinalizacaoOS', 'o dia e a hora da finalizacao')}
  ${recorte('function _tituloFinalizacaoOS', 'a dica da data de finalizacao')}
  ${recorte('function _dataCelulaListaOS', 'a celula da coluna Data')}
  ${constante('_STATUS_QUE_BAIXAM')}
  ${recorte('async function aplicarBaixaEstoqueOS', 'a reserva ao salvar a OS')}
  ${recorte('async function _estoqueSeguirStatusOS', 'a baixa de estoque pelo status')}
  ${recorte('async function darBaixaMaterialOS', 'a baixa de material')}
  ${recorte('async function estornarBaixaMaterialOS', 'o estorno da baixa')}
  ${recorte('function _ativaStatusDaOS', 'a ativa de quem foi conjugada a mao')}
  ${recorte('function _conjugadasManuaisDaOS', 'as OS conjugadas a mao')}
  ${recorte('function _conjugadasQueSeguemStatus', 'a conjugada que vai junto')}
  ${recorte('function _carimbarStatusOS', 'a escrita do status numa OS')}
  ${recorte('async function mudarStatusOS', 'a mudanca do status')}
  ${recorte('function conjugadasSemPanoDaOS', 'as conjugadas da lista de reservados')}
  ${recorte('async function salvarConjugarOS', 'o salvar da tela de conjugar')}
  const exigirEstoqueTecidos = () => true;
  /* A TELA DO "CONJUGAR": um DOM de mentira com as caixas marcadas. O que
     importa provar aqui e o que a funcao ESCREVE nas OS — e ela le a marcacao
     do modal, entao o modal precisa existir de alguma forma. */
  let _conjugarOsId = null;
  const openModal = () => {};
  const closeModal = () => { ctx.fechou = (ctx.fechou || 0) + 1; };
  const document = {
    getElementById: (id) => (id === 'conjugarAlinhar' ? { checked: !!ctx.alinhar } : null),
    querySelectorAll: () => (ctx.candidatas || []).map(c => ({
      value: c.id, checked: (ctx.marcadas || []).indexOf(c.id) >= 0
    }))
  };
  // aplicarBaixaEstoqueOS pergunta o consumo da OS ao cadastro; aqui ele vem
  // pronto pelo ctx, que e o que este teste tem a dizer sobre o assunto.
  const consumoAgregadoPorTecidoCor = () => ctx.consumo || [];
  const uid = () => 'm' + (++ctx.seq);
  const _estoqueRedesenharSeAberto = () => {};
  const renderEstoque = () => {};
  return { podeMudarStatusOS, _statusOS, _statusCelulaOS, mudarStatusOS, STATUS_OS,
           darBaixaMaterialOS, estornarBaixaMaterialOS, aplicarBaixaEstoqueOS,
           _dataFinalizacaoOS, _dataHoraFinalizacaoOS, _tituloFinalizacaoOS, _dataCelulaListaOS,
           conjugadasSemPanoDaOS, _conjugadasQueSeguemStatus,
           _ativaStatusDaOS, _conjugadasManuaisDaOS, salvarConjugarOS,
           conjugarNaOS: (id) => { _conjugarOsId = id; } };
`)(ctx);

const ctxDe = (papel, login, servidorNoAr = true, ordens = []) => {
  const ctx = { papel, login, servidorNoAr, toasts: [], salvou: 0, redesenhou: 0, seq: 0,
                STATE: { ordens, meta: {} } };
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

console.log('-- quem pode mudar o status --');
ok('1. admin muda', ctxDe('admin', 'admin@diverse.local').api.podeMudarStatusOS() === true);
ok('2. Enfesto.corte muda',
   ctxDe('usuario', 'enfesto.corte@diverse.local').api.podeMudarStatusOS() === true);
ok('3. usuario comum NAO muda (registra a folha, mas nao carimba o lote)',
   ctxDe('usuario', 'costura@diverse.local').api.podeMudarStatusOS() === false);
ok('4. sem login, nao', ctxDe(null, '').api.podeMudarStatusOS() === false);
ok('5. servidor da fabrica fora do ar: nem o admin',
   ctxDe('admin', 'admin@diverse.local', false).api.podeMudarStatusOS() === false);
ok('6. nem o Enfesto.corte, no modo nuvem',
   ctxDe('usuario', 'enfesto.corte@diverse.local', false).api.podeMudarStatusOS() === false);

console.log('');
console.log('-- o nome do login, digitado de qualquer jeito na criacao da conta --');
['enfesto.corte', 'Enfesto.Corte', 'enfesto-corte', 'enfestocorte'].forEach((n, i) => {
  ok((7 + i) + '. "' + n + '" e a mesma conta',
     ctxDe('usuario', n + '@diverse.local').api.podeMudarStatusOS() === true);
});
ok('11. um login parecido NAO entra de carona',
   ctxDe('usuario', 'enfesto@diverse.local').api.podeMudarStatusOS() === false);

console.log('');
console.log('-- o que fica gravado --');
(async () => {
  let t = ctxDe('admin', 'admin@diverse.local', true, [{ id: 'a1', os: '1234' }]);
  const os = t.ctx.STATE.ordens[0];
  ok('12. OS que ninguem tocou le "nao iniciado"', t.api._statusOS(os) === 'nao-iniciado');
  await t.api.mudarStatusOS('a1', 'andamento');
  ok('13. mudar grava a chave, quem e quando',
     os.statusOS === 'andamento' && os.statusOSPor === 'admin@diverse.local'
     && !isNaN(new Date(os.statusOSEm)), JSON.stringify(os));
  ok('14. e salva no servidor uma vez', t.ctx.salvou === 1, String(t.ctx.salvou));

  await t.api.mudarStatusOS('a1', 'nao-iniciado');
  ok('15. voltar para "nao iniciado" APAGA os campos (nao engorda o blob)',
     !('statusOS' in os) && !('statusOSPor' in os) && !('statusOSEm' in os), JSON.stringify(os));

  t = ctxDe('admin', 'admin@diverse.local', true, [{ id: 'a1', os: '1234', statusOS: 'parado' }]);
  await t.api.mudarStatusOS('a1', 'parado');
  ok('16. escolher o mesmo status nao salva a toa', t.ctx.salvou === 0, String(t.ctx.salvou));

  t = ctxDe('admin', 'admin@diverse.local', true, [{ id: 'a1', os: '1234' }]);
  await t.api.mudarStatusOS('a1', 'inventado');
  ok('17. valor que nao existe cai em "nao iniciado", nao grava lixo',
     !('statusOS' in t.ctx.STATE.ordens[0]), JSON.stringify(t.ctx.STATE.ordens[0]));

  t = ctxDe('usuario', 'costura@diverse.local', true, [{ id: 'a1', os: '1234' }]);
  await t.api.mudarStatusOS('a1', 'finalizado');
  ok('18. quem nao pode mudar nao muda nem salva',
     !('statusOS' in t.ctx.STATE.ordens[0]) && t.ctx.salvou === 0
     && /Enfesto\.corte/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));
  ok('19. e a lista e redesenhada, para o seletor voltar ao que esta gravado',
     t.ctx.redesenhou === 1, String(t.ctx.redesenhou));

  console.log('');
  console.log('-- o que aparece na coluna ACOES --');
  const osParado = { id: 'a1', os: '1234', statusOS: 'parado',
                     statusOSPor: 'enfesto.corte@diverse.local', statusOSEm: '2026-08-26T13:40:00.000Z' };
  let cel = ctxDe('admin', 'admin@diverse.local', true, [osParado]).api._statusCelulaOS(osParado);
  ok('20. quem muda ve um seletor com os quatro estados',
     /^<select/.test(cel) && (cel.match(/<option/g) || []).length === 4, cel);
  ok('21. com o estado gravado ja escolhido',
     /value="parado" selected/.test(cel), cel);
  ok('22. e a dica diz quem mexeu por ultimo',
     /enfesto\.corte/.test(cel) && /Parado/.test(cel), cel);
  cel = ctxDe('usuario', 'costura@diverse.local', true, [osParado]).api._statusCelulaOS(osParado);
  ok('23. quem so olha ve etiqueta, sem seletor',
     /^<span class="os-status ro"/.test(cel) && !/<select/.test(cel), cel);
  ok('24. e o icone do estado aparece nos dois casos', cel.includes('\u{1F534}'), cel);

  console.log('');
  console.log('-- a data de finalizacao (coluna Data, segunda linha) --');
  t = ctxDe('admin', 'admin@diverse.local', true, [{ id: 'a1', os: '1234', data: '2026-03-10' }]);
  const os2 = t.ctx.STATE.ordens[0];
  await t.api.mudarStatusOS('a1', 'andamento');
  ok('33. andamento nao carimba data de finalizacao', !('finalizadaEm' in os2), JSON.stringify(os2));
  await t.api.mudarStatusOS('a1', 'finalizado');
  ok('34. marcar Finalizado carimba o dia',
     typeof os2.finalizadaEm === 'string' && !isNaN(new Date(os2.finalizadaEm)), JSON.stringify(os2));
  const primeira = os2.finalizadaEm;
  // Um respiro: os dois carimbos no MESMO milissegundo dariam a mesma string, e
  // o teste acusaria como "nao recarimbou" algo que recarimbou.
  await new Promise(r => setTimeout(r, 5));
  await t.api.mudarStatusOS('a1', 'parado');
  ok('35. tirar de Finalizado apaga a data (a OS voltou a andar)',
     !('finalizadaEm' in os2), JSON.stringify(os2));
  await t.api.mudarStatusOS('a1', 'finalizado');
  ok('36. marcar de novo carimba o dia NOVO',
     typeof os2.finalizadaEm === 'string' && os2.finalizadaEm !== primeira,
     os2.finalizadaEm + ' vs ' + primeira);

  const A = ctxDe('admin', 'admin@diverse.local', true, []).api;
  let cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10' });
  ok('37. OS sem status mostra so a data em que foi feita',
     /10\/03\/2026/.test(cel2) && !/data-fim/.test(cel2), cel2);
  cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10', statusOS: 'finalizado',
                                finalizadaEm: '2026-08-26T13:40:00.000Z' });
  ok('38. finalizada mostra as duas: a de cima feita, a de baixo finalizada',
     /10\/03\/2026/.test(cel2) && /data-fim/.test(cel2) && /26\/08\/2026/.test(cel2), cel2);
  cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10', statusOS: 'finalizado',
                                statusOSEm: '2026-08-26T13:40:00.000Z' });
  ok('39. finalizada ANTES do campo existir vale o dia do carimbo',
     /26\/08\/2026/.test(cel2), cel2);
  cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10', statusOS: 'parado',
                                statusOSEm: '2026-08-26T13:40:00.000Z' });
  ok('40. status que nao e Finalizado nao vira data de finalizacao',
     !/data-fim/.test(cel2), cel2);

  // A HORA JUNTO COM O DIA (27/08/2026). O instante sempre esteve gravado; era
  // a tela que o cortava. A hora e a LOCAL, a mesma do relogio da fabrica —
  // por isso o teste monta o instante a partir de um Date local, e nao crava
  // "16:26" em cima de um ISO com Z, que mudaria de valor conforme o fuso.
  const instante = new Date(2026, 7, 26, 16, 26, 0);
  const finalizada = { os: '1', data: '2026-03-10', statusOS: 'finalizado',
                       finalizadaEm: instante.toISOString() };
  ok('41. a data de finalizacao leva a hora junto',
     A._dataHoraFinalizacaoOS(finalizada) === '26/08/2026 16:26',
     A._dataHoraFinalizacaoOS(finalizada));
  ok('42. e ela aparece assim na coluna Data da lista',
     /26\/08\/2026 16:26/.test(A._dataCelulaListaOS(finalizada)),
     A._dataCelulaListaOS(finalizada));
  ok('43. OS que nao terminou nao tem dia nem hora',
     A._dataHoraFinalizacaoOS({ os: '1', data: '2026-03-10', statusOS: 'andamento',
                                statusOSEm: instante.toISOString() }) === '');
  ok('44. data gravada que nao e data nao vira hora inventada',
     A._dataHoraFinalizacaoOS({ os: '1', statusOS: 'finalizado', finalizadaEm: 'nao e data' }) === '');
  ok('45. a dica separa a hora REAL da hora do carimbo em lote',
     A._tituloFinalizacaoOS(finalizada) === 'Dia e hora em que a OS foi finalizada'
     && A._tituloFinalizacaoOS({ statusOS: 'finalizado', statusOSEm: instante.toISOString() })
        === 'Dia e hora em que a OS foi marcada como finalizada');

  console.log('');
  console.log('-- o pano sai do estoque quando a OS comeca a andar --');
  // A reserva nasce ao salvar a OS (aplicarBaixaEstoqueOS, fora deste teste);
  // aqui o que se prova e o que o STATUS faz com ela.
  const comMov = (st) => {
    const t2 = ctxDe('admin', 'admin@diverse.local', true,
                     [{ id: 'e1', os: '900', data: '2026-03-10', statusOS: st }]);
    t2.ctx.STATE.estoqueMov = [
      { id: 'm1', origem: 'os', osId: 'e1', kg: 10, status: 'reservado' },
      { id: 'm2', origem: 'os', osId: 'e1', kg: 5, status: 'reservado' },
      { id: 'm3', origem: 'nf', kg: 99, tipo: 'entrada' }
    ];
    return t2;
  };
  const situacao = ctx => ctx.STATE.estoqueMov.filter(m => m.origem === 'os').map(m => m.status).join('+');
  let e = comMov();
  await e.api.mudarStatusOS('e1', 'andamento');
  ok('46. "em andamento" baixa o pano sozinho', situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'finalizado');
  ok('47. finalizado tambem baixa (quem pulou o andamento ja gastou o pano)',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'parado');
  ok('48. parado idem: parou DEPOIS de comecar', situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov('andamento');
  e.ctx.STATE.estoqueMov.forEach(m => { if (m.origem === 'os') m.status = 'consumido'; });
  await e.api.mudarStatusOS('e1', 'nao-iniciado');
  ok('49. voltar para "nao iniciado" estorna: a OS nao gastou pano nenhum',
     situacao(e.ctx) === 'reservado+reservado', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'andamento');
  ok('50. a entrada de NF nao e tocada por nada disso',
     e.ctx.STATE.estoqueMov.find(m => m.id === 'm3').status === undefined
     && e.ctx.STATE.estoqueMov.find(m => m.id === 'm3').kg === 99);

  // E o outro lado da mesma regra: o movimento NASCE conforme o status, porque o
  // consumo e recalculado toda vez que a OS e salva — e uma OS em producao pode
  // ser salva a qualquer momento (corrigir uma camada na folha, por exemplo).
  const salvando = async (st) => {
    const t3 = ctxDe('admin', 'admin@diverse.local', true,
                     [{ id: 's1', os: '901', data: '2026-03-10', statusOS: st }]);
    t3.ctx.STATE.estoqueMov = [];
    t3.ctx.consumo = [{ tecidoNome: 'Malha', corNome: 'Preto', kg: 12 }];
    await t3.api.aplicarBaixaEstoqueOS(t3.ctx.STATE.ordens[0]);
    return t3.ctx.STATE.estoqueMov.map(m => m.status).join('+');
  };
  ok('51. OS nao iniciada: o pano nasce RESERVADO', await salvando(undefined) === 'reservado');
  ok('52. OS em andamento salva de novo: o pano nasce ja BAIXADO',
     await salvando('andamento') === 'consumido', await salvando('andamento'));
  ok('53. e finalizada tambem — corrigir a folha nao desfaz a baixa',
     await salvando('finalizado') === 'consumido', await salvando('finalizado'));

  /* O PANO NAO VOLTA PARA A PRATELEIRA NO MEIO DO CAMINHO (Junior, 27/08/2026):
     "se as OS mudam para status parado ou voltam para em andamento, isso nao faz
     os tecidos reservados voltarem para o estoque reservado". E o que a fabrica
     ve: o rolo foi cortado no enfesto; a OS parar depois disso nao remonta o
     rolo. So "nao iniciado" — a OS que nao comecou — devolve a reserva. */
  e = comMov();
  await e.api.mudarStatusOS('e1', 'andamento');
  await e.api.mudarStatusOS('e1', 'parado');
  ok('54. andamento -> parado: o pano continua baixado',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  await e.api.mudarStatusOS('e1', 'andamento');
  ok('55. e voltando a andar tambem — nada volta para reservado',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  await e.api.mudarStatusOS('e1', 'finalizado');
  ok('56. ate o fim da OS, um caminho so: baixado continua baixado',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  await e.api.mudarStatusOS('e1', 'nao-iniciado');
  ok('57. e so "nao iniciado" devolve a reserva',
     situacao(e.ctx) === 'reservado+reservado', situacao(e.ctx));

  console.log('');
  console.log('-- a OS conjugada nao reserva pano (o enfesto e o mesmo) --');
  /* Junior, 28/08/2026: "os de grade conjugada nao reserva tecido, pois esse
     tipo de os representa a fase corpo 2 cm.rec. e o mesmo enfesto, separado em
     duas os diferentes. O tecido reservado correspondente e o da fase corpo 2".

     Ate 28/08 isso acontecia por ACIDENTE: as grades conjugadas estavam sem
     comprimento e largura, entao a passiva nascia com 0 x 0 e kg zero. Quatro
     das cinco ja foram medidas — o acidente acabou, e sem a guarda a proxima OS
     conjugada salva dobraria a reserva do tecido. O teste existe porque o erro
     nao apareceria na tela: apareceria no "Disponivel" de um pano que ninguem
     ia tirar da prateleira. */
  const salvandoConj = async (osExtra) => {
    const t4 = ctxDe('admin', 'admin@diverse.local', true,
                     [Object.assign({ id: 'c1', os: '902', data: '2026-03-10' }, osExtra)]);
    t4.ctx.STATE.estoqueMov = [];
    t4.ctx.consumo = [{ tecidoNome: 'Malha Algodao', corNome: 'Branco', kg: 106.722 }];
    await t4.api.aplicarBaixaEstoqueOS(t4.ctx.STATE.ordens[0]);
    return t4.ctx.STATE.estoqueMov;
  };
  ok('58. a OS ATIVA continua reservando o pano das duas',
     (await salvandoConj({})).length === 1);
  ok('59. a PASSIVA (conjugadaPaiId) nao gera movimento nenhum',
     (await salvandoConj({ conjugadaPaiId: 'pai' })).length === 0,
     JSON.stringify(await salvandoConj({ conjugadaPaiId: 'pai' })));

  // A guarda fica DEPOIS do filtro por osId de proposito: passiva que ja tenha
  // movimento gravado (nascido antes desta regra) e LIMPA ao salvar de novo, em
  // vez de a reserva velha ficar pendurada para sempre.
  const t5 = ctxDe('admin', 'admin@diverse.local', true,
                   [{ id: 'c2', os: '903', data: '2026-03-10', conjugadaPaiId: 'pai' }]);
  t5.ctx.STATE.estoqueMov = [
    { id: 'velho', origem: 'os', osId: 'c2', kg: 106.722, status: 'reservado' },
    { id: 'nf', origem: 'nf', kg: 99, tipo: 'entrada' }
  ];
  t5.ctx.consumo = [{ tecidoNome: 'Malha Algodao', corNome: 'Branco', kg: 106.722 }];
  await t5.api.aplicarBaixaEstoqueOS(t5.ctx.STATE.ordens[0]);
  ok('60. e a reserva antiga de uma passiva e apagada ao salvar de novo',
     t5.ctx.STATE.estoqueMov.length === 1
     && t5.ctx.STATE.estoqueMov[0].id === 'nf',
     JSON.stringify(t5.ctx.STATE.estoqueMov));

  console.log('');
  console.log('-- a conjugada segue TODOS os status da ativa --');
  /* Junior, 28/08/2026: "insira funcao no programa que altera status da os
     conjugada para finalizada, quando a os ativa tem seu status alterado para
     finalizada". As duas sao o MESMO enfesto: o pano e estendido uma vez e
     cortado uma vez. Terminar uma e deixar a outra aberta descreve um trabalho
     que nao existe — e era o que acontecia, porque quem carimba carimba a OS
     que esta olhando. */
  const parDe = (st) => ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: st },
    { id: 'so', os: '0500' }
  ]);
  let p = parDe();
  await p.api.mudarStatusOS('at', 'finalizado');
  const pa = () => p.ctx.STATE.ordens[1];
  ok('68. finalizar a ativa finaliza a conjugada', pa().statusOS === 'finalizado',
     JSON.stringify(pa()));
  ok('69. e carimba a data nela tambem, que e o que a coluna Data le',
     !!pa().finalizadaEm && !isNaN(new Date(pa().finalizadaEm)), pa().finalizadaEm);
  ok('70. as duas viajam na MESMA gravacao (uma so, nao duas)',
     p.ctx.salvou === 1, String(p.ctx.salvou));
  ok('71. e o aviso diz que a conjugada foi junto',
     p.ctx.toasts.some(t => /0497/.test(t)), JSON.stringify(p.ctx.toasts));

  // So a ATIVA arrasta. Sem esta trava, duas OS que se apontassem ficariam se
  // carimbando em circulo — a mesma razao do guard em deveGerarConjugada.
  p = parDe();
  await p.api.mudarStatusOS('pa', 'finalizado');
  ok('72. a PASSIVA nao arrasta a ativa', p.ctx.STATE.ordens[0].statusOS === undefined,
     JSON.stringify(p.ctx.STATE.ordens[0]));

  // TODOS os estados propagam (Junior, 28/08/2026: "a os conjugada deve seguir
  // todas as alteracoes de status da os ativa"). Nao ha estado em que uma
  // esteja e a outra nao: e o mesmo enfesto na mesa.
  p = parDe();
  await p.api.mudarStatusOS('at', 'andamento');
  ok('73. "em andamento" propaga tambem', pa().statusOS === 'andamento', JSON.stringify(pa()));
  p = parDe();
  await p.api.mudarStatusOS('at', 'parado');
  ok('73b. e "parado" idem', pa().statusOS === 'parado', JSON.stringify(pa()));

  // Ja finalizada nao e recarimbada: a data dela e o dia em que ela terminou.
  p = parDe('finalizado');
  p.ctx.STATE.ordens[1].finalizadaEm = '2026-01-01T10:00:00.000Z';
  await p.api.mudarStatusOS('at', 'finalizado');
  ok('74. conjugada ja finalizada mantem a data dela',
     pa().finalizadaEm === '2026-01-01T10:00:00.000Z', pa().finalizadaEm);

  // OS sozinha continua sozinha, e OS cuja irma sumiu nao quebra nada.
  p = parDe();
  await p.api.mudarStatusOS('so', 'finalizado');
  ok('75. OS sem conjugada segue seu caminho', p.ctx.STATE.ordens[2].statusOS === 'finalizado');
  const orfa = ctxDe('admin', 'admin@diverse.local', true,
                     [{ id: 'x', os: '0499', conjugadaId: 'sumiu' }]);
  await orfa.api.mudarStatusOS('x', 'finalizado');
  ok('76. conjugada excluida depois: finaliza a ativa e nao reclama',
     orfa.ctx.STATE.ordens[0].statusOS === 'finalizado' && orfa.ctx.salvou === 1);
  ok('77. _conjugadasQueSeguemStatus devolve LISTA (uma ativa pode puxar mais de uma)',
     Array.isArray(orfa.api._conjugadasQueSeguemStatus(orfa.ctx.STATE.ordens[0], 'finalizado')));

  /* DESFAZER TAMBEM ACOMPANHA (Junior, 28/08/2026). Tirar a ativa de
     "Finalizado" e deixar a conjugada finalizada travaria a dupla: dali em
     diante so a mao desfaria a segunda, e a lista mostraria metade de um
     enfesto terminada e metade nao. */
  const parFinalizado = () => {
    const t = ctxDe('admin', 'admin@diverse.local', true, [
      { id: 'at', os: '0498', conjugadaId: 'pa' },
      { id: 'pa', os: '0497', conjugadaPaiId: 'at' }
    ]);
    return t;
  };
  let d = parFinalizado();
  await d.api.mudarStatusOS('at', 'finalizado');
  await d.api.mudarStatusOS('at', 'nao-iniciado');
  const dp = () => d.ctx.STATE.ordens[1];
  ok('78. tirar a ativa de finalizado tira a conjugada tambem',
     dp().statusOS === undefined, JSON.stringify(dp()));
  ok('79. e apaga a data de finalizacao dela junto',
     !('finalizadaEm' in dp()), JSON.stringify(dp()));

  // A conjugada vai para o MESMO estado da ativa, nao para um estado escolhido
  // aqui: as duas sao o mesmo enfesto, e ficar em estados diferentes e o que se
  // esta consertando.
  d = parFinalizado();
  await d.api.mudarStatusOS('at', 'finalizado');
  await d.api.mudarStatusOS('at', 'andamento');
  ok('80. finalizado -> em andamento: a conjugada vai para em andamento tambem',
     dp().statusOS === 'andamento' && !('finalizadaEm' in dp()), JSON.stringify(dp()));

  /* A conjugada e ALINHADA a ativa, venha ela de onde vier. Antes so o par
     finalizado->desfazer acompanhava, e uma conjugada em estado proprio ficava
     para tras — meio enfesto num estado, meio noutro. */
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'finalizado' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'parado' }
  ]);
  await d.api.mudarStatusOS('at', 'andamento');
  ok('81. conjugada em estado proprio e alinhada a ativa',
     d.ctx.STATE.ordens[1].statusOS === 'andamento', JSON.stringify(d.ctx.STATE.ordens[1]));

  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'andamento' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'finalizado' }
  ]);
  await d.api.mudarStatusOS('at', 'parado');
  ok('82. andamento -> parado leva a conjugada junto, e apaga a finalizacao dela',
     d.ctx.STATE.ordens[1].statusOS === 'parado'
     && !('finalizadaEm' in d.ctx.STATE.ordens[1]), JSON.stringify(d.ctx.STATE.ordens[1]));

  // O unico caso em que ela NAO e tocada: ja estar no estado pedido. Recarimbar
  // reescreveria a data de finalizacao dela — que e o dia em que ELA terminou.
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'andamento' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'parado', statusOSEm: 'ontem' }
  ]);
  await d.api.mudarStatusOS('at', 'parado');
  ok('83. conjugada ja no estado pedido nao e recarimbada',
     d.ctx.STATE.ordens[1].statusOSEm === 'ontem', JSON.stringify(d.ctx.STATE.ordens[1]));

  /* ----------------------------------------------------------------------
     CONJUGAR OS À MÃO (10/09/2026, Junior: "o usuário deve ser capaz de
     conjugar duas ou mais OS, sem interferir no cadastro das grades de cada
     OS" — e, antes disso, "de forma que elas respondam pela mudanca de status
     da OS ativa").

     A amarra da grade só serve para OS que ainda não existem. Esta serve para
     as que já estão na lista: marca-se `conjugadaStatusPaiId` na OS que segue,
     e mais nada. Campo separado do `conjugadaPaiId` da grade DE PROPÓSITO —
     aquele também significa "não reserva pano", e conjugar duas OS à mão não
     pode zerar a reserva de tecido de ninguém em silêncio.
     ---------------------------------------------------------------------- */
  console.log('');
  console.log('-- conjugar OS a mao --');
  const maoDe = (extra) => ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at' },
    { id: 'c', os: '0512', conjugadaStatusPaiId: 'at' },
    { id: 'so', os: '0600' }
  ].concat(extra || []));

  p = maoDe();
  await p.api.mudarStatusOS('at', 'andamento');
  ok('93. as OS conjugadas a mao seguem a ativa',
     p.ctx.STATE.ordens[1].statusOS === 'andamento' && p.ctx.STATE.ordens[2].statusOS === 'andamento',
     JSON.stringify(p.ctx.STATE.ordens.slice(1, 3)));
  ok('94. e sao duas ou mais, nao so uma', p.api._conjugadasManuaisDaOS(p.ctx.STATE.ordens[0]).length === 2);
  ok('95. numa gravacao so', p.ctx.salvou === 1, String(p.ctx.salvou));
  ok('96. a OS que nao foi conjugada nao e tocada',
     p.ctx.STATE.ordens[3].statusOS === undefined);

  // Finalizar e desfazer: o grupo inteiro vai e volta.
  p = maoDe();
  await p.api.mudarStatusOS('at', 'finalizado');
  ok('97. finalizar leva todas, com a data',
     p.ctx.STATE.ordens.slice(1, 3).every(o => o.statusOS === 'finalizado' && !!o.finalizadaEm));
  await p.api.mudarStatusOS('at', 'nao-iniciado');
  ok('98. desfazer traz todas de volta, e apaga a data',
     p.ctx.STATE.ordens.slice(1, 3).every(o => o.statusOS === undefined && o.finalizadaEm === undefined),
     JSON.stringify(p.ctx.STATE.ordens.slice(1, 3)));

  // Quem segue nao arrasta: e a mesma regra da conjugada da grade.
  p = maoDe();
  await p.api.mudarStatusOS('b', 'parado');
  ok('99. a OS que segue nao arrasta a ativa nem a irma',
     p.ctx.STATE.ordens[0].statusOS === undefined && p.ctx.STATE.ordens[2].statusOS === undefined,
     JSON.stringify(p.ctx.STATE.ordens));

  // A que ja esta no estado pedido nao e recarimbada: a data dela e o dia em
  // que ELA terminou.
  p = maoDe();
  p.ctx.STATE.ordens[1].statusOS = 'finalizado';
  p.ctx.STATE.ordens[1].finalizadaEm = '2026-09-02T10:00:00.000Z';
  await p.api.mudarStatusOS('at', 'finalizado');
  ok('100. a que ja estava finalizada mantem a data dela',
     p.ctx.STATE.ordens[1].finalizadaEm === '2026-09-02T10:00:00.000Z');

  /* AS DUAS AMARRAS SE MISTURAM. A OS conjugada a mao pode ter a conjugada
     DELA pela grade — carimbar metade do conjunto e o mesmo pe quebrado de
     sempre. Por isso o caminho e percorrido em largura. */
  p = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at', conjugadaId: 'bf' },
    { id: 'bf', os: '0499', conjugadaPaiId: 'b' }
  ]);
  await p.api.mudarStatusOS('at', 'andamento');
  ok('101. a conjugada DA CONJUGADA tambem vai (a mao puxa a da grade)',
     p.ctx.STATE.ordens[2].statusOS === 'andamento', JSON.stringify(p.ctx.STATE.ordens[2]));

  // Ciclo: A segue B e B segue A. Sem a trava do "ja visto" isto rodaria para
  // sempre — e o programa inteiro para junto.
  p = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0435', conjugadaStatusPaiId: 'b' },
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at' }
  ]);
  await p.api.mudarStatusOS('at', 'parado');
  ok('102. duas OS que se apontam nao entram em loop',
     p.ctx.STATE.ordens[1].statusOS === 'parado');

  /* O TECIDO NAO SE MEXE. E a diferenca que separa esta amarra da outra: a
     conjugada da GRADE e a fase 2 do mesmo enfesto e nao reserva pano; a
     conjugada A MAO e uma OS inteira, com o pano dela na prateleira. Zerar a
     reserva aqui sumiria com material de verdade, em silencio. */
  const comPano = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at' }
  ]);
  comPano.ctx.consumo = [{ tecidoNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', kg: 50 }];
  await comPano.api.aplicarBaixaEstoqueOS(comPano.ctx.STATE.ordens[1]);
  ok('103. a OS conjugada a mao CONTINUA reservando o pano dela',
     (comPano.ctx.STATE.estoqueMov || []).length === 1
     && comPano.ctx.STATE.estoqueMov[0].kg === 50,
     JSON.stringify(comPano.ctx.STATE.estoqueMov));
  // E a da grade continua nao reservando, como sempre.
  const semPano = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500', conjugadaPaiId: 'at' }
  ]);
  semPano.ctx.consumo = [{ tecidoNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', kg: 50 }];
  await semPano.api.aplicarBaixaEstoqueOS(semPano.ctx.STATE.ordens[1]);
  ok('104. e a conjugada da GRADE segue sem reservar (o pano esta na ativa)',
     (semPano.ctx.STATE.estoqueMov || []).length === 0,
     JSON.stringify(semPano.ctx.STATE.estoqueMov));
  // A lista de material reservado tambem nao pode confundir as duas: a linha
  // filha "o pano esta na OS X" so vale para a da grade.
  ok('105. a lista de reservados nao trata a conjugada a mao como filha',
     comPano.api.conjugadasSemPanoDaOS('at', new Set()).length === 0);

  /* O QUE A TELA ESCREVE. Marcar e desmarcar caixas e o unico caminho pelo qual
     a amarra manual nasce e morre — um erro aqui amarra a OS errada, e o status
     de amanha vai junto com ela. */
  console.log('');
  console.log('-- o salvar da tela de conjugar --');
  const telaDe = (ordens) => {
    const feito = ctxDe('admin', 'admin@diverse.local', true, ordens);
    feito.ctx.candidatas = ordens;
    return feito;
  };

  let tc = telaDe([
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500' },
    { id: 'c', os: '0512' },
    { id: 'so', os: '0600' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['b', 'c'];
  await tc.api.salvarConjugarOS();
  ok('106. marcar duas OS amarra as duas a ativa',
     tc.ctx.STATE.ordens[1].conjugadaStatusPaiId === 'at'
     && tc.ctx.STATE.ordens[2].conjugadaStatusPaiId === 'at',
     JSON.stringify(tc.ctx.STATE.ordens));
  ok('107. e nao encosta em quem nao foi marcado',
     !('conjugadaStatusPaiId' in tc.ctx.STATE.ordens[3]));
  ok('108. gravou uma vez e fechou a janela', tc.ctx.salvou === 1 && tc.ctx.fechou === 1);
  ok('109. e NAO carimbou status nenhum: conjugar nao e carimbar',
     tc.ctx.STATE.ordens.every(o => o.statusOS === undefined),
     JSON.stringify(tc.ctx.STATE.ordens));

  // Desmarcar solta a OS, e ela fica exatamente como estava.
  tc = telaDe([
    { id: 'at', os: '0435' },
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at', statusOS: 'andamento' },
    { id: 'c', os: '0512', conjugadaStatusPaiId: 'at' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['c'];
  await tc.api.salvarConjugarOS();
  ok('110. desmarcar solta a OS do grupo',
     !('conjugadaStatusPaiId' in tc.ctx.STATE.ordens[1])
     && tc.ctx.STATE.ordens[2].conjugadaStatusPaiId === 'at',
     JSON.stringify(tc.ctx.STATE.ordens));
  ok('111. e a OS solta fica com o status que ja tinha',
     tc.ctx.STATE.ordens[1].statusOS === 'andamento');

  /* O ALINHAR E OPCIONAL, e por isso: carimbar sozinho reescreveria a data de
     finalizacao de quem terminou em outro dia. Em branco, ninguem e tocado
     hoje; marcado, as escolhidas entram no estado da ativa agora. */
  tc = telaDe([
    { id: 'at', os: '0435', statusOS: 'andamento' },
    { id: 'b', os: '0500' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['b'];
  await tc.api.salvarConjugarOS();
  ok('112. sem alinhar, a OS conjugada nao muda de estado hoje',
     tc.ctx.STATE.ordens[1].statusOS === undefined, JSON.stringify(tc.ctx.STATE.ordens[1]));

  tc = telaDe([
    { id: 'at', os: '0435', statusOS: 'andamento' },
    { id: 'b', os: '0500' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['b'];
  tc.ctx.alinhar = true;
  await tc.api.salvarConjugarOS();
  ok('113. com alinhar, ela entra no estado da ativa agora',
     tc.ctx.STATE.ordens[1].statusOS === 'andamento', JSON.stringify(tc.ctx.STATE.ordens[1]));

  // Quem so consulta nao conjuga: e a mesma permissao de carimbar o status.
  tc = telaDe([{ id: 'at', os: '0435' }, { id: 'b', os: '0500' }]);
  tc.ctx.papel = 'consulta';
  tc.ctx.login = 'costura@diverse.local';
  const semPermissao = monta(t.ctx);
  semPermissao.conjugarNaOS('at');
  tc.ctx.marcadas = ['b'];
  await semPermissao.salvarConjugarOS();
  ok('114. quem nao pode carimbar status tambem nao conjuga',
     !('conjugadaStatusPaiId' in tc.ctx.STATE.ordens[1]) && tc.ctx.salvou === 0,
     JSON.stringify(tc.ctx.STATE.ordens[1]));

  console.log('');
  console.log('-- e ela APARECE na lista, dizendo onde o pano esta --');
  /* Junior, 28/08/2026: "mostra a passiva na lista como conjugada, o pano esta
     na ativa". Nao reservar e certo, mas sumir sem explicacao e o que faz
     alguem procurar a OS na lista de material e concluir que o pano dela foi
     esquecido. Ela volta como linha filha da ativa. */
  const t6 = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'pai',    os: '0498', data: '2026-08-20' },
    { id: 'fil',    os: '0497', data: '2026-08-20', conjugadaPaiId: 'pai' },
    { id: 'outra',  os: '0499', data: '2026-08-20' }
  ]);
  const F = t6.api.conjugadasSemPanoDaOS;
  ok('61. a ativa traz a conjugada dela', F('pai', new Set()).map(o => o.os).join() === '0497');
  ok('62. uma OS sem conjugada nao traz ninguem', F('outra', new Set()).length === 0);
  ok('63. sem pai nao ha o que parear', F('', new Set()).length === 0
     && F(undefined, new Set()).length === 0);
  ok('64. a passiva que AINDA tem movimento proprio nao se repete aqui',
     F('pai', new Set(['fil'])).length === 0);
  ok('65. e a lista tambem aceita um array de ids, nao so um Set',
     F('pai', ['fil']).length === 0 && F('pai', []).length === 1);

  // A linha filha e montada dentro da tabela de reservados, colada na ativa —
  // se alguem separar as duas em tabelas diferentes, a leitura que o Junior
  // pediu ("o pano esta na de cima") se perde e o teste cai.
  const secao = src.slice(src.indexOf('OSs · material reservado'));
  ok('66. a filha sai logo depois da linha da ativa, na mesma tabela',
     /linhaOS\(p\.pai\)\s*\+\s*p\.filhas\.map\(c => linhaConjugada\(c, p\.pai\)\)/.test(secao),
     secao.slice(secao.indexOf('<tbody>'), secao.indexOf('<tbody>') + 200));
  ok('67. e ela diz em qual OS o pano esta, com o numero da ativa',
     /o pano está na OS \$\{esc\(pai\.osNumero\)/.test(src));

  console.log('');
  console.log('-- o filtro por status da lista de OS Salvas --');
  // O <select> de mentira: e tudo o que _filtroStatusListaOS toca (value e
  // innerHTML), entao da para conferir as opcoes sem navegador.
  const comSelect = (escolhido, ordens) => {
    const sel = { value: escolhido, innerHTML: '' };
    const ctx = { papel: 'admin', login: 'admin@diverse.local', servidorNoAr: true,
                  toasts: [], salvou: 0, redesenhou: 0, STATE: { ordens }, sel };
    const api = new Function('ctx', `
      const document = { getElementById: () => ctx.sel };
      const esc = (s) => String(s == null ? '' : s);
      ${constante('STATUS_OS')}
      ${recorte('function _statusOS', 'a leitura do status')}
      ${recorte('function _filtroStatusListaOS', 'o filtro por status')}
      return { _filtroStatusListaOS };
    `)(ctx);
    return { ctx, api, sel };
  };
  const osDoFiltro = [
    { id: '1', os: '0483' },                          // nao iniciado
    { id: '2', os: '0484', statusOS: 'parado' },
    { id: '3', os: '0485', statusOS: 'finalizado' },
    { id: '4', os: '0486', statusOS: 'finalizado' }
  ];
  let f = comSelect('', osDoFiltro);
  ok('27. sem escolha, o filtro nao corta nada', f.api._filtroStatusListaOS(osDoFiltro) === '');
  ok('28. as opcoes sao "todos" mais os quatro estados',
     (f.sel.innerHTML.match(/<option/g) || []).length === 5, f.sel.innerHTML);
  ok('29. cada opcao ja diz quantas OS tem naquele estado',
     /Todos os status \(4\)/.test(f.sel.innerHTML)
     && /Finalizado \(2\)/.test(f.sel.innerHTML)
     && /Parado \(1\)/.test(f.sel.innerHTML)
     && /Não iniciado \(1\)/.test(f.sel.innerHTML), f.sel.innerHTML);
  f = comSelect('finalizado', osDoFiltro);
  ok('30. o escolhido volta como chave e continua marcado',
     f.api._filtroStatusListaOS(osDoFiltro) === 'finalizado'
     && /value="finalizado" selected/.test(f.sel.innerHTML), f.sel.innerHTML);

  console.log('');
  console.log('-- o status mora em DOIS lugares, e monta num so --');
  /* Ate 02/09/2026 ele morava so na coluna ACOES da lista de OS Salvas, e este
     teste dizia isso. Junior pediu o status tambem no cabecalho da janela da
     folha: quem carimba e o corte, e o corte trabalha com a folha aberta.

     A regra que sobra — e que este teste guarda — e a que importa: os dois
     lugares montam pela MESMA _statusCelulaOS. Duas telas desenhando o status
     por caminhos diferentes e como uma ganha um estado que a outra nao tem.
     Se aparecer um uso novo, o teste cai e a conversa acontece antes.

     15/09/2026: entrou o TERCEIRO lugar — a lista de OS de cada campo do
     fluxo (Estoque de corte, Costurando, Em transito...), que ganhou a mesma
     coluna de acoes da lista de OS Salvas. A regra segue de pe, que e o que
     este teste guarda: os tres montam pela MESMA _statusCelulaOS. */
  const usos = (src.match(/_statusCelulaOS\(/g) || []).length;
  ok('31. _statusCelulaOS e chamada em tres lugares (a definicao + os tres)',
     usos === 4, String(usos));
  const campo = src.slice(src.indexOf('function renderFaseOsLista'));
  ok('31b. a terceira esta na coluna col-actions da lista de um campo do fluxo',
     /col-actions row-actions">\s*\$\{_statusCelulaOS\(o\)\}/.test(campo.slice(0, 4000)),
     campo.slice(campo.indexOf('col-actions'), campo.indexOf('col-actions') + 120));
  const lista = src.slice(src.indexOf('function renderListaOS'));
  ok('32. uma chamada esta na coluna col-actions da lista',
     /col-actions row-actions">\s*\$\{_statusCelulaOS\(o\)\}/.test(lista),
     lista.slice(lista.indexOf('col-actions'), lista.indexOf('col-actions') + 120));
  const folha = src.slice(src.indexOf('function renderStatusFolhaOS'));
  ok('33. a outra esta no cabecalho da folha, e nao redesenha OS que sumiu',
     /_statusCelulaOS\(o, 'folha'\)/.test(folha.slice(0, 1600))
     && /if \(!o\) \{ box\.innerHTML = ''; return; \}/.test(folha.slice(0, 1600)),
     folha.slice(0, 200));
  ok('34. e o cabecalho da folha e .no-print — status nao e dado do papel',
     /id="print-status-os"/.test(html)
     && /page-header no-print[\s\S]{0,900}id="print-status-os"/.test(html));

  /* O BOTAO CONJUGAR NA BARRA DA FOLHA (10/09/2026, Junior: "insira a acao do
     botao conjugar na barra de tarefas da janela de visualizacao da OS").

     Mesma barra e mesma permissao do status, pela mesma razao: quem esta com a
     folha aberta e o corte, e mandar voltar a lista para amarrar duas OS do
     mesmo trabalho e o passo a mais que ja tinha trazido o status para ca. */
  ok('35. o botao esta na barra .no-print da folha, e nasce escondido',
     /page-header no-print[\s\S]{0,1600}id="print-conjugar-os"/.test(html)
     && /class="btn hidden" id="print-conjugar-os"/.test(html), 'botao fora da barra');
  ok('36. ele chama a acao, que abre a MESMA janela da lista',
     /onclick="conjugarOsAtual\(\)"/.test(html)
     && /function conjugarOsAtual\(\)[\s\S]{0,200}abrirModalConjugarOS\(printOsAtual\.id\)/.test(src),
     'acao desligada do botao');
  ok('37. quem nao carimba status nao ve o botao (e a OS que sumiu o esconde)',
     /bt\.classList\.toggle\('hidden', !\(o && podeMudarStatusOS\(\)\)\)/.test(folha.slice(0, 1600)),
     folha.slice(0, 1200));
  ok('38. e o botao e acertado ANTES dos returns, senao a barra ficaria mentindo',
     folha.indexOf("print-conjugar-os") < folha.indexOf("if (!box) return;"),
     'o toggle ficou depois do return');

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
