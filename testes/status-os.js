/* Rode com:  node testes/status-os.js

   O STATUS DA OS na coluna AÇÕES da lista de OS Salvas.

   15/09/2026 o status deixou de ser genérico. Eram quatro estados — não
   iniciado, em andamento, parado, finalizado — e viraram a ETAPA em que a OS
   está, nesta ordem: Não iniciado, Preparando matéria-prima, Enfestando,
   Cortando, Ensacado, Costurando | Descalvado, Costurando | São Carlos,
   Retirando fio, Parado, Estoque. Só "Parado" fica fora da fila: o checklist
   diz a etapa e ele diz que ela travou ali.

   "ENSACADO" É O FIM DO CORTE. O "Finalizado" à parte existiu por uma tarde e
   saiu: dois jeitos de dizer que algo acabou — a etapa e um carimbo dizendo que
   acabou — é a porta para a lista dizer uma coisa e a prateleira outra. Quem
   carimba `finalizadaEm` é "Ensacado" (STATUS_FIM), porque a data responde por
   uma OPERAÇÃO e não pela OS: ensacar é o último ato do corte, e dali em diante
   a peça é da costura.

   E o status passou a NASCER DO CHECKLIST, com carimbo à mão por cima que vale
   até a próxima etapa ser marcada. É o miolo novo deste teste: a derivação, a
   validade do carimbo, a data de fim de quem terminou pela FOLHA (sem ninguém
   carimbar), e a garantia de que a troca de nomes não soltou o pano que já
   estava baixado (o "em andamento" que baixava o estoque agora se chama
   "enfestando", e as OS gravadas com a chave velha leem o checklist).

   O que este teste protege, e já protegia:

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
  ${constante('STATUS_FIM')}
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
  // O status nasce do checklist: a derivacao e as duas funcoes que ela usa
  // entram inteiras, sem duble — o que se quer provar aqui e justamente que a
  // etapa marcada vira o texto do status.
  ${recorte('function osEtapaMarcada', 'a etapa marcada no checklist')}
  ${src.match(/^const ETAPA_SC_RE = .+$/m)[0]}
  ${src.match(/^const _osRecebidaSC = .+$/m)[0]}
  ${src.match(/^const COSTURA_SC_RE = .+$/m)[0]}
  ${src.match(/^const _osCosturaEmSC = .+$/m)[0]}
  ${recorte('function _marcasDoStatus', 'as marcas de um status')}
  ${recorte('function _statusDoChecklistOS', 'o status que o checklist diz')}
  ${recorte('function _ultimaMarcacaoChecklist', 'a ultima etapa marcada')}
  ${recorte('function _statusOS', 'a leitura do status')}
  ${constante('STATUS_PONTO')}
  ${recorte('function _statusPingo', 'o pingo do status')}
  ${recorte('function _statusEstilo', 'o fundo da caixa do status')}
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
  return { podeMudarStatusOS, _statusOS, _statusDoChecklistOS, _statusCelulaOS, mudarStatusOS, STATUS_OS,
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
  await t.api.mudarStatusOS('a1', 'enfestando');
  ok('13. mudar grava a chave, quem e quando',
     os.statusOS === 'enfestando' && os.statusOSPor === 'admin@diverse.local'
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
  await t.api.mudarStatusOS('a1', 'estoque');
  ok('18. quem nao pode mudar nao muda nem salva',
     !('statusOS' in t.ctx.STATE.ordens[0]) && t.ctx.salvou === 0
     && /Enfesto\.corte/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));
  ok('19. e a lista e redesenhada, para o seletor voltar ao que esta gravado',
     t.ctx.redesenhou === 1, String(t.ctx.redesenhou));

  console.log('');
  console.log('-- o status NASCE DO CHECKLIST --');
  /* A folha ja e marcada etapa por etapa; o status le aquilo em vez de pedir o
     mesmo apontamento duas vezes. `etapasSeq` e o carimbo de QUANDO cada etapa
     foi marcada (Date.now() no app), e e ele que decide qual vale. */
  const osCheck = (etapas, check, seq) => ({
    id: 'c1', os: '0501', etapas,
    progresso: { etapasCheck: check, etapasSeq: seq }
  });
// A UNIDADE DA COSTURA SAI DO NOME DA ETAPA (15/09/2026). Os 34 desenhos listam
// AS DUAS — "Costura CM.LISA | Descalvado" e "| São Carlos" —, e quem está no
// chão marca a que fez. A "Costura" pura é a etapa das OS antigas, de antes de
// a segunda unidade existir, e conta como Descalvado.
  // "Recebido em Descalvado" entrou na lista em 17/09/2026, quando ela passou a
  // acender um status proprio (Estoque com fio). Sem estar aqui, a caixa era
  // ignorada pelo `leitura` e os testes dela passavam sem provar nada.
  const FLUXO = ['Preparo de matéria-prima', 'Enfesto', 'Corte',
                 'Recebido em São Carlos', 'Costura', 'Costura CM.LISA | São Carlos',
                 'Recebido em Descalvado', 'Retirada de fios', 'Ensaque', 'Estoque'];
  const leitura = (check, seq) => {
    const o = osCheck(FLUXO, check, seq);
    return ctxDe('admin', 'admin@diverse.local', true, [o]).api._statusOS(o);
  };
  ok('12a. nada marcado: nao iniciado', leitura({}, {}) === 'nao-iniciado');
  // O nome do status e o da etapa nao sao iguais — "Enfestando" vem de
  // "Enfesto", "Preparando materia-prima" de "Preparo de materia-prima". A
  // ligacao e por regex, e e isso que estas linhas guardam.
  ok('12b. "Preparo de materia-prima" acende Preparando materia-prima',
     leitura({ 'Preparo de matéria-prima': true }, { 'Preparo de matéria-prima': 1 }) === 'materia-prima',
     leitura({ 'Preparo de matéria-prima': true }, { 'Preparo de matéria-prima': 1 }));
  ok('12c. "Enfesto" acende Enfestando',
     leitura({ 'Enfesto': true }, { 'Enfesto': 2 }) === 'enfestando');

  /* O ENFESTO VEM DA TABELA DE ENFESTOS, e não do checklist (15/09/2026,
     Junior: "o status Enfestando deve estar correlacionado com a etapa Enfesto
     que já existe na folha de OS, pelos check box das etapas fase 1, fase 2,
     fase 3"). Não há etapa "Enfesto" no cadastro — ela é de outro tipo: uma
     linha por FASE da grade, cada uma com a sua caixa (progresso.enfestosCheck).
     Era por isso que o status existia e nunca acendia sozinho. */
  const comEnfesto = (check, seq, etapasCheck, etapasSeq) => {
    const o = osCheck(FLUXO, etapasCheck || {}, etapasSeq || {});
    o.progresso.enfestosCheck = check;
    o.progresso.enfestosSeq = seq;
    return ctxDe('admin', 'admin@diverse.local', true, [o]).api._statusOS(o);
  };
  ok('12c-1. a fase 1 do enfesto marcada acende Enfestando',
     comEnfesto({ 1: true }, { 1: 5000 }) === 'enfestando',
     comEnfesto({ 1: true }, { 1: 5000 }));
  // Qualquer fase serve: marcar a 2 diz a mesma coisa que marcar a 1 — ha pano
  // estendido na mesa agora. Exigir a 1 faria o status mentir quando alguem
  // marcasse fora de ordem, que e coisa de chao de fabrica.
  ok('12c-2. qualquer fase do enfesto acende, nao so a primeira',
     comEnfesto({ 2: true }, { 2: 5000 }) === 'enfestando');
  ok('12c-3. nenhuma fase marcada: o enfesto nao acende',
     comEnfesto({}, {}) === 'nao-iniciado');
  // E o Corte, marcado DEPOIS, passa a frente: as duas fontes disputam pelo
  // mesmo carimbo de relogio.
  ok('12c-4. Corte marcado depois do enfesto: vence Cortando',
     comEnfesto({ 1: true }, { 1: 5000 }, { 'Corte': true }, { 'Corte': 9000 }) === 'cortando');
  ok('12c-5. e marcado ANTES, o enfesto segue mandando',
     comEnfesto({ 1: true }, { 1: 9000 }, { 'Corte': true }, { 'Corte': 5000 }) === 'enfestando');
  // Fase marcada antes desta versao nao tem carimbo: vale o desempate por ordem,
  // que poe o enfesto entre o preparo e o corte, onde ele esta no chao.
  ok('12c-6. fase sem carimbo ainda acende, pela ordem da fila',
     comEnfesto({ 1: true }, {}) === 'enfestando',
     comEnfesto({ 1: true }, {}));
  ok('12d. "Corte" acende Cortando',
     leitura({ 'Enfesto': true, 'Corte': true }, { 'Enfesto': 2, 'Corte': 3 }) === 'cortando');
  ok('12e. "Retirada de fios" acende Retirando fio',
     leitura({ 'Corte': true, 'Retirada de fios': true }, { 'Corte': 3, 'Retirada de fios': 9 }) === 'fios');
  ok('12f. "Ensaque" acende Ensacado',
     leitura({ 'Corte': true, 'Ensaque': true }, { 'Corte': 3, 'Ensaque': 10 }) === 'ensacado');
  ok('12g. "Estoque" acende Estoque',
     leitura({ 'Corte': true, 'Estoque': true }, { 'Corte': 3, 'Estoque': 11 }) === 'estoque');
  // As duas costuras dividem a MESMA etapa do checklist; quem as separa e a
  // caixa "Recebido em Sao Carlos", do mesmo jeito que nos campos do fluxo.
  ok('12h. Costura sem passar por Sao Carlos: Costurando | Descalvado',
     leitura({ 'Corte': true, 'Costura': true }, { 'Corte': 3, 'Costura': 5 }) === 'costurando');
  ok('12i. a costura de Sao Carlos marcada: Costurando | Sao Carlos',
     leitura({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
             { 'Corte': 3, 'Recebido em São Carlos': 4, 'Costura CM.LISA | São Carlos': 5 }) === 'costurando-sc');
  // E o inverso, que e o caso que a medicao achou na OS 0507: a peca voltou
  // para ser costurada AQUI, e a caixa de recebimento nao pode mandar nela.
  ok('12i-b. recebida em SC, mas costurando em Descalvado: manda a ETAPA',
     leitura({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura': true },
             { 'Corte': 3, 'Recebido em São Carlos': 4, 'Costura': 5 }) === 'costurando');
  // Vale a marcada por ULTIMO, e nao a que esta mais adiante na lista: desmarcar
  // e voltar atras tem de levar o status junto.
  ok('12j. vale a etapa marcada por ULTIMO, mesmo sendo anterior no fluxo',
     leitura({ 'Corte': true, 'Costura': true }, { 'Costura': 5, 'Corte': 9 }) === 'cortando');
  // OS antiga, de antes de etapasSeq existir: desempata pela ordem REAL da
  // producao, que nao e a ordem do seletor (nele "Ensacado" vem antes das
  // costuras; no chao, depois).
  ok('12k. sem etapasSeq, vale a mais adiantada na ordem real da producao',
     leitura({ 'Corte': true, 'Costura': true, 'Ensaque': true }, {}) === 'ensacado',
     leitura({ 'Corte': true, 'Costura': true, 'Ensaque': true }, {}));

  /* O ENSAQUE TEM DUAS UNIDADES desde 17/09/2026, pela MESMA linha das duas
     costuras: o ensaque daqui e o fim do corte em Descalvado, e "Recebido em
     Sao Carlos" e o saco na prateleira DE LA. Antes os dois acendiam um
     "Ensacado" so, e quem lia a lista nao sabia de que lado estava o pano —
     a divisao existia nos campos do estoque de corte, e nao no status. */
  ok('12t. o ensaque daqui acende Ensacado | Descalvado',
     leitura({ 'Ensaque': true }, { 'Ensaque': 3 }) === 'ensacado',
     leitura({ 'Ensaque': true }, { 'Ensaque': 3 }));
  ok('12u. e "Recebido em Sao Carlos" acende Ensacado | Sao Carlos',
     leitura({ 'Corte': true, 'Recebido em São Carlos': true },
             { 'Corte': 3, 'Recebido em São Carlos': 4 }) === 'ensacado-sc',
     leitura({ 'Corte': true, 'Recebido em São Carlos': true },
             { 'Corte': 3, 'Recebido em São Carlos': 4 }));
  /* O LOOKAHEAD NEGATIVO no `re` do ensaque daqui, olhado de perto: sem ele, a
     etapa de Sao Carlos casaria com os DOIS status, e o desempate cairia na
     ordem da tabela em vez do lugar onde o pano esta. E o teste acima que
     quebraria — mas quebraria sem dizer POR QUE, entao a regra fica escrita. */
  const reDaqui = ctxDe('admin', 'a@b', true, []).api.STATUS_OS.find(x => x.k === 'ensacado').re;
  ok('12v. o `re` do ensaque daqui ignora a etapa de Sao Carlos',
     reDaqui.test('Ensaque') && !reDaqui.test('Recebido em São Carlos')
     && !reDaqui.test('Ensaque São Carlos'),
     String(reDaqui));
  /* E o ensaque daqui marcado DEPOIS traz a OS de volta: e a regra de sempre,
     vale a etapa marcada por ultimo. */
  ok('12x. e o ensaque daqui, marcado depois, traz a OS de volta',
     leitura({ 'Recebido em São Carlos': true, 'Ensaque': true },
             { 'Recebido em São Carlos': 4, 'Ensaque': 9 }) === 'ensacado',
     leitura({ 'Recebido em São Carlos': true, 'Ensaque': true },
             { 'Recebido em São Carlos': 4, 'Ensaque': 9 }));

  /* CHEGAR NAO E LIMPAR (17/09/2026, Junior: "Estoque com fio, derivado das OS
     que sao preenchidas o check box Recebido em Descalvado. Esse volume migra
     para Retirando fio quando essa check box e preenchida").

     A peca volta de Sao Carlos costurada, com os fios soltos, e FICA parada
     esperando a mesa de limpeza. Ate aqui "Recebido em Descalvado" acendia
     direto o "Retirando fio": a OS que tinha acabado de descer do caminhao
     aparecia como se ja estivesse sendo limpa, e o volume parado se somava ao
     volume em trabalho. */
  ok('12y. "Recebido em Descalvado" acende Estoque com fio',
     leitura({ 'Corte': true, 'Recebido em Descalvado': true },
             { 'Corte': 3, 'Recebido em Descalvado': 8 }) === 'estoque-fio',
     leitura({ 'Corte': true, 'Recebido em Descalvado': true },
             { 'Corte': 3, 'Recebido em Descalvado': 8 }));
  ok('12z. e a caixa da RETIRADA e que move para Retirando fio',
     leitura({ 'Recebido em Descalvado': true, 'Retirada de fios': true },
             { 'Recebido em Descalvado': 8, 'Retirada de fios': 9 }) === 'fios',
     leitura({ 'Recebido em Descalvado': true, 'Retirada de fios': true },
             { 'Recebido em Descalvado': 8, 'Retirada de fios': 9 }));
  /* O `re` da retirada nao pode voltar a pegar a chegada: era
     /fios|recebido em descalvado/i, e e essa volta que este teste barra. */
  const reFios = ctxDe('admin', 'a@b', true, []).api.STATUS_OS.find(x => x.k === 'fios').re;
  ok('12z-b. o `re` da retirada ignora a caixa de chegada',
     reFios.test('Retirada de fios') && !reFios.test('Recebido em Descalvado'),
     String(reFios));

  // A ORDEM DO SELETOR e a que o Junior escreveu, e nao a do fluxo: "Ensacado"
  // vem antes das costuras, e os dois de fora da fila (Parado e Cancelado, este
  // desde 17/09/2026) vem antes de "Estoque". E o que a pessoa le no seletor,
  // entao esta escrita aqui por inteiro — mudar a lista sem querer tem de
  // derrubar o teste.
  ok('12s. a fila do seletor esta na ordem pedida',
     ctxDe('admin', 'a@b', true, []).api.STATUS_OS.map(x => x.rotulo).join(' / ') ===
     ['Não iniciado', 'Preparando matéria-prima', 'Enfestando', 'Cortando',
      'Ensacado | Descalvado', 'Ensacado | São Carlos',
      'Estoque em trânsito | Desc x São Carlos', 'Estoque em trânsito | São Carlos X Desc.',
      'Costurando | Descalvado', 'Costurando | São Carlos',
      'Estoque com fio | Descalvado', 'Estoque com fio | São Carlos', 'Retirando fio',
      'Parado', 'Cancelado', 'Estoque'].join(' / '),
     ctxDe('admin', 'a@b', true, []).api.STATUS_OS.map(x => x.rotulo).join(' / '));

  console.log('');
  console.log('-- o carimbo a mao vale ATE a proxima etapa ser marcada --');
  /* Ele existe para adiantar o que a folha ainda nao sabe: entre comecar a
     enfestar e marcar "Enfesto" ha um dia inteiro. Marcada a etapa seguinte, a
     folha passou a saber mais do que o carimbo — e carimbo velho por cima dela
     e o estado em que a lista diz uma coisa e a OS diz outra. */
  const comCarimbo = (statusOS, quandoMs, check, seq) => {
    const o = osCheck(FLUXO, check, seq);
    o.statusOS = statusOS;
    o.statusOSPor = 'enfesto.corte@diverse.local';
    o.statusOSEm = new Date(quandoMs).toISOString();
    return ctxDe('admin', 'admin@diverse.local', true, [o]).api._statusOS(o);
  };
  ok('12l. carimbo mais novo que a ultima etapa: vale o carimbo',
     comCarimbo('enfestando', 5000, { 'Corte': true }, { 'Corte': 1000 }) === 'enfestando');
  ok('12m. etapa marcada DEPOIS do carimbo: o checklist retoma',
     comCarimbo('enfestando', 1000, { 'Corte': true }, { 'Corte': 5000 }) === 'cortando');
  ok('12n. vale tambem para "Parado": etapa nova quer dizer que a OS voltou a andar',
     comCarimbo('parado', 1000, { 'Corte': true }, { 'Corte': 5000 }) === 'cortando');
  ok('12o. e para "Finalizado"',
     comCarimbo('estoque', 1000, { 'Corte': true }, { 'Corte': 5000 }) === 'cortando');
  ok('12p. sem etapa marcada nenhuma, o carimbo vale sozinho',
     comCarimbo('parado', 1000, {}, {}) === 'parado');
  // As OS gravadas antes de 15/09/2026 tem a chave 'andamento', que saiu da
  // lista. Elas nao podem virar "nao iniciado": caem no checklist, que sabe
  // mais do que a chave velha sabia.
  ok('12q. a chave antiga "andamento" cai no checklist, e nao em "nao iniciado"',
     comCarimbo('andamento', 9999, { 'Corte': true }, { 'Corte': 1000 }) === 'cortando',
     comCarimbo('andamento', 9999, { 'Corte': true }, { 'Corte': 1000 }));
  ok('12r. e uma "andamento" sem checklist nenhum le "nao iniciado"',
     comCarimbo('andamento', 9999, {}, {}) === 'nao-iniciado');

  console.log('');
  console.log('-- o que aparece na coluna ACOES --');
  const osParado = { id: 'a1', os: '1234', statusOS: 'parado',
                     statusOSPor: 'enfesto.corte@diverse.local', statusOSEm: '2026-08-26T13:40:00.000Z' };
  let cel = ctxDe('admin', 'admin@diverse.local', true, [osParado]).api._statusCelulaOS(osParado);
  // O seletor oferece a FILA INTEIRA: as oito etapas mais "nao iniciado" e os
  // dois de fora da fila (parado, finalizado). O numero sai da propria tabela,
  // e nao de um 11 escrito aqui: status novo na fila nao pode derrubar o teste,
  // so o seletor que deixasse de oferecer o que a tabela tem.
  ok('20. quem muda ve um seletor com a fila inteira',
     /^<select/.test(cel)
     && (cel.match(/<option/g) || []).length === ctxDe('admin', 'a@b', true, []).api.STATUS_OS.length, cel);
  ok('20b. e o rotulo da lista vai ABREVIADO, que e o que cabe na coluna',
     /Costurando \| SC</.test(cel) && !/Costurando \| São Carlos</.test(cel), cel);
  ok('21. com o estado gravado ja escolhido',
     /value="parado"[^>]* selected/.test(cel), cel);
  ok('22. e a dica diz quem mexeu por ultimo',
     /enfesto\.corte/.test(cel) && /Parado/.test(cel), cel);
  cel = ctxDe('usuario', 'costura@diverse.local', true, [osParado]).api._statusCelulaOS(osParado);
  ok('23. quem so olha ve etiqueta, sem seletor',
     /^<span class="os-status ro"/.test(cel) && !/<select/.test(cel), cel);
  /* O ICONE E UM PINGO REDONDO NA COR DO STATUS (17/09/2026). Era um emoji, e
     o emoji escolhia a forma junto com a cor -- losango, quadrado e circulo em
     tres azuis parecidos. Aqui se prova o que mudou: a etiqueta traz o pingo
     pintado com a cor que a TABELA da para 'parado', e dentro do seletor, onde
     so cabe texto, vai o mesmo circulo em caractere, com a cor na opcao. */
  const tabela = ctxDe('admin', 'a@b', true, []).api.STATUS_OS;
  const corParado = tabela.find(x => x.k === 'parado').cor;
  ok('24. a etiqueta traz o pingo redondo na cor do status',
     cel.includes('class="st-pingo"') && cel.includes('background:' + corParado), cel);
  const celSel = ctxDe('admin', 'admin@diverse.local', true, [osParado]).api._statusCelulaOS(osParado);
  ok('24b. e o seletor leva o mesmo circulo, com uma cor por opcao',
     (celSel.match(/●/g) || []).length === tabela.length
     && celSel.includes('style="color:' + corParado + ';"'), celSel);
  ok('24c. e as cores da tabela nao se repetem',
     new Set(tabela.map(x => x.cor)).size === tabela.length,
     tabela.map(x => x.k + '=' + x.cor).join(' '));

  console.log('');
  console.log('-- a data de finalizacao (coluna Data, segunda linha) --');
  t = ctxDe('admin', 'admin@diverse.local', true, [{ id: 'a1', os: '1234', data: '2026-03-10' }]);
  const os2 = t.ctx.STATE.ordens[0];
  await t.api.mudarStatusOS('a1', 'enfestando');
  ok('33. enfestando nao carimba data de finalizacao', !('finalizadaEm' in os2), JSON.stringify(os2));
  await t.api.mudarStatusOS('a1', 'ensacado');
  ok('34. ENSACADO carimba o dia — e o fim da producao (15/09/2026)',
     typeof os2.finalizadaEm === 'string' && !isNaN(new Date(os2.finalizadaEm)), JSON.stringify(os2));
  const primeira = os2.finalizadaEm;
  // Um respiro: os dois carimbos no MESMO milissegundo dariam a mesma string, e
  // o teste acusaria como "nao recarimbou" algo que recarimbou.
  await new Promise(r => setTimeout(r, 5));
  await t.api.mudarStatusOS('a1', 'costurando');
  /* A DATA FICA. Com o fim no ensaque, sair dali e o caminho NORMAL — a OS
     ensacada segue para a costura da outra unidade —, e apagar a data faria
     toda OS perder o dia em que foi feita no minuto seguinte. */
  ok('35. seguir para a costura NAO apaga a data: ela e um fato, nao um estado',
     os2.finalizadaEm === primeira, JSON.stringify(os2));
  await t.api.mudarStatusOS('a1', 'nao-iniciado');
  ok('35b. so voltar ao comeco limpa: OS nao iniciada nao tem dia de termino',
     !('finalizadaEm' in os2), JSON.stringify(os2));
  await t.api.mudarStatusOS('a1', 'ensacado');
  ok('36. ensacar de novo carimba o dia NOVO',
     typeof os2.finalizadaEm === 'string' && os2.finalizadaEm !== primeira,
     os2.finalizadaEm + ' vs ' + primeira);

  const A = ctxDe('admin', 'admin@diverse.local', true, []).api;
  let cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10' });
  ok('37. OS sem status mostra so a data em que foi feita',
     /10\/03\/2026/.test(cel2) && !/data-fim/.test(cel2), cel2);
  cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10', statusOS: 'estoque',
                                finalizadaEm: '2026-08-26T13:40:00.000Z' });
  ok('38. finalizada mostra as duas: a de cima feita, a de baixo finalizada',
     /10\/03\/2026/.test(cel2) && /data-fim/.test(cel2) && /26\/08\/2026/.test(cel2), cel2);
  cel2 = A._dataCelulaListaOS({ os: '1', data: '2026-03-10', statusOS: 'ensacado',
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
  const finalizada = { os: '1', data: '2026-03-10', statusOS: 'estoque',
                       finalizadaEm: instante.toISOString() };
  ok('41. a data de finalizacao leva a hora junto',
     A._dataHoraFinalizacaoOS(finalizada) === '26/08/2026 16:26',
     A._dataHoraFinalizacaoOS(finalizada));
  ok('42. e ela aparece assim na coluna Data da lista',
     /26\/08\/2026 16:26/.test(A._dataCelulaListaOS(finalizada)),
     A._dataCelulaListaOS(finalizada));
  ok('43. OS que nao terminou nao tem dia nem hora',
     A._dataHoraFinalizacaoOS({ os: '1', data: '2026-03-10', statusOS: 'enfestando',
                                statusOSEm: instante.toISOString() }) === '');
  ok('44. data gravada que nao e data nao vira hora inventada',
     A._dataHoraFinalizacaoOS({ os: '1', statusOS: 'estoque', finalizadaEm: 'nao e data' }) === '');
  /* A DICA FALA DE CORTE, E NAO DE OS (15/09/2026). Junior: "essa data de
     finalizacao e para informar que a operacao de corte foi finalizada". A OS
     ensacada segue para a costura — chamar aquilo de "OS finalizada" fazia a
     folha mentir para quem a pegava na expedicao. */
  ok('45. a dica separa a hora REAL da hora do carimbo em lote',
     A._tituloFinalizacaoOS(finalizada) === 'Dia e hora em que o corte foi finalizado (OS ensacada)'
     && A._tituloFinalizacaoOS({ statusOS: 'estoque', statusOSEm: instante.toISOString() })
        === 'Dia e hora em que a OS foi marcada como Ensacado — o fim do corte',
     A._tituloFinalizacaoOS(finalizada));
  ok('45e. e nenhuma das dicas diz que a OS inteira terminou',
     ![A._tituloFinalizacaoOS(finalizada),
       A._tituloFinalizacaoOS({ statusOS: 'estoque', statusOSEm: instante.toISOString() })]
       .some(t => /OS foi finalizada|OS terminou/.test(t)));

  // O CORTE QUE TERMINOU PELA FOLHA. Com o status nascendo do checklist, marcar
  // o ensaque na folha fecha o corte sem ninguem carimbar nada — e ai nao ha
  // `finalizadaEm` nem `statusOSEm` para a coluna Data ler. A data sai de
  // `etapasSeq`, que e o instante em que a caixa foi marcada: e a hora real, e
  // nao uma reconstrucao.
  const terminouNaFolha = {
    os: '1', data: '2026-03-10',
    etapas: ['Corte', 'Ensaque'],
    progresso: { etapasCheck: { 'Corte': true, 'Ensaque': true },
                 etapasSeq: { 'Corte': 1000, 'Ensaque': instante.getTime() } }
  };
  ok('45a. terminada pelo checklist: o status e Ensacado, sem carimbo nenhum',
     A._statusOS(terminouNaFolha) === 'ensacado', A._statusOS(terminouNaFolha));
  ok('45b. e a data de fim e a hora em que a caixa Ensaque foi marcada',
     A._dataHoraFinalizacaoOS(terminouNaFolha) === '26/08/2026 16:26',
     A._dataHoraFinalizacaoOS(terminouNaFolha));
  ok('45c. a dica diz que a hora veio da folha, e nao de um carimbo',
     A._tituloFinalizacaoOS(terminouNaFolha)
       === 'Dia e hora em que a caixa do Ensaque foi marcada no checklist da folha',
     A._tituloFinalizacaoOS(terminouNaFolha));
  // Desmarcar a caixa tira a OS do fim: a data some junto, pelo mesmo motivo
  // que apagar o carimbo apaga — OS que voltou a andar nao terminou.
  const voltouAAndar = JSON.parse(JSON.stringify(terminouNaFolha));
  voltouAAndar.progresso.etapasCheck = { 'Corte': true };
  voltouAAndar.progresso.etapasSeq = { 'Corte': instante.getTime() + 1 };
  ok('45d. desmarcar a caixa Ensaque tira a data de fim junto',
     A._dataHoraFinalizacaoOS(voltouAAndar) === '' && A._statusOS(voltouAAndar) === 'cortando',
     A._statusOS(voltouAAndar) + ' / ' + A._dataHoraFinalizacaoOS(voltouAAndar));


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
  await e.api.mudarStatusOS('e1', 'enfestando');
  ok('46. "enfestando" baixa o pano sozinho', situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'estoque');
  ok('47. finalizado tambem baixa (quem pulou o enfesto ja gastou o pano)',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'parado');
  ok('48. parado idem: parou DEPOIS de comecar', situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  e = comMov('enfestando');
  e.ctx.STATE.estoqueMov.forEach(m => { if (m.origem === 'os') m.status = 'consumido'; });
  await e.api.mudarStatusOS('e1', 'nao-iniciado');
  ok('49. voltar para "nao iniciado" estorna: a OS nao gastou pano nenhum',
     situacao(e.ctx) === 'reservado+reservado', situacao(e.ctx));
  e = comMov();
  await e.api.mudarStatusOS('e1', 'enfestando');
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
  ok('52. OS enfestando salva de novo: o pano nasce ja BAIXADO',
     await salvando('enfestando') === 'consumido', await salvando('enfestando'));
  ok('53. e finalizada tambem — corrigir a folha nao desfaz a baixa',
     await salvando('estoque') === 'consumido', await salvando('estoque'));

  /* O PANO NAO VOLTA PARA A PRATELEIRA NO MEIO DO CAMINHO (Junior, 27/08/2026):
     "se as OS mudam para status parado ou voltam para em andamento, isso nao faz
     os tecidos reservados voltarem para o estoque reservado". E o que a fabrica
     ve: o rolo foi cortado no enfesto; a OS parar depois disso nao remonta o
     rolo. So "nao iniciado" — a OS que nao comecou — devolve a reserva. */
  e = comMov();
  await e.api.mudarStatusOS('e1', 'enfestando');
  await e.api.mudarStatusOS('e1', 'parado');
  ok('54. enfestando -> parado: o pano continua baixado',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  await e.api.mudarStatusOS('e1', 'enfestando');
  ok('55. e voltando a andar tambem — nada volta para reservado',
     situacao(e.ctx) === 'consumido+consumido', situacao(e.ctx));
  await e.api.mudarStatusOS('e1', 'estoque');
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
  await p.api.mudarStatusOS('at', 'ensacado');
  const pa = () => p.ctx.STATE.ordens[1];
  ok('68. finalizar a ativa finaliza a conjugada', pa().statusOS === 'ensacado',
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
  await p.api.mudarStatusOS('pa', 'estoque');
  ok('72. a PASSIVA nao arrasta a ativa', p.ctx.STATE.ordens[0].statusOS === undefined,
     JSON.stringify(p.ctx.STATE.ordens[0]));

  // TODOS os estados propagam (Junior, 28/08/2026: "a os conjugada deve seguir
  // todas as alteracoes de status da os ativa"). Nao ha estado em que uma
  // esteja e a outra nao: e o mesmo enfesto na mesa.
  p = parDe();
  await p.api.mudarStatusOS('at', 'enfestando');
  ok('73. "enfestando" propaga tambem', pa().statusOS === 'enfestando', JSON.stringify(pa()));
  p = parDe();
  await p.api.mudarStatusOS('at', 'parado');
  ok('73b. e "parado" idem', pa().statusOS === 'parado', JSON.stringify(pa()));

  // Ja finalizada nao e recarimbada: a data dela e o dia em que ela terminou.
  p = parDe('estoque');
  p.ctx.STATE.ordens[1].finalizadaEm = '2026-01-01T10:00:00.000Z';
  await p.api.mudarStatusOS('at', 'estoque');
  ok('74. conjugada ja finalizada mantem a data dela',
     pa().finalizadaEm === '2026-01-01T10:00:00.000Z', pa().finalizadaEm);

  // OS sozinha continua sozinha, e OS cuja irma sumiu nao quebra nada.
  p = parDe();
  await p.api.mudarStatusOS('so', 'estoque');
  ok('75. OS sem conjugada segue seu caminho', p.ctx.STATE.ordens[2].statusOS === 'estoque');
  const orfa = ctxDe('admin', 'admin@diverse.local', true,
                     [{ id: 'x', os: '0499', conjugadaId: 'sumiu' }]);
  await orfa.api.mudarStatusOS('x', 'estoque');
  ok('76. conjugada excluida depois: finaliza a ativa e nao reclama',
     orfa.ctx.STATE.ordens[0].statusOS === 'estoque' && orfa.ctx.salvou === 1);
  ok('77. _conjugadasQueSeguemStatus devolve LISTA (uma ativa pode puxar mais de uma)',
     Array.isArray(orfa.api._conjugadasQueSeguemStatus(orfa.ctx.STATE.ordens[0], 'estoque')));

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
  await d.api.mudarStatusOS('at', 'estoque');
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
  await d.api.mudarStatusOS('at', 'estoque');
  await d.api.mudarStatusOS('at', 'enfestando');
  ok('80. finalizado -> enfestando: a conjugada vai para enfestando tambem',
     dp().statusOS === 'enfestando' && !('finalizadaEm' in dp()), JSON.stringify(dp()));

  /* A conjugada e ALINHADA a ativa, venha ela de onde vier. Antes so o par
     finalizado->desfazer acompanhava, e uma conjugada em estado proprio ficava
     para tras — meio enfesto num estado, meio noutro. */
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'estoque' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'parado' }
  ]);
  await d.api.mudarStatusOS('at', 'enfestando');
  ok('81. conjugada em estado proprio e alinhada a ativa',
     d.ctx.STATE.ordens[1].statusOS === 'enfestando', JSON.stringify(d.ctx.STATE.ordens[1]));

  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'enfestando' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'estoque' }
  ]);
  await d.api.mudarStatusOS('at', 'parado');
  ok('82. enfestando -> parado leva a conjugada junto, e apaga a finalizacao dela',
     d.ctx.STATE.ordens[1].statusOS === 'parado'
     && !('finalizadaEm' in d.ctx.STATE.ordens[1]), JSON.stringify(d.ctx.STATE.ordens[1]));

  // O unico caso em que ela NAO e tocada: ja estar no estado pedido. Recarimbar
  // reescreveria a data de finalizacao dela — que e o dia em que ELA terminou.
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'enfestando' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'parado', statusOSEm: 'ontem' }
  ]);
  await d.api.mudarStatusOS('at', 'parado');
  ok('83. conjugada ja no estado pedido nao e recarimbada',
     d.ctx.STATE.ordens[1].statusOSEm === 'ontem', JSON.stringify(d.ctx.STATE.ordens[1]));

  /* CARIMBAR O MESMO STATUS DE NOVO CONSERTA O GRUPO (18/09/2026, Junior: "a
     alteracao de status de OS conjugada deve ser idempotente").

     Antes, escolher o status que a OS ja tinha saia na primeira linha da
     funcao, e um grupo desencontrado nao tinha gesto que o juntasse: o clique
     que deveria consertar era exatamente o que nao fazia nada. Foi o caso da
     0557 (manda) com a 0554 (segue) — a amarra foi feita DEPOIS do carimbo da
     mestre, entao a seguidora nunca recebeu aquele status.

     Idempotente aqui quer dizer as duas coisas: repetir CONVERGE (quem esta
     atrasado sobe) e repetir NAO ACUMULA (quem ja chegou nao e recarimbado,
     nao grava e nao mexe no estoque). */
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0557', statusOS: 'ensacado-sc', statusOSEm: 'ontem' },
    { id: 'pa', os: '0554', conjugadaStatusPaiId: 'at', statusOS: 'enfestando' }
  ]);
  await d.api.mudarStatusOS('at', 'ensacado-sc');
  ok('83b. repetir o status da que manda alinha a seguidora atrasada',
     d.ctx.STATE.ordens[1].statusOS === 'ensacado-sc', JSON.stringify(d.ctx.STATE.ordens[1]));
  ok('83c. e a data de quem ja estava no alvo NAO e reescrita',
     d.ctx.STATE.ordens[0].statusOSEm === 'ontem', JSON.stringify(d.ctx.STATE.ordens[0]));
  ok('83d. isso vale uma gravacao — o conserto precisa chegar ao servidor',
     d.ctx.salvou === 1, String(d.ctx.salvou));
  ok('83e. e o aviso diz que quem foi alinhada foi a conjugada, nao a clicada',
     d.ctx.toasts.some(t => /0554/.test(t) && /alinhada/i.test(t)),
     JSON.stringify(d.ctx.toasts));

  // Grupo inteiro no alvo: repetir continua sendo clique sem efeito nenhum.
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0557', statusOS: 'ensacado-sc', statusOSEm: 'ontem' },
    { id: 'pa', os: '0554', conjugadaStatusPaiId: 'at', statusOS: 'ensacado-sc', statusOSEm: 'ontem' }
  ]);
  await d.api.mudarStatusOS('at', 'ensacado-sc');
  ok('83f. grupo ja alinhado: repetir nao grava nada',
     d.ctx.salvou === 0 && d.ctx.STATE.ordens[1].statusOSEm === 'ontem',
     String(d.ctx.salvou) + ' ' + JSON.stringify(d.ctx.STATE.ordens[1]));

  // A mesma convergencia na amarra da GRADE, que e o outro caminho.
  d = ctxDe('admin', 'admin@diverse.local', true, [
    { id: 'at', os: '0498', conjugadaId: 'pa', statusOS: 'cortando' },
    { id: 'pa', os: '0497', conjugadaPaiId: 'at', statusOS: 'enfestando' }
  ]);
  await d.api.mudarStatusOS('at', 'cortando');
  ok('83g. na amarra da grade tambem: repetir alinha a que ficou para tras',
     d.ctx.STATE.ordens[1].statusOS === 'cortando', JSON.stringify(d.ctx.STATE.ordens[1]));

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
  await p.api.mudarStatusOS('at', 'enfestando');
  ok('93. as OS conjugadas a mao seguem a ativa',
     p.ctx.STATE.ordens[1].statusOS === 'enfestando' && p.ctx.STATE.ordens[2].statusOS === 'enfestando',
     JSON.stringify(p.ctx.STATE.ordens.slice(1, 3)));
  ok('94. e sao duas ou mais, nao so uma', p.api._conjugadasManuaisDaOS(p.ctx.STATE.ordens[0]).length === 2);
  ok('95. numa gravacao so', p.ctx.salvou === 1, String(p.ctx.salvou));
  ok('96. a OS que nao foi conjugada nao e tocada',
     p.ctx.STATE.ordens[3].statusOS === undefined);

  // Finalizar e desfazer: o grupo inteiro vai e volta.
  p = maoDe();
  await p.api.mudarStatusOS('at', 'ensacado');
  ok('97. finalizar leva todas, com a data',
     p.ctx.STATE.ordens.slice(1, 3).every(o => o.statusOS === 'ensacado' && !!o.finalizadaEm));
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
  p.ctx.STATE.ordens[1].statusOS = 'estoque';
  p.ctx.STATE.ordens[1].finalizadaEm = '2026-09-02T10:00:00.000Z';
  await p.api.mudarStatusOS('at', 'estoque');
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
  await p.api.mudarStatusOS('at', 'enfestando');
  ok('101. a conjugada DA CONJUGADA tambem vai (a mao puxa a da grade)',
     p.ctx.STATE.ordens[2].statusOS === 'enfestando', JSON.stringify(p.ctx.STATE.ordens[2]));

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
    { id: 'b', os: '0500', conjugadaStatusPaiId: 'at', statusOS: 'enfestando' },
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
     tc.ctx.STATE.ordens[1].statusOS === 'enfestando');

  /* O ALINHAR E OPCIONAL, e por isso: carimbar sozinho reescreveria a data de
     finalizacao de quem terminou em outro dia. Em branco, ninguem e tocado
     hoje; marcado, as escolhidas entram no estado da ativa agora. */
  tc = telaDe([
    { id: 'at', os: '0435', statusOS: 'enfestando' },
    { id: 'b', os: '0500' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['b'];
  await tc.api.salvarConjugarOS();
  ok('112. sem alinhar, a OS conjugada nao muda de estado hoje',
     tc.ctx.STATE.ordens[1].statusOS === undefined, JSON.stringify(tc.ctx.STATE.ordens[1]));

  tc = telaDe([
    { id: 'at', os: '0435', statusOS: 'enfestando' },
    { id: 'b', os: '0500' }
  ]);
  tc.api.conjugarNaOS('at');
  tc.ctx.marcadas = ['b'];
  tc.ctx.alinhar = true;
  await tc.api.salvarConjugarOS();
  ok('113. com alinhar, ela entra no estado da ativa agora',
     tc.ctx.STATE.ordens[1].statusOS === 'enfestando', JSON.stringify(tc.ctx.STATE.ordens[1]));

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
      ${constante('STATUS_FIM')}
      // O status le o checklist antes do carimbo: sem estas, _statusOS nao roda.
      ${recorte('function osEtapaMarcada', 'a etapa marcada no checklist')}
      ${src.match(/^const ETAPA_SC_RE = .+$/m)[0]}
      ${src.match(/^const _osRecebidaSC = .+$/m)[0]}
      ${src.match(/^const COSTURA_SC_RE = .+$/m)[0]}
      ${src.match(/^const _osCosturaEmSC = .+$/m)[0]}
      ${recorte('function _marcasDoStatus', 'as marcas de um status')}
      ${recorte('function _statusDoChecklistOS', 'o status que o checklist diz')}
      ${recorte('function _ultimaMarcacaoChecklist', 'a ultima etapa marcada')}
      ${recorte('function _statusOS', 'a leitura do status')}
      ${constante('STATUS_PONTO')}
      ${recorte('function _filtroStatusListaOS', 'o filtro por status')}
      return { _filtroStatusListaOS };
    `)(ctx);
    return { ctx, api, sel };
  };
  const osDoFiltro = [
    { id: '1', os: '0483' },                          // nao iniciado
    { id: '2', os: '0484', statusOS: 'parado' },
    { id: '3', os: '0485', statusOS: 'estoque' },
    { id: '4', os: '0486', statusOS: 'estoque' }
  ];
  let f = comSelect('', osDoFiltro);
  ok('27. sem escolha, o filtro nao corta nada', f.api._filtroStatusListaOS(osDoFiltro) === '');
  // "Todos" mais SÓ os estados que alguma OS da lista tem: aqui, nao iniciado,
  // parado e finalizado. Os outros oito da fila nao aparecem — eles nao existem
  // nesta lista, e oferecer oito linhas "(0)" e procurar no meio do que nao ha.
  ok('28. as opcoes sao "todos" mais os estados que EXISTEM na lista',
     (f.sel.innerHTML.match(/<option/g) || []).length === 4, f.sel.innerHTML);
  ok('28b. status sem nenhuma OS fica de fora',
     !/Enfestando/.test(f.sel.innerHTML) && !/Cortando/.test(f.sel.innerHTML), f.sel.innerHTML);
  ok('29. cada opcao ja diz quantas OS tem naquele estado',
     /Todos os status \(4\)/.test(f.sel.innerHTML)
     && /Estoque \(2\)/.test(f.sel.innerHTML)
     && /Parado \(1\)/.test(f.sel.innerHTML)
     && /Não iniciado \(1\)/.test(f.sel.innerHTML), f.sel.innerHTML);
  f = comSelect('estoque', osDoFiltro);
  ok('30. o escolhido volta como chave e continua marcado',
     f.api._filtroStatusListaOS(osDoFiltro) === 'estoque'
     && /value="estoque"[^>]* selected/.test(f.sel.innerHTML), f.sel.innerHTML);

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
