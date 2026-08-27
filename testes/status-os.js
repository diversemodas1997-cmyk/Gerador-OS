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
  ${recorte('function _recusarPorModoNuvem', 'a recusa do modo nuvem')}
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
  ${recorte('async function mudarStatusOS', 'a mudanca do status')}
  const exigirEstoqueTecidos = () => true;
  // aplicarBaixaEstoqueOS pergunta o consumo da OS ao cadastro; aqui ele vem
  // pronto pelo ctx, que e o que este teste tem a dizer sobre o assunto.
  const consumoAgregadoPorTecidoCor = () => ctx.consumo || [];
  const uid = () => 'm' + (++ctx.seq);
  const _estoqueRedesenharSeAberto = () => {};
  const renderEstoque = () => {};
  return { podeMudarStatusOS, _statusOS, _statusCelulaOS, mudarStatusOS, STATUS_OS,
           darBaixaMaterialOS, estornarBaixaMaterialOS, aplicarBaixaEstoqueOS,
           _dataFinalizacaoOS, _dataHoraFinalizacaoOS, _tituloFinalizacaoOS, _dataCelulaListaOS };
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
  console.log('-- o status mora SO na lista de OS Salvas --');
  // A celula e montada em um lugar so, dentro da coluna ACOES do renderListaOS.
  // Se alguem levar o status para a folha, a OE ou o painel de fases, o teste cai
  // e a conversa acontece antes.
  const usos = (src.match(/_statusCelulaOS\(/g) || []).length;
  ok('31. _statusCelulaOS e chamada uma vez so (a definicao + a chamada)',
     usos === 2, String(usos));
  const lista = src.slice(src.indexOf('function renderListaOS'));
  ok('32. e a chamada esta dentro da coluna col-actions da lista',
     /col-actions row-actions">\s*\$\{_statusCelulaOS\(o\)\}/.test(lista),
     lista.slice(lista.indexOf('col-actions'), lista.indexOf('col-actions') + 120));

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
