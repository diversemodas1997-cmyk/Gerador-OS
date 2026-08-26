/* Rode com:  node testes/avisos.js

   O MURAL DE AVISOS: o que os outros escreveram e carimbaram nas OS.

   O mural NÃO guarda evento nenhum — ele é lido das próprias OS (as notas da
   folha e o carimbo de status), e a marca de "já li" mora no navegador de cada
   um. Isso resolve o egresso (o blob desce inteiro a cada abertura), mas põe
   toda a responsabilidade em três contas que este teste protege:

     · o que entra na lista (janela de 30 dias, nota sem data fica fora);
     · o que conta como NÃO LIDO — o que chegou depois da última visita e não
       foi escrito por mim (ninguém precisa ser avisado de si mesmo);
     · a PRIMEIRA visita não pode nascer com trezentos avisos: um contador que
       nasce cheio ninguém zera, ignora.

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

// localStorage de mentira: um objeto. E o que o mural usa para lembrar da
// ultima visita, entao o teste precisa de um por cenario.
const monta = (ctx) => new Function('ctx', `
  const localStorage = {
    getItem: (k) => (k in ctx.ls ? ctx.ls[k] : null),
    setItem: (k, v) => { ctx.ls[k] = String(v); }
  };
  const currentUser = ctx.login ? { email: ctx.login } : null;
  const STATE = ctx.STATE;
  ${constante('AVISOS_DIAS')}
  ${constante('AVISOS_MAX')}
  ${constante('STATUS_OS')}
  ${recorte('function _obsNotas', 'as notas da OS')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _statusOS', 'a leitura do status')}
  ${recorte('function _avisosNascimento', 'o nascimento de um registro')}
  ${recorte('function _expCancelSet', 'as expedicoes canceladas')}
  ${recorte('function _avisosEventos', 'a lista de eventos')}
  ${constante('AVISOS_LOTE')}
  ${recorte('function _avisosAgrupar', 'o agrupamento do lote')}
  ${recorte('function _avisosChaveVistos', 'a chave da ultima visita')}
  ${recorte('function _avisosChaveLimpos', 'a chave da limpeza')}
  ${recorte('function _avisosLimpoAte', 'o corte da limpeza')}
  ${recorte('function _avisosMarcarVistos', 'a marca da visita')}
  ${recorte('function _avisosVistos', 'a leitura da ultima visita')}
  ${recorte('function _avisosNaoLidos', 'a conta do que e novo')}
  const limpar = (quando) => { ctx.ls[_avisosChaveLimpos()] = quando; };
  const desfazer = () => { delete ctx.ls[_avisosChaveLimpos()]; };
  return { _avisosEventos, _avisosNaoLidos, _avisosVistos, _avisosMarcarVistos, _avisosLimpoAte,
           _avisosAgrupar, limpar, desfazer };
`)(ctx);

const diasAtras = (d) => new Date(Date.now() - d * 24 * 60 * 60 * 1000).toISOString();
const ctxDe = (login, ordens, ls = {}, exp = {}) => {
  const ctx = { login, ls, STATE: { ordens,
    expedicaoCargas: exp.cargas || [], expedicaoJanelas: exp.janelas || [],
    expedicaoExcecoes: exp.excecoes || [] } };
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const ordens = () => ([
  { id: 'a', os: '0483', obsNotas: [{ login: 'costura@diverse.local', texto: 'faltou pano na 3a fase', em: diasAtras(1) }] },
  { id: 'b', os: '0484', statusOS: 'parado', statusOSPor: 'enfesto.corte@diverse.local', statusOSEm: diasAtras(2) },
  { id: 'c', os: '0485', obsNotas: [{ login: 'admin@diverse.local', texto: 'recado velho', em: diasAtras(60) }] },
  { id: 'd', os: '0486', obsNotas: [{ login: 'costura@diverse.local', texto: 'migrada, sem data', anterior: true }] },
  { id: 'e', os: '0487' },                                    // nada aconteceu nela
  { id: 'f', os: '0488', statusOS: 'nao-iniciado' }           // status que e a ausencia de status
]);

console.log('-- o que entra no mural --');
let t = ctxDe('admin@diverse.local', ordens());
let ev = t.api._avisosEventos();
ok('1. entra a observacao e o carimbo de status', ev.length === 2, JSON.stringify(ev.map(e => e.os)));
ok('2. mais recente em cima', ev[0].os === '0483' && ev[1].os === '0484', JSON.stringify(ev.map(e => e.os)));
ok('3. a observacao traz OS, quem, quando e o texto',
   ev[0].tipo === 'obs' && ev[0].osId === 'a' && ev[0].quem === 'costura@diverse.local'
   && /faltou pano/.test(ev[0].texto), JSON.stringify(ev[0]));
ok('4. o status traz o estado carimbado',
   ev[1].tipo === 'status' && ev[1].status === 'parado'
   && ev[1].quem === 'enfesto.corte@diverse.local', JSON.stringify(ev[1]));
ok('5. recado de 60 dias fica fora da janela', !ev.some(e => e.os === '0485'), JSON.stringify(ev.map(e => e.os)));
ok('6. nota migrada, sem data, nao vira aviso (nao da para dizer se e nova)',
   !ev.some(e => e.os === '0486'), JSON.stringify(ev.map(e => e.os)));
ok('7. "nao iniciado" nao e evento — e a ausencia de carimbo',
   !ev.some(e => e.os === '0488'), JSON.stringify(ev.map(e => e.os)));

console.log('');
console.log('-- o que e NAO LIDO --');
t = ctxDe('admin@diverse.local', ordens());
ok('8. primeira visita nasce lida (ninguem zera contador que nasce cheio)',
   t.api._avisosNaoLidos().length === 0, String(t.api._avisosNaoLidos().length));
ok('9. e a marca ficou guardada para a proxima vez',
   typeof t.ctx.ls['avisosVistos:admin@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));

// Visita antiga guardada: tudo o que veio depois e novidade.
t = ctxDe('admin@diverse.local', ordens(), { 'avisosVistos:admin@diverse.local': diasAtras(5) });
ok('10. depois de uma visita antiga, os dois eventos sao novos',
   t.api._avisosNaoLidos().length === 2, String(t.api._avisosNaoLidos().length));

// O proprio recado nunca e aviso.
t = ctxDe('costura@diverse.local', ordens(), { 'avisosVistos:costura@diverse.local': diasAtras(5) });
const naoLidos = t.api._avisosNaoLidos();
ok('11. o que EU escrevi nao me avisa',
   naoLidos.length === 1 && naoLidos[0].tipo === 'status',
   JSON.stringify(naoLidos.map(e => e.os + ':' + e.tipo)));

// A marca e por login: dois turnos no mesmo computador nao herdam o "ja li".
t = ctxDe('enfesto.corte@diverse.local', ordens(), { 'avisosVistos:costura@diverse.local': diasAtras(5) });
ok('12. a marca de um login nao vale para o outro',
   t.api._avisosNaoLidos().length === 0
   && typeof t.ctx.ls['avisosVistos:enfesto.corte@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));

// Marcou como visto agora: nada pendente ate acontecer coisa nova.
t = ctxDe('admin@diverse.local', ordens(), { 'avisosVistos:admin@diverse.local': diasAtras(5) });
t.api._avisosMarcarVistos(new Date().toISOString());
ok('13. abrir o mural zera o que estava pendente',
   t.api._avisosNaoLidos().length === 0, String(t.api._avisosNaoLidos().length));
const comNovo = ordens();
comNovo[0].obsNotas[0].editadoEm = new Date(Date.now() + 1000).toISOString();
t.ctx.STATE.ordens = comNovo;
ok('14. e a observacao corrigida depois disso volta a avisar',
   t.api._avisosNaoLidos().length === 1, String(t.api._avisosNaoLidos().length));

console.log('');
console.log('-- sem login --');
t = ctxDe('', ordens(), {});
ok('15. quem nao entrou na conta ve o mural, sem nada pendente',
   t.api._avisosEventos().length === 2 && t.api._avisosNaoLidos().length === 0,
   String(t.api._avisosNaoLidos().length));

console.log('');
console.log('-- a OS gerada e a OE montada --');
// Pedido de 26/08/2026: o programa avisa tambem quando nasce uma OS e quando
// uma expedicao e montada - era o que so se descobria abrindo a lista.
const msAtras = (h) => Date.now() - h * 3600 * 1000;
const nascidas = [
  { id: 'n1', os: '0504', criadoEm: new Date(msAtras(3)).toISOString(), criadoPor: 'admin@diverse.local',
    modeloNome: 'Camiseta Basica', grade: { total: 12 } },
  { id: 'n2', os: '0100', criadoEm: new Date(msAtras(24 * 40)).toISOString() }   // fora da janela
];
const janelas = [{ id: 'j1', nome: 'Matinal' }, { id: 'j2', nome: 'Tarde' }];
const cargas = [
  // A OE de 27/08: nasce com a primeira carga, e a segunda nao a faz nascer de novo.
  { id: 'c1', janelaId: 'j1', data: '2026-08-27', osId: 'n1', criadaEm: new Date(msAtras(5)).toISOString(), criadaPor: 'admin@diverse.local' },
  { id: 'c2', janelaId: 'j1', data: '2026-08-27', osId: 'n2', criadaEm: new Date(msAtras(2)).toISOString(), criadaPor: 'admin@diverse.local' },
  // Carga ANTIGA, sem criadaEm: o milissegundo sai do proprio id (uid()).
  { id: 'id_' + msAtras(8) + '_42', janelaId: 'j2', data: '2026-08-26', osId: 'n1' },
  // Expedicao CANCELADA: nao vira aviso.
  { id: 'c4', janelaId: 'j1', data: '2026-08-28', osId: 'n1', criadaEm: new Date(msAtras(1)).toISOString() }
];
const excecoes = [{ janelaId: 'j1', data: '2026-08-28', tipo: 'cancelada' }];
t = ctxDe('costura@diverse.local', nascidas, {}, { cargas, janelas, excecoes });
ev = t.api._avisosEventos();
const porTipo = (k) => ev.filter(e => e.tipo === k);
ok('16. a OS gerada vira aviso', porTipo('os').length === 1 && porTipo('os')[0].os === '0504',
   JSON.stringify(ev.map(e => e.tipo + ':' + (e.os || e.dataOe))));
ok('17. e leva o modelo e as pecas por camada',
   /Camiseta Basica/.test(porTipo('os')[0].texto) && /12/.test(porTipo('os')[0].texto),
   porTipo('os')[0].texto);
ok('18. OS criada ha 40 dias fica fora da janela', !porTipo('os').some(e => e.os === '0100'),
   JSON.stringify(porTipo('os').map(e => e.os)));
ok('19. a OE montada vira UM aviso, nao um por carga',
   porTipo('oe').filter(e => e.dataOe === '2026-08-27').length === 1,
   JSON.stringify(porTipo('oe').map(e => e.dataOe)));
const oe27 = porTipo('oe').find(e => e.dataOe === '2026-08-27');
ok('20. com a janela, quantas OS estao na carga e quem montou',
   oe27.janela === 'Matinal' && oe27.n === 2 && oe27.quem === 'admin@diverse.local', JSON.stringify(oe27));
ok('21. a OE nasce com a PRIMEIRA carga (5h atras), nao com a ultima',
   Math.abs(Date.parse(oe27.em) - msAtras(5)) < 2000, oe27.em);
ok('22. carga antiga sem criadaEm: o milissegundo sai do proprio id',
   porTipo('oe').some(e => e.dataOe === '2026-08-26'), JSON.stringify(porTipo('oe').map(e => e.dataOe)));
ok('23. expedicao CANCELADA nao vira aviso',
   !porTipo('oe').some(e => e.dataOe === '2026-08-28'), JSON.stringify(porTipo('oe').map(e => e.dataOe)));

console.log('');
console.log('-- acao em lote vira UMA linha --');
// O lote de 26/08 (200 OS finalizadas de uma vez pelo script) afogava o mural:
// 200 linhas iguais empurravam o recado da producao para fora das 100 linhas.
const mesmoMinuto = '2026-08-26T10:52:53.823Z';
const lote = [];
for (let i = 0; i < 12; i++) {
  lote.push({ id: 'L' + i, os: '02' + String(10 + i), statusOS: 'finalizado',
              statusOSPor: 'admin@diverse.local', statusOSEm: mesmoMinuto });
}
// Duas no mesmo minuto continuam sendo duas noticias.
lote.push({ id: 'p1', os: '0301', statusOS: 'parado', statusOSPor: 'enfesto.corte@diverse.local',
            statusOSEm: '2026-08-26T11:10:00.000Z' });
lote.push({ id: 'p2', os: '0302', statusOS: 'parado', statusOSPor: 'enfesto.corte@diverse.local',
            statusOSEm: '2026-08-26T11:10:10.000Z' });
const A = ctxDe('costura@diverse.local', [], {}).api;
const juntos = A._avisosAgrupar([
  ...lote.map(o => ({ tipo: 'status', osId: o.id, os: o.os, quem: o.statusOSPor,
                      em: o.statusOSEm, status: o.statusOS }))
]);
const doLote = juntos.find(e => e.tipo === 'status-lote');
ok('24. doze carimbos iguais no mesmo minuto viram uma linha so',
   !!doLote && doLote.n === 12, JSON.stringify(juntos.map(e => e.tipo + (e.n ? ':' + e.n : ''))));
ok('25. a linha do lote diz o estado e quem fez',
   doLote.status === 'finalizado' && doLote.quem === 'admin@diverse.local', JSON.stringify(doLote));
ok('26. e guarda os numeros das OS para mostrar os primeiros',
   Array.isArray(doLote.oss) && doLote.oss.length === 12, JSON.stringify(doLote.oss));
ok('27. duas no mesmo minuto NAO agrupam (ainda sao duas noticias)',
   juntos.filter(e => e.tipo === 'status' && e.status === 'parado').length === 2,
   JSON.stringify(juntos.map(e => e.tipo)));
ok('28. e o mural nao afoga: 14 carimbos viram 3 linhas',
   juntos.length === 3, String(juntos.length));

console.log('');
console.log('-- limpar a lista --');
// Limpar e marca de LEITURA, nao apagamento: o que sai e a linha do mural, e o
// recado segue na OS. Vale so para quem limpou, e so neste computador.
t = ctxDe('admin@diverse.local', ordens(), { 'avisosVistos:admin@diverse.local': diasAtras(5) });
ok('29. antes de limpar, os dois eventos estao na lista',
   t.api._avisosEventos().length === 2 && t.api._avisosNaoLidos().length === 2,
   String(t.api._avisosEventos().length));
t.api.limpar(new Date().toISOString());
ok('30. limpar esvazia a lista', t.api._avisosEventos().length === 0, String(t.api._avisosEventos().length));
ok('31. e zera o contador junto', t.api._avisosNaoLidos().length === 0, String(t.api._avisosNaoLidos().length));
ok('32. as OS nao foram tocadas — o recado segue la',
   t.ctx.STATE.ordens[0].obsNotas[0].texto === 'faltou pano na 3a fase'
   && t.ctx.STATE.ordens[1].statusOS === 'parado', JSON.stringify(t.ctx.STATE.ordens[1]));
// O que acontecer DEPOIS da limpeza volta a aparecer.
const depois = ordens();
depois.push({ id: 'g', os: '0490', statusOS: 'andamento', statusOSPor: 'costura@diverse.local',
              statusOSEm: new Date(Date.now() + 2000).toISOString() });
t.ctx.STATE.ordens = depois;
ok('33. o que acontece depois da limpeza aparece',
   t.api._avisosEventos().length === 1 && t.api._avisosNaoLidos().length === 1,
   String(t.api._avisosEventos().length));
t.api.desfazer();
ok('34. "ver os avisos de novo" traz a lista de volta',
   t.api._avisosEventos().length === 3, String(t.api._avisosEventos().length));
// A limpeza e de quem limpou: o outro login continua com os avisos dele.
t.api.limpar(new Date().toISOString());
const outro = ctxDe('costura@diverse.local', depois, t.ctx.ls);
ok('35. limpar a minha lista nao limpa a do outro turno',
   outro.api._avisosEventos().length === 3, String(outro.api._avisosEventos().length));

console.log('');
console.log('-- a lista acumulada mora atras do sino, junto do icone do perfil --');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
// O pedido de 26/08/2026: os avisos nao sao uma pagina do menu, sao uma lista
// acumulada que se expande no clique do sino — e o sino fica junto do icone do
// perfil, que e onde se olha para saber quem esta logado.
ok('36. o sino esta dentro do bloco do perfil logado',
   /auth-identity[\s\S]{0,900}?id="avisosSino"/.test(html)
   && html.indexOf('id="avisosSino"') < html.indexOf('</aside>'), 'sino fora do bloco do perfil');
ok('37. o contador fica no sino', /id="avisosSino"[\s\S]{0,400}?id="avisosBadge"/.test(html), 'badge solto');
ok('38. o painel existe e nasce escondido, fora da barra lateral',
   /class="avisos-painel hidden" id="avisosPainel"/.test(html)
   && html.indexOf('id="avisosPainel"') > html.indexOf('</aside>'), 'painel dentro da aside ou visivel');
ok('39. nao ha mais pagina de avisos no menu (a lista e o painel)',
   !/data-page="avisos"/.test(html), 'sobrou a pagina antiga');
ok('40. clicar no sino abre e fecha',
   /function toggleAvisos/.test(src) && /function abrirAvisos/.test(src) && /function fecharAvisos/.test(src),
   'faltou abrir/fechar');
ok('41. e abrir a OS pelo aviso fecha o painel antes',
   /fecharAvisos\(\); verOS\(/.test(src), 'o painel ficaria por cima da folha');
ok('42. o painel tem o botao de limpar',
   /class="avisos-limpar" onclick="limparAvisos\(\)"/.test(html), 'sem botao de limpar');
ok('43. e o painel vazio oferece o desfazer',
   /mostrarTodosAvisos\(\)/.test(src), 'sem "ver os avisos de novo"');

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
