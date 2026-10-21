/* Rode com:  node testes/diario-alteracoes.js

   O DIÁRIO DE ALTERAÇÕES (21/10/2026, Junior: "insira caixa de seleção para
   ver, editar e ser notificado para todos"). É dele que sai o Avisar das áreas
   que não tinham aviso próprio. O que este teste protege:
     · cada gravação de uma área vira UMA linha, com quem, quando, a tela e o
       nome da ação (o do Desfazer, ou o da ação liberada);
     · o que já tem aviso próprio (as OS, a expedição, o planejamento) e o
       próprio diário não entram;
     · a mesma ação gravando duas chaves é uma notícia só;
     · o Avisar das áreas novas nasce desligado, e a tela pode ter o seu;
     · toda linha da tela de acessos tem as três caixas, e as sem efeito (traço) são
       as combinadas com o Junior.
   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n}', i) + 2);
}
const constante = (nome) => {
  // Lista termina em "\n];", objeto em "\n};", e a de uma linha no primeiro ";".
  const m = src.match(new RegExp('^const ' + nome + ' = [^\\n]*;$', 'm'))
    || src.match(new RegExp('^const ' + nome + ' = \\[[\\s\\S]*?\\];|^const ' + nome + ' = \\{[\\s\\S]*?\\n\\};', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  let currentRole = ctx.papel || 'usuario';
  const currentUser = { email: ctx.login };
  let _desfazerRotulo = ctx.rotulo || '';
  const _paginaAtual = () => ctx.pagina || '';
  const _telaBase = p => p;
  let n = 0;
  const uid = () => 'd' + (++n);
  const saveState = async (k) => { ctx.gravou = (ctx.gravou || []).concat([k]); };
  const setTimeout = (fn) => { ctx.timer = fn; return 1; };
  const clearTimeout = () => {};
  const esc = s => String(s);
  const LOGINS_STATUS_OS = [], LOGINS_ESTOQUE_TECIDOS = [];
  ${constante('ACOES_POR_AREA')}
  ${constante('DIARIO_AREA_DA_CHAVE')}
  ${constante('DIARIO_AREA_DA_ACAO')}
  ${constante('DIARIO_SEM')}
  ${constante('DIARIO_AREAS_FORA')}
  const DIARIO_DIAS = 30, DIARIO_MAX = 400;
  let _diarioAcao = null, _diarioTimer = null;
  const _acessoTelasAbertas = new Set();
  ${['AREAS_ACESSO', 'ACESSO_PADRAO', 'VISAO_PAGINAS', 'VISAO_CONCEDIDA', 'VISAO_SEM_TELA', 'EDITAR_TRAVADO', 'AVISAR_TRAVADO',
     'AREAS_SO_VISAO', 'AREAS_SO_AVISO', 'AVISOS_DA_AREA', 'AVISO_PADRAO_DESLIGADO', 'TELAS_DA_AREA', '_TELA_DA_FOLHA'].map(constante).join('\n')}
  ${['function _obsQuemSou', 'function _obsNomeLogin', 'function _chaveLogin', 'function _acessoChaveConta', 'function _acessosTabela',
     'function _visaoTabela', 'function _avisosContaTabela', 'function contaRecebeAviso', 'function _areaDaPagina', 'function contaTemAcesso',
     'function contaVeArea', 'function _telasTabela', 'function _ajusteTela', 'function contaEditaTela', 'function contaVeTela',
     'function _telasAjustadas', 'function _acessosMenuHtml', 'function _areaDaAcao', 'function _diarioAcaoLiberada',
     'function _diarioLista', 'function _diarioAnotar', 'function _diarioDaGravacao'].map(f => recorte(f, f)).join('\n')}
  const _desfazerFrase = (k) => 'alteração em ' + k;
  return { _diarioDaGravacao, _diarioAcaoLiberada, contaRecebeAviso, _acessosMenuHtml,
           rotulo: r => { _desfazerRotulo = r; } };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const novo = (extra) => {
  const ctx = Object.assign({ login: 'nathaly@diverse.local', pagina: 'cad-grades', STATE: { meta: {}, diario: [] } }, extra || {});
  return { ctx, api: monta(ctx) };
};

console.log('-- o que vira linha do diário --');
let t = novo({ rotulo: 'cadastro da grade M-G-G1 | CM.LISA' });
t.api._diarioDaGravacao('grades', '[1]', '[1,2]');
let d = t.ctx.STATE.diario;
ok('1. gravar uma grade vira uma linha de Cadastros, com a tela e o nome do Desfazer',
   d.length === 1 && d[0].area === 'cadastros' && d[0].tela === 'cad-grades' && d[0].rotulo === 'cadastro da grade M-G-G1 | CM.LISA' && d[0].quem === 'nathaly@diverse.local', d);
ok('2. e grava o diário um pouco depois (junta as gravações da ação)', typeof t.ctx.timer === 'function' && !t.ctx.gravou);
t.api._diarioDaGravacao('desenhos', '[1]', '[1,2]');
ok('3. a mesma ação gravando outra chave é uma notícia só', t.ctx.STATE.diario.length === 1, t.ctx.STATE.diario);
t = novo();
t.api._diarioDaGravacao('ordens', '[1]', '[2]');
t.api._diarioDaGravacao('diario', '[1]', '[2]');
t.api._diarioDaGravacao('expedicaoCargas', '[1]', '[2]');
t.api._diarioDaGravacao('osPlanejadas', '[1]', '[2]');
ok('4. OS, diário, expedição e planejamento ficam de fora (têm aviso próprio)', t.ctx.STATE.diario.length === 0, t.ctx.STATE.diario);
t.api._diarioDaGravacao('grades', '[1]', '[1]');
ok('5. gravar sem mudar nada não é notícia', t.ctx.STATE.diario.length === 0);
t = novo({ pagina: 'config' });
t.api._diarioAcaoLiberada('conceder acesso');
t.api._diarioDaGravacao('meta', '{"a":1}', '{"a":2}');
ok('6. a concessão de acesso (meta) entra como Contas, pela ação liberada', t.ctx.STATE.diario.length === 1 && t.ctx.STATE.diario[0].area === 'contas' && t.ctx.STATE.diario[0].rotulo === 'conceder acesso', t.ctx.STATE.diario);
t = novo();
t.api._diarioDaGravacao('meta', '{"a":1}', '{"a":2}');
ok('7. meta sem ação conhecida não entra (configuração miúda)', t.ctx.STATE.diario.length === 0);
t = novo();
t.api._diarioAcaoLiberada('alocar OS na expedição');
t.api._diarioDaGravacao('meta', '{"a":1}', '{"a":2}');
ok('8. ação de Expedição também fica de fora', t.ctx.STATE.diario.length === 0);
t = novo({ STATE: { meta: {}, diario: [{ id: 'v', area: 'cadastros', rotulo: 'x', quem: 'a', em: '2020-01-01T00:00:00Z' }] } });
t.api._diarioDaGravacao('cores', '[1]', '[2]');
ok('9. linha com mais de 30 dias sai do diário', t.ctx.STATE.diario.length === 1 && t.ctx.STATE.diario[0].id !== 'v', t.ctx.STATE.diario);

console.log('\n-- quem recebe --');
t = novo();
const N = 'nathaly@diverse.local';
ok('10. Avisar das áreas novas nasce desligado', t.api.contaRecebeAviso('cadastros', N) === false && t.api.contaRecebeAviso('estoque-fases', N) === false);
ok('11. e o das que já avisavam continua ligado', t.api.contaRecebeAviso('expedicao', N) === true && t.api.contaRecebeAviso('os', N) === true);
t = novo({ STATE: { meta: { acessosAvisos: { nathaly: { cadastros: true } }, acessosTelas: { nathaly: { 'cad-equipe': { avisar: false } } } }, diario: [] } });
ok('12. ligado na área, a tela pode desligar só ela', t.api.contaRecebeAviso('cadastros', N, 'cad-grades') === true && t.api.contaRecebeAviso('cadastros', N, 'cad-equipe') === false);

console.log('\n-- a tela de acessos --');
t = novo({ papel: 'admin' });
const html = t.api._acessosMenuHtml({ id: 'u1', email: N });
const linhas = html.split('<div class="acesso-linha').slice(2);
const celulas = l => (l.match(/<input|class="acesso-nada"/g) || []).length;
ok('13. toda linha tem as três colunas (caixa ou traço)', linhas.length > 15 && linhas.every(l => celulas(l) === 3),
   linhas.filter(l => celulas(l) !== 3).map(l => l.slice(0, 120)));
const linha = nome => linhas.find(l => l.includes('<b>' + nome + '</b>')) || '';
const caixas = l => [...l.matchAll(/<input[^>]*>|<span class="acesso-nada"/g)].map(m => /acesso-nada/.test(m[0]) ? '#' : '.').join('');
ok('14. travadas: Editar e Avisar do Ranking e do Início', caixas(linha('Ranking de produção')) === '.##' && caixas(linha('Início')) === '.##');
ok('15. travada: Avisar do Relatório de produção (Ver e Editar livres)', caixas(linha('Relatório de produção')) === '..#');
ok('16. Contas de acesso: Ver e Editar travados (só do admin), Avisar livre', caixas(linha('Contas de acesso')) === '##.');
ok('17. as áreas sem tela ganham o Ver (status, fila, dados, @todos)', ['Status da OS', 'Fila de produção', 'Dados, backup e pastas', 'Mensagens para @todos']
   .every(n => caixas(linha(n)) === '...'));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
