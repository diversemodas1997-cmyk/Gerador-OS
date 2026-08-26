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
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('function _obsQuando', 'a data da nota')}
  ${recorte('function _chaveLogin', 'a chave do login')}
  ${recorte('function podeMudarStatusOS', 'quem muda o status')}
  ${recorte('function exigirStatusOS', 'a recusa do status')}
  ${recorte('function _recusarPorModoNuvem', 'a recusa do modo nuvem')}
  ${recorte('function _statusOS', 'a leitura do status')}
  ${recorte('function _statusCelulaOS', 'a celula do status')}
  ${recorte('async function mudarStatusOS', 'a mudanca do status')}
  return { podeMudarStatusOS, _statusOS, _statusCelulaOS, mudarStatusOS, STATUS_OS };
`)(ctx);

const ctxDe = (papel, login, servidorNoAr = true, ordens = []) => {
  const ctx = { papel, login, servidorNoAr, toasts: [], salvou: 0, redesenhou: 0,
                STATE: { ordens } };
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
