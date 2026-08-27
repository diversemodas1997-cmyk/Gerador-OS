/* Rode com:  node testes/mensagens.js

   O CAMPO DE MENSAGENS: o recado de todo mundo, num canal só.

   Pedido do Junior em 26/08/2026. Sem conversa privada e sem grupo, de
   propósito — recado de produção que só duas pessoas veem é o mesmo que recado
   no papel: some, e ninguém mais sabe que existiu.

   O que este teste protege:

     · o recado vai assinado por QUEM mandou (o RLS do servidor recusa gravar em
       nome de outro; aqui se garante que o app manda o autor certo);
     · o que EU mandei nunca conta como não lido;
     · a marca de leitura é por login e mora no navegador, como as outras;
     · a mensagem não cabe no blob dos dados — ela vive na tabela `mensagens`,
       e é isso que o app tem de consultar;
     · servidor sem a tabela (a nuvem, por exemplo) esconde o ícone em vez de
       piscar erro.

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

// O Supabase de mentira: guarda o que foi inserido e devolve o que mandarem.
const monta = (ctx) => new Function('ctx', `
  const toast = (m, t) => ctx.toasts.push((t || '') + ':' + m);
  const localStorage = {
    getItem: (k) => (k in ctx.ls ? ctx.ls[k] : null),
    setItem: (k, v) => { ctx.ls[k] = String(v); }
  };
  const document = { getElementById: () => ctx.campo, addEventListener: () => {} };
  const currentUser = ctx.user;
  let currentRole = ctx.papel || 'usuario';
  const _recusarPorModoNuvem = () => !!ctx.modoNuvem;
  // O select das MENSAGENS termina em .limit(); o das REACOES, em .in(). O mesmo
  // objeto responde aos dois, e o .in devolve lista vazia: reacao nao e assunto
  // deste teste, mas carregarMensagens passa por la.
  const supa = {
    from: (tabela) => ({
      select: () => ({
        order: () => ({ limit: async () => (tabela === 'mensagens' ? ctx.resposta : { data: [] }) }),
        in: async () => ({ data: ctx.reacoes || [] })
      }),
      insert: (linha) => { ctx.inserido = linha; return {
        select: () => ({ maybeSingle: async () => ({ data: { id: 'novo', ...linha, criado_em: new Date().toISOString() } }) }) }; },
      delete: () => ({ eq: async () => ({ error: null }) })
    }),
    rpc: async () => ({ data: ctx.perfis || [] })
  };
  let _mensagens = ctx.mensagens || [];
  let _msgIndisponivel = false, _msgContagem = null, _msgCanal = null;
  // O canal ganhou resposta (responde_a), polegar e mencao em 27/08/2026; o
  // que este teste precisa saber deles e so que existem.
  let _msgRespondendoA = null, _msgReacoes = [], _msgReacoesIndisponivel = false;
  let _msgPerfis = [], _msgMencao = null, _msgEditando = null;
  const STATE = { ordens: [] };
  function _msgBarraResposta() {}
  function _msgFecharMencao() {}
  ${constante('MSG_MAX')}
  ${recorte('function _msgQuemSou', 'quem sou nas mensagens')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('async function carregarMensagens', 'a leitura das mensagens')}
  ${recorte('async function enviarMensagem', 'o envio')}
  ${recorte('function _msgChaveVistas', 'a chave da leitura')}
  ${recorte('function _msgVistas', 'a marca de leitura')}
  ${recorte('function _msgMarcarVistas', 'a gravacao da marca')}
  ${recorte('function _msgNaoLidas', 'a conta do que nao li')}
  // O polegar e a lista de mencao tem os seus proprios caminhos e nao sao o
  // assunto deste teste (que e ler, mandar e contar o que nao foi lido);
  // carregarMensagens passa por eles, entao aqui eles existem e nao fazem nada.
  async function carregarReacoes() { ctx.pediuReacoes = (ctx.pediuReacoes || 0) + 1; }
  async function carregarPerfis() { ctx.pediuPerfis = (ctx.pediuPerfis || 0) + 1; }
  function mensagensAberto() { return false; }
  function renderMensagens() { ctx.desenhou = (ctx.desenhou || 0) + 1; }
  return {
    carregarMensagens, enviarMensagem, _msgNaoLidas, _msgVistas, _msgMarcarVistas,
    lista: () => _mensagens, indisponivel: () => _msgIndisponivel
  };
`)(ctx);

const ctxDe = (extra = {}) => {
  const ctx = Object.assign({
    user: { id: 'u1', email: 'costura@diverse.local' },
    toasts: [], ls: {}, campo: { value: '', focus() {} },
    mensagens: [], resposta: { data: [], error: null }
  }, extra);
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};
const agora = () => new Date().toISOString();
const antes = (min) => new Date(Date.now() - min * 60000).toISOString();

(async () => {
  console.log('-- ler o canal --');
  let t = ctxDe({ resposta: { data: [
    { id: 'm2', criado_em: antes(5), autor_id: 'u2', autor: 'corte@diverse.local', texto: 'segundo' },
    { id: 'm1', criado_em: antes(9), autor_id: 'u1', autor: 'costura@diverse.local', texto: 'primeiro' }
  ], error: null } });
  await t.api.carregarMensagens();
  ok('1. o canal chega em ordem de leitura (mais antigo primeiro)',
     t.api.lista().map(m => m.id).join(',') === 'm1,m2', t.api.lista().map(m => m.id).join(','));
  ok('2. e o servidor com a tabela nao marca indisponivel', t.api.indisponivel() === false);

  t = ctxDe({ resposta: { data: null, error: { message: 'relation "mensagens" does not exist' } } });
  await t.api.carregarMensagens();
  ok('3. servidor sem a tabela: o campo fica indisponivel, sem erro na cara',
     t.api.indisponivel() === true && t.api.lista().length === 0, '');

  console.log('');
  console.log('-- mandar recado --');
  t = ctxDe();
  t.ctx.campo.value = '  faltou pano na 3a fase  ';
  await t.api.enviarMensagem();
  ok('4. o recado vai assinado por quem mandou',
     t.ctx.inserido && t.ctx.inserido.autor_id === 'u1'
     && t.ctx.inserido.autor === 'costura@diverse.local', JSON.stringify(t.ctx.inserido));
  ok('5. sem espaco sobrando nas pontas', t.ctx.inserido.texto === 'faltou pano na 3a fase',
     JSON.stringify(t.ctx.inserido.texto));
  ok('6. e o campo esvazia para o proximo', t.ctx.campo.value === '', t.ctx.campo.value);
  ok('7. a mensagem entra na lista na hora', t.api.lista().length === 1, String(t.api.lista().length));

  t = ctxDe();
  t.ctx.campo.value = '   ';
  await t.api.enviarMensagem();
  ok('8. recado vazio nao vai', !t.ctx.inserido, JSON.stringify(t.ctx.inserido));

  t = ctxDe({ user: null });
  t.ctx.campo.value = 'oi';
  await t.api.enviarMensagem();
  ok('9. sem login, nao manda e explica', !t.ctx.inserido && /conta/.test(t.ctx.toasts.join(' ')),
     t.ctx.toasts.join(' | '));

  t = ctxDe({ modoNuvem: true });
  t.ctx.campo.value = 'oi';
  await t.api.enviarMensagem();
  ok('10. modo nuvem (servidor da fabrica fora do ar) nao manda', !t.ctx.inserido, '');

  t = ctxDe();
  t.ctx.campo.value = 'x'.repeat(3000);
  await t.api.enviarMensagem();
  ok('11. recado e recado: corta em 2000 caracteres', t.ctx.inserido.texto.length === 2000,
     String(t.ctx.inserido.texto.length));

  console.log('');
  console.log('-- o que ainda nao li --');
  const conversa = [
    { id: 'a', criado_em: antes(30), autor_id: 'u2', autor: 'corte@diverse.local', texto: 'a' },
    { id: 'b', criado_em: antes(2), autor_id: 'u2', autor: 'corte@diverse.local', texto: 'b' },
    { id: 'c', criado_em: antes(1), autor_id: 'u1', autor: 'costura@diverse.local', texto: 'meu' }
  ];
  t = ctxDe({ mensagens: conversa, ls: { 'mensagensVistas:costura@diverse.local': antes(10) } });
  ok('12. conta so o que chegou depois da minha ultima leitura',
     t.api._msgNaoLidas().map(m => m.id).join(',') === 'b', t.api._msgNaoLidas().map(m => m.id).join(','));
  ok('13. o que EU mandei nunca conta como nao lido',
     !t.api._msgNaoLidas().some(m => m.id === 'c'), '');
  t.api._msgMarcarVistas(agora());
  ok('14. ler zera a conta', t.api._msgNaoLidas().length === 0, String(t.api._msgNaoLidas().length));
  t = ctxDe({ mensagens: conversa, ls: {} });
  ok('15. primeira vez nesta conta nasce lida (nao comeca com 200 recados)',
     t.api._msgNaoLidas().length === 0, String(t.api._msgNaoLidas().length));
  ok('16. e a marca fica guardada por login',
     typeof t.ctx.ls['mensagensVistas:costura@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));
  t = ctxDe({ mensagens: conversa, ls: { 'mensagensVistas:outro@diverse.local': antes(10) } });
  ok('17. a marca de um login nao vale para o outro',
     t.api._msgNaoLidas().length === 0
     && typeof t.ctx.ls['mensagensVistas:costura@diverse.local'] === 'string', JSON.stringify(t.ctx.ls));

  console.log('');
  console.log('-- onde a mensagem mora --');
  ok('18. na tabela `mensagens`, e nao no blob que desce inteiro',
     /from\('mensagens'\)/.test(src) && !/STATE\.mensagens/.test(src), 'mensagem no blob');
  ok('19. com Realtime, para o recado chegar sem recarregar',
     /channel\('mensagens_all'\)/.test(src) && /table: 'mensagens'/.test(src), 'sem realtime');
  const sql = fs.readFileSync(path.join(__dirname, '..', 'sql', 'supabase-mensagens.sql'), 'utf8');
  ok('20. o script do servidor cria a tabela e liga o RLS',
     /create table if not exists mensagens/.test(sql) && /enable row level security/.test(sql), '');
  ok('21. escrever exige ser voce mesmo (autor_id = auth.uid())',
     /with check \(autor_id = auth\.uid\(\)\)/.test(sql), 'qualquer um assinaria por outro');
  ok('22. e apagar e so do proprio autor (ou do admin)',
     /for delete/.test(sql) && /autor_id = auth\.uid\(\)/.test(sql) && /role = 'admin'/.test(sql), '');

  console.log('');
  console.log('-- na tela --');
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  ok('23. o icone fica junto do perfil, com o proprio numero',
     /id="msgBotao"/.test(html) && /id="msgBadge"/.test(html)
     && html.indexOf('id="msgBotao"') < html.indexOf('</aside>'), 'icone fora do bloco do perfil');
  ok('24. o painel tem a lista e o campo de escrever',
     /id="msg-lista"/.test(html) && /id="msg-texto"/.test(html) && /enviarMensagem\(\)/.test(html), '');
  ok('25. Enter manda o recado', /_msgTecla\(event\)/.test(html) && /ev\.key === 'Enter'/.test(src), '');

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
