/* Rode com:  node testes/papel-cabecalho.js

   AS OPÇÕES DE EDIÇÃO QUE APARECIAM TARDE.

   Quem escreve no programa é só o admin: "+ Nova OS", o botão da barra lateral,
   editar/duplicar/excluir — tudo `admin-only`, tudo escondido enquanto o papel do
   usuário não é conhecido. Esconder é o padrão certo. O problema era QUANDO a
   resposta chegava: a pergunta de 50 bytes "qual é o meu papel?" era a última da
   fila, depois da sondagem do servidor, do download do blob de 1,8 MB, das
   compras, do catálogo de SKUs e do loadState. O cabeçalho aparecia pronto, em
   modo consulta, e as opções de edição pingavam na tela segundos depois.

   Duas defesas, e é o que este teste guarda:
     • o papel é LEMBRADO desta máquina, por USUÁRIO — a tela nasce certa;
     • falha de rede não rebaixa o admin a espectador.

   Lembrar é seguro porque o papel guardado só decide o que a tela MOSTRA: toda
   escrita passa por exigirAdmin/exigirEdicao, que olham o papel de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = [
  cortaLinha('const PAPEL_CACHE_CHAVE'),
  corta('function _lembrarPapel'),
  corta('function aplicarPapelLembrado'),
  corta('async function carregarPapel')
].join('\n');

// Um ambiente por cenário: localStorage de verdade (em memória), o usuário da
// sessão e o que o servidor responde à consulta do papel.
function ambiente({ guardado = {}, usuario = { id: 'u1' }, resposta = { data: { role: 'admin' } } } = {}) {
  const reg = { guardado, telaAplicada: 0 };
  const fn = new Function('REG', 'USUARIO', 'RESPOSTA', `
    const localStorage = {
      getItem: k => (k in REG.guardado ? REG.guardado[k] : null),
      setItem: (k, v) => { REG.guardado[k] = String(v); },
      removeItem: k => { delete REG.guardado[k]; }
    };
    const console = { warn: () => {} };
    let currentUser = USUARIO;
    let currentRole = null;
    const supa = { from: () => ({ select: () => ({ eq: () => ({ maybeSingle: async () => RESPOSTA }) }) }) };
    // A tela: só interessa QUANTAS vezes foi repintada e com que papel.
    function aplicarPermissoesUI() { REG.telaAplicada++; REG.ultimoPapelNaTela = currentRole; }
    ${motor}
    return {
      aplicarPapelLembrado, carregarPapel, _lembrarPapel,
      papel: () => currentRole,
      forcarPapel: (p) => { currentRole = p; },
      semUsuario: () => { currentUser = null; }
    };
  `);
  return { api: fn(reg, usuario, resposta), reg };
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

const CHAVE = 'papelUsuario';

/* ---------- 1. a tela nasce com o papel da última vez ---------- */

let { api, reg } = ambiente({ guardado: { [CHAVE]: JSON.stringify({ uid: 'u1', role: 'admin' }) } });
eq('papel lembrado do mesmo usuário é aplicado na hora', api.aplicarPapelLembrado(), true);
eq('e o papel vale como admin antes de qualquer resposta do servidor', api.papel(), 'admin');
eq('a tela foi pintada uma vez', reg.telaAplicada, 1);

// Primeira abertura nesta máquina: nada lembrado, então segue o padrão (consulta).
({ api, reg } = ambiente({ guardado: {} }));
eq('sem nada lembrado, não pinta nada', api.aplicarPapelLembrado(), false);
eq('e o papel continua indefinido (a tela esconde por padrão)', api.papel(), null);

// OUTRA conta na mesma máquina não herda o papel da anterior — este é o caso em
// que lembrar poderia mostrar edição a quem não edita.
({ api } = ambiente({
  guardado: { [CHAVE]: JSON.stringify({ uid: 'u1', role: 'admin' }) },
  usuario: { id: 'u2' }
}));
eq('papel de outro usuário não é aproveitado', api.aplicarPapelLembrado(), false);
eq('e o papel fica indefinido para a conta nova', api.papel(), null);

// Lixo no armazenamento não pode derrubar a abertura do programa.
({ api } = ambiente({ guardado: { [CHAVE]: 'isto não é json' } }));
eq('memória corrompida não quebra nem pinta', api.aplicarPapelLembrado(), false);

// Sessão sem usuário: nada a vestir.
({ api } = ambiente({ guardado: { [CHAVE]: JSON.stringify({ uid: 'u1', role: 'admin' }) } }));
api.semUsuario();
eq('sem usuário na sessão não pinta', api.aplicarPapelLembrado(), false);

/* ---------- 2. a resposta do servidor manda, e é lembrada ---------- */

(async () => {
  ({ api, reg } = ambiente({ resposta: { data: { role: 'admin' } } }));
  await api.carregarPapel();
  eq('o papel do servidor é adotado', api.papel(), 'admin');
  eq('e fica lembrado para a próxima abertura',
    JSON.parse(reg.guardado[CHAVE]).role, 'admin');
  eq('o usuário do papel lembrado é o da sessão',
    JSON.parse(reg.guardado[CHAVE]).uid, 'u1');
  eq('a tela é repintada porque o papel mudou (de indefinido para admin)', reg.telaAplicada, 1);

  // Papel que não mudou não precisa repintar a tela.
  ({ api, reg } = ambiente({ resposta: { data: { role: 'admin' } } }));
  api.forcarPapel('admin');
  await api.carregarPapel();
  eq('papel igual ao que já estava não repinta', reg.telaAplicada, 0);

  // Sem linha em user_roles = usuário comum (é o padrão da casa).
  ({ api, reg } = ambiente({ resposta: { data: null } }));
  await api.carregarPapel();
  eq('sem papel cadastrado, vale usuario', api.papel(), 'usuario');
  eq('e usuario também é lembrado', JSON.parse(reg.guardado[CHAVE]).role, 'usuario');

  // TROCA de papel no servidor: o lembrado é corrigido e a tela acompanha.
  ({ api, reg } = ambiente({
    guardado: { [CHAVE]: JSON.stringify({ uid: 'u1', role: 'admin' }) },
    resposta: { data: { role: 'usuario' } }
  }));
  api.aplicarPapelLembrado();
  await api.carregarPapel();
  eq('papel rebaixado no servidor vence o lembrado', api.papel(), 'usuario');
  eq('o lembrado é corrigido', JSON.parse(reg.guardado[CHAVE]).role, 'usuario');
  eq('a tela foi repintada (uma vez pelo lembrado, outra pela correção)', reg.telaAplicada, 2);

  /* ---------- 3. falha de rede não rebaixa quem já era admin ---------- */

  ({ api, reg } = ambiente({
    guardado: { [CHAVE]: JSON.stringify({ uid: 'u1', role: 'admin' }) },
    resposta: { error: { message: 'sem conexão' } }
  }));
  api.aplicarPapelLembrado();
  await api.carregarPapel();
  eq('erro na consulta mantém o papel lembrado', api.papel(), 'admin');
  eq('e não apaga o que estava lembrado', JSON.parse(reg.guardado[CHAVE]).role, 'admin');

  // Sem nada lembrado e com erro, o padrão seguro é consulta.
  ({ api } = ambiente({ guardado: {}, resposta: { error: { message: 'sem conexão' } } }));
  await api.carregarPapel();
  eq('erro sem papel lembrado cai em usuario (esconder é o padrão)', api.papel(), 'usuario');

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
