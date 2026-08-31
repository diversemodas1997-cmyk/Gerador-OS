/* Rode com:  node testes/reacao-nomes.js

   QUEM MARCOU O RECADO — o nome, e não só o número.

   Pedido do Junior em 31/08/2026. O polegar já contava quantos tinham marcado,
   e o próprio comentário no app.js dizia que o tooltip servia para responder
   "quem já viu?" — mas ele mostrava só o número. Numa fábrica de dez pessoas,
   "3 pessoas marcaram" não diz se quem precisava ver, viu.

   O que este teste protege:

     · "você" vem primeiro e com esse nome — é como a pessoa se reconhece na
       lista, e vê-lo no meio dos outros por login é ler o próprio nome como se
       fosse de terceiro;
     · o resto sai em ordem alfabética, para a lista não mudar de ordem a cada
       render (a ordem de chegada das reações não é estável);
     · a concordância acompanha o número: "marcou" para um, "marcaram" para
       vários — errar isso é o tipo de coisa que ninguém reporta e todo mundo lê;
     · a lista para em 8 nomes e diz quantos sobraram: um tooltip com trinta
       nomes não informa, entulha;
     · o nome tem TRÊS fontes, e a falta de uma não pode apagar a pessoa da
       lista — senão o número diria 3 e os nomes seriam 2.

   Recorta as funções do app.js de verdade.  */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(assinatura) {
  const i = src.indexOf(assinatura);
  if (i < 0) { console.error('não achei ' + assinatura + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('não achei o fim de ' + assinatura); process.exit(1); }
  return src.slice(i, j + 2);
}

// O contexto de que as funções precisam: quem sou eu, a lista de perfis e os
// recados já carregados (que carregam o nome do autor junto).
const monta = (ctx) => new Function('ctx', `
  const _msgQuemSou = () => ctx.eu;
  const _msgPerfis  = ctx.perfis || [];
  const _mensagens  = ctx.mensagens || [];
  ${recorte('function _obsNomeLogin(login)')}
  ${recorte('function _msgNomeDe(userId)')}
  ${recorte('function _msgTextoReacao(quemReagiu)')}
  return { _msgNomeDe, _msgTextoReacao };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

const CTX = {
  eu: 'u-junior',
  perfis: [
    { user_id: 'u-junior',  login: 'junior'  },
    { user_id: 'u-nathaly', login: 'nathaly' },
    { user_id: 'u-ana',     login: 'ana'     }
  ],
  // O Carlos não está nos perfis, mas escreveu um recado — e todo recado leva o
  // nome do autor junto. É a segunda fonte do nome.
  mensagens: [{ id: 'm1', autor_id: 'u-carlos', autor: 'carlos@diverse.com.br' }]
};
const app = monta(CTX);
const r = (...ids) => ids.map(id => ({ mensagem_id: 'm1', user_id: id }));

// 1. As três fontes do nome, e o recuo quando não há nenhuma.
ok('1. eu sou "você"',            app._msgNomeDe('u-junior') === 'você');
ok('1b. perfil dá o login',       app._msgNomeDe('u-ana') === 'ana');
ok('1c. recado dá o nome',        app._msgNomeDe('u-carlos') === 'carlos',
   'esperava carlos, veio ' + app._msgNomeDe('u-carlos'));
ok('1d. desconhecido não some',   app._msgNomeDe('u-zzz') === 'alguém');
ok('1e. id vazio não quebra',     app._msgNomeDe('') === 'alguém');

// 2. Concordância: um marcou, vários marcaram.
ok('2. um: "marcou"',   app._msgTextoReacao(r('u-ana')) === 'ana marcou este recado',
   app._msgTextoReacao(r('u-ana')));
ok('2b. dois: "marcaram"',
   app._msgTextoReacao(r('u-ana', 'u-nathaly')) === 'ana e nathaly marcaram este recado',
   app._msgTextoReacao(r('u-ana', 'u-nathaly')));
ok('2c. três usa vírgula e "e"',
   app._msgTextoReacao(r('u-ana', 'u-nathaly', 'u-carlos')) === 'ana, carlos e nathaly marcaram este recado',
   app._msgTextoReacao(r('u-ana', 'u-nathaly', 'u-carlos')));

// 3. "você" na frente, mesmo chegando por último.
ok('3. "você" vem primeiro',
   app._msgTextoReacao(r('u-ana', 'u-junior')) === 'você e ana marcaram este recado',
   app._msgTextoReacao(r('u-ana', 'u-junior')));
ok('3b. só eu',
   app._msgTextoReacao(r('u-junior')) === 'você marcou este recado',
   app._msgTextoReacao(r('u-junior')));

// 4. Ordem alfabética não depende da ordem de chegada — senão a lista dança a
//    cada render e a mesma reação parece outra.
ok('4. ordem estável',
   app._msgTextoReacao(r('u-nathaly', 'u-ana')) === app._msgTextoReacao(r('u-ana', 'u-nathaly')));

// 5. O corte em 8, com o resto contado.
const muitos = ['u-junior'].concat(
  Array.from({ length: 12 }, (_, i) => 'u-x' + String(i).padStart(2, '0')));
const txt = app._msgTextoReacao(r(...muitos));
ok('5. começa por você',      /^você, /.test(txt), txt);
ok('5b. mostra 8 nomes',      txt.split(' e mais ')[0].split(/,\s|\se\s/).length === 8, txt);
ok('5c. conta os que sobram', / e mais 5 marcaram este recado$/.test(txt), txt);

// 6. Sem ninguém, frase nenhuma — o botão cai no texto de convite.
ok('6. lista vazia devolve vazio', app._msgTextoReacao([]) === '');

console.log(falhas ? `\n${falhas} FALHA(S)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
