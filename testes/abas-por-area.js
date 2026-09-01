/* Rode com:  node testes/abas-por-area.js

   CADA ÁREA TEM ENDEREÇO, E ABRE EM ABA PRÓPRIA.

   O programa é uma página só: trocar de área é esconder uma <section> e mostrar
   outra. Sem endereço, abrir uma segunda aba caía sempre no Início — não havia
   como ter a lista de OS numa aba e a expedição na outra, e voltar a uma delas
   custava refazer o caminho pelo menu, perdendo rolagem e filtro.

   O que este teste protege:

     · o item de menu é um LINK (<a href="#area">), não um <button>. É o que faz
       o Ctrl+clique e o "abrir em nova aba" existirem — um <button> não é link
       para o navegador, e nenhum desses gestos funcionava;
     · o href e o data-page dizem a MESMA área (um link para a área errada é
       pior do que link nenhum);
     · Ctrl/⌘/Shift/Alt-clique NÃO é interceptado: quem segura a tecla está
       pedindo outra aba, e a aba de agora tem de ficar onde está;
     · trocar de área escreve o endereço, e só para área que tem menu — a folha
       de OS depende de qual OS está na mão, e "#print-os" numa aba nova abriria
       uma folha sem OS;
     · endereço desconhecido cai no Início, em vez de tela nenhuma;
     · o programa ABRE pelo endereço (F5 volta para onde se estava), inclusive
       depois de entrar na conta. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const css = fs.readFileSync(path.join(__dirname, '..', 'styles.css'), 'utf8');

function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// -- o menu --
const itens = [...html.matchAll(/<a class="nav-btn[^"]*" href="#([^"]+)" data-page="([^"]+)"[^>]*>/g)];
ok('1. todo item de menu e um link <a href="#area">',
  itens.length >= 25 && !/<button class="nav-btn/.test(html), itens.length);
ok('2. o href e o data-page apontam para a MESMA area',
  itens.every(m => m[1] === m[2]), itens.filter(m => m[1] !== m[2]).map(m => m[0]));
ok('3. o link nao fica sublinhado (continua parecendo o menu de sempre)',
  /\.nav-btn \{[^}]*text-decoration:\s*none/.test(css));
ok('3b. e cada item diz como se abre outra aba — o gesto e invisivel sem isso',
  itens.length > 0 && (html.match(/title="Ctrl\+clique/g) || []).length === itens.length,
  (html.match(/title="Ctrl\+clique/g) || []).length);

// -- as funcoes de endereco, com um DOM de mentira --
const paginasComMenu = itens.map(m => m[2]);
const monta = (hash, paginaNaTela) => {
  const ctx = { hash, ida: [] };
  const doc = {
    querySelector: sel => {
      let m = /^\.nav-btn\[data-page="([^"]+)"\]$/.exec(sel);
      if (m) return paginasComMenu.includes(m[1]) ? {} : null;
      if (sel === 'section.page:not(.hidden)') return { dataset: { page: paginaNaTela } };
      return null;
    }
  };
  const f = new Function('ctx', 'document', 'location', 'goto', `
    ${recorte('function _paginaTemEndereco', 'quem tem endereco')}
    ${recorte('function _paginaDoEndereco', 'a area do endereco')}
    ${recorte('function _paginaInicial', 'a area de abertura')}
    ${recorte('function _enderecoDaPagina', 'escrever o endereco')}
    return { _paginaTemEndereco, _paginaDoEndereco, _paginaInicial, _enderecoDaPagina };
  `);
  const location = { get hash() { return ctx.hash; }, set hash(v) { ctx.hash = v; ctx.escreveu = v; } };
  return Object.assign(f(ctx, doc, location, p => ctx.ida.push(p)), { ctx });
};

console.log('');
console.log('-- o endereco --');
let A = monta('#expedicao', 'home');
ok('4. area conhecida no endereco e a area de abertura', A._paginaInicial() === 'expedicao', A._paginaInicial());
A = monta('#coisa-que-nao-existe', 'home');
ok('5. endereco desconhecido cai no Inicio', A._paginaInicial() === 'home', A._paginaInicial());
A = monta('', 'home');
ok('6. sem endereco, Inicio', A._paginaInicial() === 'home', A._paginaInicial());
A = monta('#print-os', 'home');
ok('7. a folha de OS nao e endereco — ela depende de qual OS esta na mao',
  A._paginaInicial() === 'home', A._paginaInicial());

console.log('');
console.log('-- escrever o endereco ao trocar de area --');
A = monta('#home', 'home');
A._enderecoDaPagina('estoque');
ok('8. trocar de area escreve o endereco', A.ctx.hash === '#estoque', A.ctx.hash);
A = monta('#estoque', 'estoque');
A.ctx.escreveu = undefined;
A._enderecoDaPagina('estoque');
ok('9. ir para a area em que ja se esta nao mexe no historico',
  A.ctx.escreveu === undefined, A.ctx.escreveu);
A = monta('#estoque', 'estoque');
A.ctx.escreveu = undefined;
A._enderecoDaPagina('print-os');
ok('10. area sem menu nao escreve endereco nenhum',
  A.ctx.escreveu === undefined && A.ctx.hash === '#estoque', A.ctx.hash);

console.log('');
console.log('-- o que o navegador tem de continuar fazendo --');
// O .forEach do goto (que so tira a classe .active) tambem casa com um
// indexOf ingenuo: procura o que INSTALA o clique.
const handler = src.slice(src.indexOf(".nav-btn').forEach(b => b.addEventListener('click'"));
ok('11. Ctrl/Cmd/Shift/Alt-clique NAO e interceptado — e assim que se abre outra aba',
  /if \(e\.ctrlKey \|\| e\.metaKey \|\| e\.shiftKey \|\| e\.altKey\) return;/.test(handler.slice(0, 600)));
ok('12. o clique normal continua trocando de area sem recarregar',
  /e\.preventDefault\(\);\s*\n\s*goto\(b\.dataset\.page\);/.test(handler.slice(0, 600)));
ok('13. voltar e avancar do navegador andam entre as areas',
  /addEventListener\('hashchange'/.test(src));
ok('14. e o hashchange so age quando a area do endereco nao e a que ja esta na tela',
  /const atual = document\.querySelector\('section\.page:not\(\.hidden\)'\)\?\.dataset\?\.page;\s*\n\s*if \(page && page !== atual\) goto\(page\);/.test(src));

console.log('');
console.log('-- por onde o programa abre --');
ok('15. goto escreve o endereco da area', /_enderecoDaPagina\(page\);/.test(recorte('function goto(page)', 'o goto')));
ok('16. o programa abre pelo endereco, e nao sempre no Inicio',
  (src.match(/goto\(_paginaInicial\(\)\)/g) || []).length >= 3,
  (src.match(/goto\(_paginaInicial\(\)\)/g) || []).length);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
