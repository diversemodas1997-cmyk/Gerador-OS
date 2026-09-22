/* Rode com:  node testes/handlers-existem.js

   TODO BOTAO CHAMA UMA FUNCAO QUE EXISTE.

   Este teste nasceu de um estrago real (22/09/2026). Ao retirar o
   recolher/estender do Ranking, o corte levou junto trinta linhas que moravam
   logo abaixo dele e nao tinham nada com o Ranking: `abrirListaPorStatus` e os
   dois campos PENDENTES que o `goto` consome. O programa continuou compilando
   — ninguem chama aquela funcao no codigo, ela e chamada de dentro de um
   `onclick=` escrito em texto —, e o quadro "Volume das OS por status" passou a
   nao abrir a lista ao ser clicado. Junior achou na tela o que o compilador nao
   tinha como achar.

   O que se verifica: TODO handler inline (onclick, onchange, oninput, onsubmit,
   onkeydown) do app.js e do index.html aponta para um nome que o app.js define
   — funcao declarada, const de seta ou window.<nome>. E um teste de texto, de
   proposito: e assim que o navegador vai procurar a funcao na hora do clique.

   Ele nao prova que o botao FAZ a coisa certa; prova que ele nao morre em
   silencio, que e o defeito que passa por qualquer revisao de codigo. */
const fs = require('fs');
const path = require('path');

const raiz = path.join(__dirname, '..');
const app = fs.readFileSync(path.join(raiz, 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');

// O que o app.js poe ao alcance de um onclick: funcao de topo, const de seta no
// topo do arquivo e tudo o que e pendurado no window.
const definidos = new Set();
for (const re of [
  /^(?:async )?function ([A-Za-z0-9_$]+)\(/gm,
  /^window\.([A-Za-z0-9_$]+) =/gm,
  /^(?:const|let|var) ([A-Za-z0-9_$]+) = (?:async )?\(/gm,
  /^(?:const|let|var) ([A-Za-z0-9_$]+) = (?:async )?function/gm
]) {
  for (const m of app.matchAll(re)) definidos.add(m[1]);
}

// O que o navegador ja tem, e por isso nao precisa estar no app.js.
const NATIVOS = new Set(['window', 'document', 'event', 'this', 'alert', 'confirm',
  'print', 'open', 'setTimeout', 'parseInt', 'parseFloat', 'Number', 'String',
  // Palavras de controle que aparecem quando o handler traz codigo inline
  // ("onkeydown=\"if(event.key==='Enter'){...}\"").
  'if', 'for', 'while', 'switch', 'return', 'try']);

const HANDLER = /on(?:click|change|input|submit|keydown|keyup|focus|blur)="\s*([A-Za-z0-9_$.]+)\s*\(/g;
const achados = new Map();   // nome -> [arquivos]
for (const [arquivo, txt] of [['app.js', app], ['index.html', html]]) {
  for (const m of txt.matchAll(HANDLER)) {
    const nome = m[1].split('.')[0];
    if (NATIVOS.has(nome)) continue;
    if (!achados.has(nome)) achados.set(nome, new Set());
    achados.get(nome).add(arquivo);
  }
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const orfaos = [...achados.keys()].filter(n => !definidos.has(n)).sort();
console.log(`-- ${achados.size} funcoes chamadas de handlers inline, ${definidos.size} definidas no app.js --`);
ok('1. todo handler inline aponta para uma funcao que existe',
   orfaos.length === 0,
   orfaos.map(n => n + ' (em ' + [...achados.get(n)].join(', ') + ')'));

/* Os atalhos que a fabrica usa todo dia entram nomeados, para a falta deles
   falhar com o nome na tela em vez de virar mais uma linha numa lista grande. */
const ESSENCIAIS = [
  'abrirListaPorStatus',   // o quadro "Volume das OS por status" -> lista de OS
  '_rankingAbrirGrupo',    // uma celula do Ranking -> as OS daquele cruzamento
  'goto', 'verOS', 'salvarOS', 'renderListaOS',
  'moverNaFila', 'definirPosicaoFila',       // a fila de producao
  'gerarOCdaCompra', 'imprimirOC',           // a ordem de compra
  'compraAdicionar', 'compraRemover'
];
ESSENCIAIS.forEach((n, i) => {
  ok((i + 2) + '. ' + n + ' esta definida e ao alcance do clique', definidos.has(n), n);
});

/* E os dois campos PENDENTES que o `goto` LE ao abrir a lista de OS. Ler uma
   variavel que ninguem declarou lanca ReferenceError e derruba a navegacao
   inteira para aquela tela — foi metade do estrago de 22/09. */
['_listaOsGrupoPendente', '_listaOsStatusPendente'].forEach((n, i) => {
  ok((ESSENCIAIS.length + 2 + i) + '. ' + n + ' continua declarado (o goto le este campo)',
     new RegExp('^(?:let|var|const) ' + n + '\\b', 'm').test(app), n);
});

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
