/* Rode com:  node testes/busca-desenho-os.js

   A BUSCA DO DESENHO TÉCNICO no formulário da OS.

   São mais de cem desenhos numa lista suspensa só — achar o certo ali era rolar
   a lista inteira lendo código por código. Pedido do Junior em 26/08/2026: um
   campo de busca no seletor de desenho da tela de produção de OS.

   O que este teste protege:

     · a busca varre tudo o que identifica o desenho (código, descrição, modelo,
       cor e SKU) e casa TODOS os termos digitados;
     · o desenho JÁ ESCOLHIDO nunca sai da lista — filtrar não pode desfazer em
       silêncio uma escolha que já estava na OS;
     · a escolha sobrevive ao recorte (o value do seletor não se perde);
     · a conta ao lado diz quantos sobraram de quantos.

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

const monta = (ctx) => new Function('ctx', `
  const esc = (s) => String(s == null ? '' : s).replace(/"/g, '&quot;');
  const STATE = ctx.STATE;
  const document = { getElementById: (id) => ctx.el[id] || null };
  ${recorte('function _textoBuscaDesenho', 'o texto que a busca varre')}
  ${recorte('function _rotuloDesenhoOS', 'o rotulo da opcao')}
  ${recorte('function _desenhoSkuProduto', 'o SKU do produto')}
  ${recorte('function _desenhoModelo', 'o modelo do desenho')}
  ${recorte('function _desenhoVariacao', 'a variacao do desenho')}
  ${recorte('function _desenhosAgrupados', 'o agrupamento por semelhanca')}
  ${recorte('function _rotuloGrupoDesenho', 'o rotulo do grupo')}
  ${recorte('function _rotuloDesenhoNoGrupo', 'o rotulo dentro do grupo')}
  ${recorte('function filtrarDesenhosOS', 'o filtro do desenho')}
  return { filtrarDesenhosOS, _textoBuscaDesenho };
`)(ctx);

const DESENHOS = [
  { id: 'd1', codigo: '008', desc: 'Camiseta gola careca', modeloId: 'm1', corPrincipalId: 'c1', skuLinha: 'CM.LISA-PRE' },
  { id: 'd2', codigo: '0023', desc: 'Tricolor manga longa', modeloId: 'm2', corPrincipalId: 'c1', corSecundariaId: 'c2', skuLinha: 'BM.TRI-PRE' },
  { id: 'd3', codigo: '0016', desc: 'Polo piquet', modeloId: 'm3', corPrincipalId: 'c3', skuLinha: 'PM.LISA-BEG' }
];
const MODELOS = [{ id: 'm1', nome: 'Camiseta Básica' }, { id: 'm2', nome: 'Blusa Moletom Tricolor' },
                 { id: 'm3', nome: 'Camiseta Polo' }];
const CORES = [{ id: 'c1', nome: 'Preto' }, { id: 'c2', nome: 'Mostarda' }, { id: 'c3', nome: 'Bege' }];

const ctxDe = (busca = '', escolhido = '') => {
  const ctx = {
    STATE: { desenhos: DESENHOS, modelos: MODELOS, cores: CORES },
    el: {
      'f-desenho': { value: escolhido, innerHTML: '' },
      'f-desenho-busca': { value: busca },
      'f-desenho-conta': { textContent: '', classList: { toggle: (c, v) => { ctx.contaFiltrando = v; } } }
    }
  };
  return { ctx, api: monta(ctx) };
};
const opcoes = (ctx) => (ctx.el['f-desenho'].innerHTML.match(/<option/g) || []).length - 1;  // fora o "— selecione —"

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

console.log('-- o que a busca varre --');
let t = ctxDe();
const casa = (d, termo) => {
  const alvo = t.api._textoBuscaDesenho(d);
  return termo.toLowerCase().split(/\s+/).filter(Boolean).every(x => alvo.includes(x));
};
ok('1. acha pelo codigo', casa(DESENHOS[0], '008'));
ok('2. acha pela descricao', casa(DESENHOS[2], 'piquet'));
ok('3. acha pelo modelo', casa(DESENHOS[1], 'moletom'));
ok('4. acha pela cor', casa(DESENHOS[1], 'mostarda') && !casa(DESENHOS[0], 'mostarda'));
ok('5. acha pelo SKU', casa(DESENHOS[2], 'pm.lisa'));
ok('6. dois termos valem juntos', casa(DESENHOS[1], 'moletom preto') && !casa(DESENHOS[0], 'moletom preto'));

console.log('');
console.log('-- a lista recortada --');
t = ctxDe('');
t.api.filtrarDesenhosOS();
ok('7. sem busca, a lista inteira', opcoes(t.ctx) === 3, String(opcoes(t.ctx)));
ok('8. e a conta diz quantos sao', /3 desenhos/.test(t.ctx.el['f-desenho-conta'].textContent),
   t.ctx.el['f-desenho-conta'].textContent);

t = ctxDe('polo');
t.api.filtrarDesenhosOS();
ok('9. com busca, so o que casa', opcoes(t.ctx) === 1 && /Polo piquet/.test(t.ctx.el['f-desenho'].innerHTML),
   t.ctx.el['f-desenho'].innerHTML);
ok('10. e a conta diz quantos de quantos',
   /1 de 3 desenhos/.test(t.ctx.el['f-desenho-conta'].textContent) && t.ctx.contaFiltrando === true,
   t.ctx.el['f-desenho-conta'].textContent);

t = ctxDe('veludo');
t.api.filtrarDesenhosOS();
ok('11. busca sem resultado deixa so o "selecione"', opcoes(t.ctx) === 0, String(opcoes(t.ctx)));

console.log('');
console.log('-- o desenho ja escolhido --');
t = ctxDe('polo', 'd1');            // escolhido é o 008, a busca é "polo"
t.api.filtrarDesenhosOS();
ok('12. o escolhido continua na lista, mesmo fora da busca',
   /value="d1"/.test(t.ctx.el['f-desenho'].innerHTML), t.ctx.el['f-desenho'].innerHTML);
ok('13. e continua marcado (a escolha da OS nao se perde)',
   /value="d1" selected/.test(t.ctx.el['f-desenho'].innerHTML), t.ctx.el['f-desenho'].innerHTML);
ok('14. junto com o que a busca achou', opcoes(t.ctx) === 2, String(opcoes(t.ctx)));

console.log('');
console.log('-- na tela --');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
ok('15. o campo de busca fica em cima da lista, no card do desenho tecnico',
   /id="f-desenho-busca"/.test(html)
   && html.indexOf('id="f-desenho-busca"') < html.indexOf('id="f-desenho"'), 'busca no lugar errado');
ok('16. digitar filtra na hora', /oninput="filtrarDesenhosOS\(\)"/.test(html), 'sem o oninput');
ok('17. abrir o formulario limpa a busca da OS anterior',
   /const buscaDes = document\.getElementById\('f-desenho-busca'\)/.test(src), 'a busca vazaria entre OS');

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
