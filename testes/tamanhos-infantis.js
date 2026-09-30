/* Rode com:  node testes/tamanhos-infantis.js

   A LINHA INFANTIL DE TAMANHOS (30/09/2026, Junior: "Insira na janela de
   cadastro de grade uma linha paralela de tamanhos, logo abaixo da linha de
   tamanhos P ao G3, com os tamanhos 2-4-6-8-10-12-14-16"). Decidido com ele:
   vale no programa inteiro, e uma grade usa UMA linha ou a outra.

   As chaves são t2…t16 e o rótulo é só o número. O que o teste guarda:
     · a grade adulta continua com as MESMAS 7 colunas de sempre (a folha monta
       as colunas por totaisPorTamanhoTomOS().keys);
     · a infantil sai com as 8 colunas 2 ao 16, e os totais batem;
     · a etiqueta escreve "2-4-6…" e não "T2-T4…";
     · os pacotes da etiqueta saem com o número do tamanho;
     · o nome da faixa de riscos: "2 ao 16".

   Recorta as funções do app.js de verdade. */
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
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a const ' + nome); process.exit(1); }
  return m[0];
};
const listaConst = (nome) => {
  const i = src.indexOf('const ' + nome + ' = [');
  const j = src.indexOf('];', i);
  if (i < 0 || j < 0) { console.error('nao achei a const ' + nome); process.exit(1); }
  return src.slice(i, j + 2);
};

const motor = [
  constante('ETIQUETA_CONTEUDO_REPOSICAO'),
  constante('ETIQUETA_CONTEUDO_REPOSICAO_BM'),
  constante('ETIQUETAS_REPOSICAO_POR_OS'),
  listaConst('ETIQUETA_COMPOSICAO_MOLETOM'),
  corta('function _composicaoPacoteMoletom'),
  corta('function totaisPorTamanhoTomOS'),
  listaConst('_PECAS_ETIQUETA_BMTRI'),
  listaConst('_PECAS_ETIQUETA_BMLISA'),
  corta('function _normNome'),
  corta('function _normSku'),
  corta('function _skuDaGrade'),
  corta('function _osEhBmTri'),
  corta('function _osEhBmLisa'),
  corta('function _pecasEtiquetaOS'),
  corta('function _coresDaPecaOS'),
  corta('function _tamanhosDaGradeExpandido'),
  corta('function _qtdePacoteEtiqueta'),
  corta('function dadosEtiquetaParaOS'),
  corta('function _riscoNomeTamanhos')
].join('\n');

function rodar(o, codigo) {
  return new Function('o', `
    const STATE = { grades: [], tecidos: [{ id: 't1' }], desenhos: [], cores: [] };
    const categoriaEfetivaTecido = () => 'malha';
    const corNomeCurto = (n) => String(n == null ? '' : n).trim();
    const tonsEfetivos = () => [1];
    const _osEhMoletom = () => false;
    const multiplicadorPecaOS = () => 2;
    const _skuDaOS = () => '';
    ${motor}
    return (${codigo});
  `)(o);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const INF = { t2: 1, t4: 1, t6: 1, t8: 1, t10: 1, t12: 1, t14: 1, t16: 1 };
const osInf = { id: 'i', os: '0700', grade: Object.assign({ total: 8, descricao: '2-4-6-8-10-12-14-16 | CM.LISA' }, INF),
                enfesto: { camadas: 10 }, fases: [], tecidos: [], progresso: {} };
const osAdu = { id: 'a', os: '0701', grade: { p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1, total: 7 },
                enfesto: { camadas: 10 }, fases: [], tecidos: [], progresso: {} };

/* ---------- as colunas da folha ---------- */
const tA = rodar(osAdu, 'totaisPorTamanhoTomOS(o)');
ok('adulta: as mesmas 7 colunas de sempre', JSON.stringify(tA.keys) === JSON.stringify(['p','m','g','gg','g1','g2','g3']), tA.keys);
ok('adulta: total geral = 7 × 10 camadas × 2', tA.totalGeral === 140, tA.totalGeral);
const tI = rodar(osInf, 'totaisPorTamanhoTomOS(o)');
ok('infantil: as 8 colunas 2 ao 16', JSON.stringify(tI.keys) === JSON.stringify(['t2','t4','t6','t8','t10','t12','t14','t16']), tI.keys);
ok('infantil: os 8 tamanhos entram na conta', tI.tamanhos.length === 8, tI.tamanhos);
ok('infantil: total geral = 8 × 10 × 2', tI.totalGeral === 160, tI.totalGeral);

/* ---------- a etiqueta ---------- */
const dI = rodar(osInf, 'dadosEtiquetaParaOS(o)');
ok('etiqueta infantil: tamanho escrito 2-4-6-8-10-12-14-16', dI.tam === '2-4-6-8-10-12-14-16', dI.tam);
const dA = rodar(osAdu, 'dadosEtiquetaParaOS(o)');
ok('etiqueta adulta: continua P-M-G-GG-G1-G2-G3', dA.tam === 'P-M-G-GG-G1-G2-G3', dA.tam);
ok('etiqueta infantil: pacotes com o número do tamanho (8 + reposição)',
   rodar(osInf, '_tamanhosDaGradeExpandido(o)').join(',') === '2,4,6,8,10,12,14,16',
   rodar(osInf, '_tamanhosDaGradeExpandido(o)'));

/* ---------- o nome da faixa ---------- */
ok('faixa infantil: "2 ao 16"', rodar(null, '_riscoNomeTamanhos(' + JSON.stringify(INF) + ')') === '2 ao 16',
   rodar(null, '_riscoNomeTamanhos(' + JSON.stringify(INF) + ')'));
ok('faixa adulta: continua "P ao G3"',
   rodar(null, '_riscoNomeTamanhos({p:1,m:1,g:1,gg:1,g1:1,g2:1,g3:1})') === 'P ao G3');

/* ---------- a linha no cadastro e no formulário ---------- */
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
ok('formulário da OS tem a linha infantil, logo depois do G3',
   /id="f-gr-g3"[\s\S]{0,600}grade-inputs-infantil[\s\S]*id="f-gr-t2"[\s\S]*id="f-gr-t16"/.test(html));
ok('janela da grade tem a linha 2 ao 16 abaixo da P ao G3',
   /\['p','m','g','gg','g1','g2','g3'\]\.map\(t => `\s*<div class="field"><label>\$\{t\.toUpperCase\(\)\}[\s\S]{0,400}grade-inputs-infantil/.test(src));
ok('grade com as duas linhas é recusada no cadastro', /Zere a outra linha antes de salvar/.test(src));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
