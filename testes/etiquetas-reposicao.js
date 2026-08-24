/* Rode com:  node testes/etiquetas-reposicao.js

   AS ETIQUETAS DE VIÉS/REPOSIÇÃO/RIBANA.

   O conjunto de etiquetas de uma OS é: uma por PACOTE de tamanho (tamanhos ×
   tonalidades) mais o pacote de reposição. Em 24/08/2026 o pedido do Junior
   mudou três coisas nesse fim de conjunto:

     1. a etiqueta de reposição sai em DUAS vias iguais (uma colada por fora do
        saco, outra dentro);
     2. as duas trazem o resumo das tonalidades numa linha só, junto das demais
        ("TONS: 1 · 2") — antes a reposição era a única etiqueta do conjunto que
        não dizia tom nenhum;
     3. a via EXTRA não participa da contagem de pacotes: o LOTE conta pacotes
        (tamanhos + 1 de reposição), e ela sai SEM linha de lote.

   O ponto que este teste guarda é o 3: etiqueta e pacote deixaram de ser a
   mesma coisa. Numerar a via extra faria a expedição procurar na doca um pacote
   que não existe.

   Recorta _tamanhosDaGradeExpandido, dadosEtiquetaParaOS e gerarPdfEtiquetas do
   app.js de verdade. A jsPDF entra dublada — em vez de desenhar, ela ANOTA o
   texto de cada página, que é exatamente o que se quer conferir. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
// As constantes do conteudo e do numero de vias saem do app.js tambem: o teste
// nao pode ter a sua propria ideia de quantas etiquetas de reposicao existem.
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
  constante('ETIQUETAS_REPOSICAO_POR_OS'),
  listaConst('ETIQUETA_COMPOSICAO_MOLETOM'),
  corta('function _tamanhosDaGradeExpandido'),
  corta('function dadosEtiquetaParaOS'),
  corta('function gerarPdfEtiquetas')
].join('\n');

// jsPDF dublê: cada página vira um array com o texto das linhas desenhadas.
const jsPdfDuble = `
  function jsPDFDuble() {
    this.paginas = [[]];
    const cur = () => this.paginas[this.paginas.length - 1];
    this.addPage = () => { this.paginas.push([]); };
    this.setFont = () => {}; this.setFontSize = () => {};
    this.setLineWidth = () => {}; this.rect = () => {}; this.line = () => {};
    this.getTextWidth = (s) => String(s).length * 1.6;   // largura plausível em mm
    this.text = (t) => { cur().push(String(t)); };
    this.output = () => 'BLOB';
  }`;

// Roda o conjunto de etiquetas de uma OS e devolve { dados, paginas }.
function etiquetasDe(o, { tons = [], moletom = false } = {}) {
  const fn = new Function('o', 'tons', 'moletom', `
    ${jsPdfDuble}
    let capturado = null;
    const window = { jspdf: { jsPDF: function () { capturado = new jsPDFDuble(); return capturado; } } };
    // Dublês: o que se mede aqui é a montagem do conjunto, não o cadastro.
    // Um tecido cadastrado: e a categoria dele que diz se a OS e de moletom.
    const STATE = { grades: [], tecidos: [{ id: 't1' }], desenhos: [], cores: [] };
    const categoriaEfetivaTecido = () => (moletom ? 'moletom' : 'malha');
    const corNomeCurto = (n) => String(n == null ? '' : n).trim();
    const tonsEfetivos = () => tons;
    const _osEhMoletom = () => moletom;
    ${motor}
    const dados = dadosEtiquetaParaOS(o);
    gerarPdfEtiquetas(dados);
    return { dados, paginas: capturado.paginas };
  `);
  return fn(o, tons, moletom);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

// OS de camiseta, grade P-M-G (1 vaga cada). O tom vem do dublê.
const osBase = (extra) => Object.assign({
  id: 'os_1', os: '0501', griffeNome: 'Diverse',
  grade: { p: 1, m: 1, g: 1, total: 3 },
  enfesto: { camadas: 10 },
  fases: [], tecidos: [], progresso: {}
}, extra || {});

const temLote = (pag) => pag.some(l => /^LOTE:/.test(l));
const loteDe = (pag) => (pag.find(l => /^LOTE:/.test(l)) || '').replace('LOTE: ', '');
const linhaTons = (pag) => pag.find(l => /^(TOM|TONS):/.test(l)) || '';
const ehRep = (pag) => pag.some(l => l.includes('Viés/Reposição/Ribana'));

/* ---------- 1. um tom: 3 pacotes de tamanho + 1 de reposição ---------- */

let r = etiquetasDe(osBase(), { tons: [1] });
eq('1 tom: pacotes contados', r.dados.totalPacotes, 4);
eq('1 tom: etiquetas impressas (a via extra da reposição)', r.dados.numEtiquetas, 5);
eq('1 tom: páginas no PDF', r.paginas.length, 5);
eq('1 tom: a 1ª via da reposição fecha a contagem', loteDe(r.paginas[3]), '4/4');
ok('1 tom: as duas últimas são a reposição', ehRep(r.paginas[3]) && ehRep(r.paginas[4]), r.paginas.map(ehRep));
ok('1 tom: a via extra não leva lote', !temLote(r.paginas[4]), r.paginas[4]);
eq('1 tom: o tom sai resumido numa linha', linhaTons(r.paginas[3]), 'TOM: 1');

/* ---------- 2. dois tons: cada tamanho rende um pacote por tom ---------- */

r = etiquetasDe(osBase(), { tons: [1, 2] });
eq('2 tons: pacotes contados (3×2 + reposição)', r.dados.totalPacotes, 7);
eq('2 tons: etiquetas impressas', r.dados.numEtiquetas, 8);
eq('2 tons: páginas no PDF', r.paginas.length, 8);
eq('2 tons: o último pacote é a reposição', loteDe(r.paginas[6]), '7/7');
ok('2 tons: a via extra não leva lote', !temLote(r.paginas[7]), r.paginas[7]);
eq('2 tons: o resumo dos tons na 1ª via', linhaTons(r.paginas[6]), 'TONS: 1 · 2');
eq('2 tons: o resumo dos tons na via extra', linhaTons(r.paginas[7]), 'TONS: 1 · 2');
ok('2 tons: as duas vias da reposição são iguais',
   JSON.stringify(r.paginas[6].filter(l => !/^LOTE:/.test(l))) === JSON.stringify(r.paginas[7]),
   [r.paginas[6], r.paginas[7]]);

// A etiqueta de TAMANHO segue com o tom no destaque, e não na linha de resumo:
// é ela que distingue dois pacotes do mesmo tamanho na hora de ensacar.
ok('2 tons: a etiqueta de tamanho leva o tom no destaque',
   r.paginas[0].some(l => /^[PMG]( tom \d)?$/.test(l) && / tom /.test(l)), r.paginas[0]);
ok('2 tons: a etiqueta de tamanho não repete a linha de resumo',
   linhaTons(r.paginas[0]) === '', r.paginas[0]);

/* ---------- 3. três tons: o resumo cresce, a contagem também ---------- */

r = etiquetasDe(osBase(), { tons: [1, 2, 3] });
eq('3 tons: pacotes contados (3×3 + reposição)', r.dados.totalPacotes, 10);
eq('3 tons: etiquetas impressas', r.dados.numEtiquetas, 11);
eq('3 tons: o resumo lista os três', linhaTons(r.paginas[10]), 'TONS: 1 · 2 · 3');
ok('3 tons: a via extra continua sem lote', !temLote(r.paginas[10]), r.paginas[10]);

/* ---------- 4. sem tonalidade registrada: a linha some ---------- */

r = etiquetasDe(osBase(), { tons: [] });
eq('sem tom: a OS ainda rende os pacotes de tamanho', r.dados.totalPacotes, 4);
eq('sem tom: nenhuma linha de tom é inventada', linhaTons(r.paginas[3]), '');
ok('sem tom: a via extra segue sem lote', !temLote(r.paginas[4]), r.paginas[4]);

/* ---------- 5. moletom: 1 pacote por tamanho, composição só nas de tamanho ---------- */

r = etiquetasDe(osBase({ fases: [{ tecidoId: 't1' }] }), { tons: [1], moletom: true });
eq('moletom: 1 pacote por tamanho + reposição', r.dados.totalPacotes, 4);
ok('moletom: a etiqueta de tamanho traz a composição',
   r.paginas[0].some(l => /Frente 36/.test(l)), r.paginas[0]);
ok('moletom: a reposição não traz composição',
   !r.paginas[4].some(l => /Frente 36/.test(l)), r.paginas[4]);

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
