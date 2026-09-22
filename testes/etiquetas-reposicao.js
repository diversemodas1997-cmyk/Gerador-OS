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
  corta('function _composicaoPacoteMoletom'),
  corta('function totaisPorTamanhoTomOS'),
  listaConst('_PECAS_ETIQUETA_BMTRI'),
  corta('function _normNome'),
  corta('function _normSku'),
  corta('function _skuDaGrade'),
  corta('function _osEhBmTri'),
  corta('function _coresDaPecaOS'),
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
    const multiplicadorPecaOS = () => (moletom ? 1 : 2);
    // Sem grade no cadastro: o tipo (BM.TRI…) é lido do nome copiado na OS.
    const _skuDaOS = () => '';
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
ok('moletom: a etiqueta de tamanho traz a composição do pacote (1 blusa × 10 camadas)',
   r.paginas[0].includes('Frente 10 · Costa 10 · Bolso 10 · Barra 10') &&
   r.paginas[0].includes('Mangas 20 · Capuz 20 · Punhos 20'), r.paginas[0]);
ok('moletom: a reposição não traz composição',
   !r.paginas[4].some(l => /Frente/.test(l)), r.paginas[4]);

/* ---------- 6. OS 0547: a composição acompanha a tonalidade ----------
   8G, 28 camadas, Tom 1 com 144 (18 camadas) e o Tom 2 balanceando 80 (10).
   Até 22/09/2026 as duas etiquetas diziam "Frente 36 … Mangas 72", um número
   fixo que não vinha da OS. */

const os547 = osBase({
  os: '0547', fases: [{ tecidoId: 't1' }],
  grade: { g: 8, total: 8 }, enfesto: { camadas: 28 },
  progresso: { totalTamanhoTons: { 1: true, 2: true }, totalTamanhoTomValor: { 1: 144 } }
});
r = etiquetasDe(os547, { tons: [1, 2], moletom: true });
eq('0547: pacotes (G × 2 tons + reposição)', r.dados.totalPacotes, 3);
ok('0547: G tom 1 leva 144 blusas',
   r.paginas[0].includes('G tom 1') &&
   r.paginas[0].includes('Frente 144 · Costa 144 · Bolso 144 · Barra 144') &&
   r.paginas[0].includes('Mangas 288 · Capuz 288 · Punhos 288'), r.paginas[0]);
ok('0547: G tom 2 leva as 80 que sobram',
   r.paginas[1].includes('G tom 2') &&
   r.paginas[1].includes('Frente 80 · Costa 80 · Bolso 80 · Barra 80') &&
   r.paginas[1].includes('Mangas 160 · Capuz 160 · Punhos 160'), r.paginas[1]);

// Dois tons sem a divisão digitada: não há número certo a imprimir.
r = etiquetasDe(osBase({ fases: [{ tecidoId: 't1' }], grade: { g: 8, total: 8 }, enfesto: { camadas: 28 },
  progresso: { totalTamanhoTons: { 1: true, 2: true } } }), { tons: [1, 2], moletom: true });
ok('2 tons sem divisão: a composição sai sem número, não com chute',
   r.paginas[0].includes('Frente · Costa · Bolso · Barra'), r.paginas[0]);

/* ---------- 7. BM.TRI: uma etiqueta por PEÇA (pedido de 22/09/2026) ----------
   Cada tamanho × tom rende sete etiquetas: Frente, Costa, Capuz, Forro de capuz,
   Barra/Punhos, Mangas e Bolso — cada uma com a cor e a quantidade da peça.
   Os componentes são os da OS 0547 como estão gravados. */

const comps547 = [
  ['Capuz', 'Preto'], ['Forro do capuz', 'Preto'], ['Bolso canguru', 'Off-White'],
  ['Punho', 'Off-White'], ['Barra', 'Off-White'], ['Viés', 'Preto'],
  ['Frente PARTE 3 Blusa Moletom Tricolor', 'Off-White'], ['Frente PARTE 1 Blusa Moletom Tricolor', 'Preto'],
  ['Frente PARTE 2 Blusa Moletom Tricolor', 'Mostarda'], ['Costa parte 1', 'Preto'],
  ['Costa parte 2', 'Mostarda'], ['Costa parte 3', 'Off-White'],
  ['Mangas parte 1', 'Preto'], ['Mangas parte 2', 'Mostarda'], ['Mangas parte 3', 'Off-White']
].map(([nome, corNome]) => ({ nome, corNome }));
const tri547 = osBase({
  os: '0547', fases: [{ tecidoId: 't1' }], componentes: comps547,
  grade: { descricao: '8G | BM.TRI | 179cm', g: 8, total: 8 }, enfesto: { camadas: 28 },
  progresso: { totalTamanhoTons: { 1: true, 2: true }, totalTamanhoTomValor: { 1: 144 } }
});
r = etiquetasDe(tri547, { tons: [1, 2], moletom: true });
eq('BM.TRI: etiquetas (1 tamanho × 2 tons × 7 peças + 2 de reposição)', r.dados.numEtiquetas, 16);
eq('BM.TRI: pacotes no LOTE (14 peças + reposição)', r.dados.totalPacotes, 15);
const ordem = r.paginas.slice(0, 7).map(pg => pg[pg.length - 1]);
eq('BM.TRI: a ordem das peças no G tom 1', ordem.join(' | '),
   'FRENTE | COSTA | CAPUZ | FORRO DE CAPUZ | BARRA/PUNHOS | MANGAS | BOLSO');
ok('BM.TRI: G tom 1 — Frente, 144, nas três cores em ordem de parte',
   r.paginas[0].includes('G tom 1') && r.paginas[0].includes('QTDE: 144') &&
   r.paginas[0].includes('COR: PRETO/MOSTARDA/OFF-WHITE'), r.paginas[0]);
ok('BM.TRI: G tom 1 — Capuz é preto e vai em dobro',
   r.paginas[2].includes('QTDE: 288') && r.paginas[2].includes('COR: PRETO'), r.paginas[2]);
ok('BM.TRI: G tom 1 — Barra/Punhos diz as duas contas',
   r.paginas[4].includes('QTDE: Barra 144 · Punhos 288') && r.paginas[4].includes('COR: OFF-WHITE'), r.paginas[4]);
ok('BM.TRI: G tom 2 — Mangas com as 80 blusas do tom',
   r.paginas[12].includes('G tom 2') && r.paginas[12].includes('MANGAS') && r.paginas[12].includes('QTDE: 160'), r.paginas[12]);
ok('BM.TRI: a etiqueta de peça não traz a lista de composição nem a grade inteira',
   !r.paginas[0].some(l => /Punhos d+ ·|^TAM:/.test(l)), r.paginas[0]);
ok('BM.TRI: a reposição continua no fim, sem peça',
   ehRep(r.paginas[14]) && ehRep(r.paginas[15]) && !r.paginas[14].includes('FRENTE'), r.paginas[14]);

// Outro moletom que não é BM.TRI segue com uma etiqueta por tamanho.
r = etiquetasDe(osBase({ fases: [{ tecidoId: 't1' }], componentes: comps547,
  grade: { descricao: 'P ao G3 | BM.LISA | 177cm', g: 8, total: 8 }, enfesto: { camadas: 28 } }),
  { tons: [1], moletom: true });
eq('BM.LISA: uma etiqueta por tamanho, como antes', r.dados.numEtiquetas, 3);
ok('BM.LISA: sem etiqueta por peça', r.dados.pecasPacotes == null, r.dados.pecasPacotes);

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
