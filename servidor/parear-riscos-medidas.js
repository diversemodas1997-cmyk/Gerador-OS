/*
 * Liga cada FASE do cadastro ao PDF DE ENCAIXE que deu a medida dela
 * (grava `fase.risco`, o mesmo campo que a importacao de risco preenche).
 *
 * POR QUE ISTO EXISTE
 *
 * Junior pediu (20/08/2026) que a geracao de OS seja BLOQUEADA quando a medida
 * da grade nao veio de um PDF. O campo que responde isso, `fase.risco`, so
 * passou a ser gravado em 19/08/2026: medido antes de comecar, 327 fases tinham
 * comprimento e largura e ZERO tinham o PDF registrado. Ligar a trava naquele
 * estado pararia a fabrica inteira. Este script e o passo que falta antes dela.
 *
 * O QUE ELE FAZ, E O QUE ELE NAO FAZ
 *
 * Ele NAO reescreve medida nenhuma. A medida cadastrada e o que a fabrica usa;
 * trocar 327 numeros por conta propria seria o oposto de tornar o cadastro
 * confiavel. Ele so PROCURA A PROVA: entre os relatorios de encaixe da pasta,
 * qual deles produz exatamente o comprimento e a largura que ja estao ali.
 *
 * Fase cujo numero nao sai de PDF nenhum fica SEM prova, de proposito — e e
 * exatamente essa que a trava vai pegar depois. Nao inventar prova e o ponto.
 *
 * COMO ELE DECIDE
 *
 *   1. LARGURA: a do relatorio, sem nada somado (|dif| <= 2 cm, a mesma folga
 *      que a tela usa).
 *   2. COMPRIMENTO: o do relatorio MAIS o excedente daquela fase — a mesma
 *      conta da tela (`excedenteEnfestoM`, recortada do app.js para nao existir
 *      uma segunda versao da regra) — contra o comprimento cadastrado.
 *   3. LINHA: a pasta do PDF (BM.LISA, CM.TRI...) tem que ser a linha do nome
 *      da grade. Sem isto a "P ao G3" da CM.TRI casa com a pasta da CM.LISA e
 *      a prova aponta para o produto errado — abrir o PDF errado e pior do que
 *      nao ter atalho.
 *
 * Empatando, desempata pela pasta de largura que bate com o nome da grade e
 * pelos tamanhos (casamento exato na frente do proporcional, que e o caso da
 * ribana 10x). Empate que sobra e RELATADO, nao chutado.
 *
 * COMO RODAR
 *
 *   node servidor/parear-riscos-medidas.js             so relata, nao grava
 *   node servidor/parear-riscos-medidas.js --gravar    grava no servidor
 *   node servidor/parear-riscos-medidas.js --csv x.csv relatorio em planilha
 *
 * A gravacao e read-modify-write com CONFERENCIA DE CARIMBO: se alguem gravou
 * no servidor entre a leitura e a escrita, ele aborta em vez de passar por cima
 * (ver project_restauracao_credenciada — o blob e uma linha so, compartilhada).
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const PASTA_RISCOS = path.join(RAIZ, 'Desenhos técnicos -grades de corte');
const VENDOR = path.join(RAIZ, 'vendor');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';

const GRAVAR = process.argv.includes('--gravar');
const CSV = (() => { const i = process.argv.indexOf('--csv'); return i > 0 ? process.argv[i + 1] : ''; })();

const TOL_LARG = 0.02;   // 2 cm — a largura e a do pano, quase nao varia
const TOL_COMP = 0.02;   // 2 cm — o comprimento cadastrado e arredondado em 2 casas

/* ---- as REGRAS vem do app.js, recortadas: uma so versao de cada uma ---- */
function regrasDoApp(STATE) {
  const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
  const de = src.indexOf('const EXCEDENTE_ENFESTO_PADRAO_CM');
  const ate = src.indexOf('function _pastaMontarGrupos');
  if (de < 0 || ate < 0) throw new Error('nao achei o bloco de excedente/risco no app.js');
  // Duas ajudantes que o bloco usa e que moram longe dele no arquivo.
  const solta = nome => {
    const i = src.indexOf('function ' + nome + '(');
    if (i < 0) throw new Error('nao achei ' + nome + ' no app.js');
    return src.slice(i, src.indexOf('\n}', i) + 2);
  };
  const normNome = solta('_normNome') + solta('_normFaseNome');
  return new Function('STATE', 'esc', 'document', 'window', `
    ${normNome}
    ${src.slice(de, ate)}
    return { excedenteEnfestoM, _riscoGradesQueCasam, _riscoGradesProporcionais, _normNome };
  `)(STATE, s => String(s == null ? '' : s),
     { getElementById: () => null, querySelectorAll: () => [] }, {});
}

/* ---- leitura dos relatorios de encaixe ---------------------------------- */
function todosOsPdfs(dir, acc = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) todosOsPdfs(p, acc);
    else if (/\.pdf$/i.test(e.name)) acc.push(p);
  }
  return acc;
}

// O pdf.js do vendor roda no Node com estes dois empurroes: um `window` e o
// caminho explicito do worker (sem ele a biblioteca procura './pdf.worker.js'
// ao lado dela e morre com "Setting up fake worker failed").
function carregarPdfjs() {
  global.window = global;
  global.navigator = { userAgent: 'node' };
  const lib = require(path.join(VENDOR, 'pdf-3.11.174.min.js'));
  lib.GlobalWorkerOptions.workerSrc = path.join(VENDOR, 'pdf.worker-3.11.174.min.js');
  return lib;
}

const num = txt => {
  const m = String(txt == null ? '' : txt).replace(',', '.').match(/-?\d+(?:\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
};

// Comprimento, largura e tamanhos de um relatorio. As FILAS sao faixas de
// altura, e nao linhas exatas: o CAD desalinha o tamanho em 1 ponto do resto da
// fila (a mesma razao do FILA=3 da tela).
async function lerPdf(pdfjsLib, arq) {
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(fs.readFileSync(arq)) }).promise;
  const tc = await (await doc.getPage(1)).getTextContent();
  const itens = tc.items.filter(i => (i.str || '').trim())
    .map(i => ({ t: i.str.trim(), x: i.transform[4], y: i.transform[5] }));
  const filas = [];
  itens.slice().sort((a, b) => b.y - a.y).forEach(i => {
    const f = filas[filas.length - 1];
    if (f && Math.abs(f.y - i.y) <= 3) { f.itens.push(i); return; }
    filas.push({ y: i.y, itens: [i] });
  });
  filas.forEach(f => f.itens.sort((a, b) => a.x - b.x));
  const txt = f => f.itens.map(i => i.t).join(' ');
  const inteiro = filas.map(txt).join(' ').replace(/\s+/g, ' ');
  const campo = rotulo => {
    const m = inteiro.match(new RegExp(rotulo + ':\\s*([\\d.,]+)'));
    return m ? num(m[1]) : null;
  };
  const iCab = filas.findIndex(f => /(^|\s)Tamanho(\s|$)/i.test(txt(f)) && /Completos/i.test(txt(f)));
  const tamanhos = {};
  const moldes = {};
  let temCompletos = false;
  if (iCab >= 0) {
    const iFim = filas.findIndex((f, k) => k > iCab && /^Encaixe$/i.test(txt(f).trim()));
    const ate = iFim > iCab ? iFim : filas.length;
    for (let k = iCab + 1; k < ate; k++) {
      const toks = filas[k].itens;
      const iTam = toks.findIndex(t => /^(P|M|G|GG|G1|G2|G3)$/i.test(t.t));
      if (iTam < 0) continue;
      const nums = toks.slice(iTam + 1).filter(t => /^\d+$/.test(t.t));
      const chave = toks[iTam].t.toLowerCase();
      const nComp = nums[0] ? parseInt(nums[0].t, 10) : 0;
      const nMold = nums[1] ? parseInt(nums[1].t, 10) : 0;
      if (nComp > 0) { tamanhos[chave] = nComp; temCompletos = true; }
      if (nMold > 0) moldes[chave] = nMold;
    }
  }
  // ENCAIXE PARCIAL: a fase que corta so ALGUMAS pecas da roupa sai com a coluna
  // Completos toda zerada (nenhuma peca inteira cabe ali) — e o "CM.REC - CORPO
  // 2" e assim. A distribuicao continua na coluna dos MOLDES: total de moldes
  // dividido por modelos pedidos da quantos moldes tem uma roupa. Mesma regra da
  // tela; so vale quando a divisao fecha EXATA em todos os tamanhos.
  if (!temCompletos && Object.keys(moldes).length) {
    const pedidos = (inteiro.match(/Modelos pedidos:\s*(\d+)/) || [])[1];
    const total = Object.values(moldes).reduce((s, v) => s + v, 0);
    const porRoupa = pedidos > 0 ? total / parseInt(pedidos, 10) : 0;
    if (Number.isInteger(porRoupa) && porRoupa > 0
        && Object.values(moldes).every(v => v % porRoupa === 0)) {
      Object.keys(moldes).forEach(k => { tamanhos[k] = moldes[k] / porRoupa; });
    }
  }
  const comp = campo('Comprimento'), larg = campo('Largura');
  const rel = path.relative(PASTA_RISCOS, arq).split(path.sep).join('/');
  return {
    rel,
    linha: (rel.split('/')[0] || '').toUpperCase(),
    pastaLargura: (String(rel.split('/')[2] || '').match(/(\d+[.,]?\d*)\s*cm/i) || [])[1] || '',
    comprimento: comp != null ? comp / 100 : null,
    largura: larg != null ? larg / 100 : null,
    tamanhos
  };
}

/* ---- o servidor da fabrica --------------------------------------------- */
async function lerBlob() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at, cab };
}

// Grava de volta CONFERINDO O CARIMBO: o filtro por updated_at faz o servidor
// recusar a escrita se alguem gravou nesse meio-tempo. Sem isto, o blob inteiro
// (uma linha so, de todo mundo) seria sobrescrito com o que lemos ha um minuto.
async function gravarBlob(cab, data, updatedAt) {
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
}

/* ---- o pareamento ------------------------------------------------------- */
const f2 = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : null; };

// A linha da grade, do nome ("2M-4G-2GG | BM.LISA | 177cm" -> BM.LISA).
function linhaDaGrade(nome) {
  const partes = String(nome || '').split('|');
  return (partes[1] || '').trim().toUpperCase();
}
// A largura escrita no nome da grade, em cm ("... | 177cm" -> "177").
function larguraDoNome(nome) {
  const m = String(nome || '').match(/(\d+[.,]?\d*)\s*cm/i);
  return m ? m[1].replace(',', '.') : '';
}

(async () => {
  console.log('Lendo os relatorios de encaixe...');
  const pdfjsLib = carregarPdfjs();
  const arqs = todosOsPdfs(PASTA_RISCOS);
  const riscos = [];
  for (const a of arqs) {
    try {
      const r = await lerPdf(pdfjsLib, a);
      if (r.comprimento != null && r.largura != null) riscos.push(r);
    } catch (e) { console.log('  ! nao consegui ler ' + path.basename(a) + ': ' + e.message); }
  }
  console.log(`  ${riscos.length} de ${arqs.length} relatorios com comprimento e largura.`);

  console.log('Lendo o cadastro do servidor...');
  const { data, updatedAt, cab } = await lerBlob();
  const grades = JSON.parse(data.grades || '[]');
  const R = regrasDoApp({ grades });

  let comMedida = 0, achadas = 0, jaTinha = 0, semProva = 0, empatadas = 0;
  const relatorio = [];

  grades.forEach(g => {
    const linhaG = linhaDaGrade(g.nome);
    const cmNome = larguraDoNome(g.nome);
    const exatas = new Set(R._riscoGradesQueCasam.length ? [] : []);   // (so para clareza abaixo)
    (g.fases || []).forEach(fase => {
      const comp = f2(fase.comp), larg = f2(fase.larg);
      if (!(comp > 0 && larg > 0)) return;
      comMedida++;
      if ((fase.risco || '').trim()) { jaTinha++; return; }

      // O RELATORIO TEM QUE SER DESTA GRADE, e nao apenas dar o mesmo numero.
      // Sem esta exigencia, "Corpo Parte 2" — a peca pequena, que mede 0,62 em
      // meia duzia de grades — casava pela medida com o PDF da grade vizinha e a
      // prova apontava para o produto errado. Prova errada e pior do que prova
      // nenhuma: quem for conferir a medida abre o arquivo de outra roupa e
      // confirma o engano.
      const daGrade = r =>
        R._riscoGradesQueCasam(r.tamanhos).some(x => x.id === g.id) ? 'exato'
        : R._riscoGradesProporcionais(r.tamanhos).some(x => x.grade.id === g.id) ? 'proporcional'
        : '';
      const cands = riscos.filter(r => {
        if (r.linha !== linhaG) return false;                       // a LINHA e obrigatoria
        if (!daGrade(r)) return false;                              // e os TAMANHOS, tambem
        if (Math.abs(r.largura - larg) > TOL_LARG) return false;
        const compCad = r.comprimento + R.excedenteEnfestoM(fase, r.comprimento);
        return Math.abs(compCad - comp) <= TOL_COMP;
      });

      if (!cands.length) {
        semProva++;
        relatorio.push([g.nome, fase.nome, comp, larg, 'SEM PROVA', '']);
        return;
      }
      // Desempate: a pasta de largura do nome da grade, depois os tamanhos
      // (exato antes de proporcional — a ribana 10x e proporcional).
      const pontos = r => {
        let p = 0;
        if (cmNome && r.pastaLargura && r.pastaLargura.replace(',', '.') === cmNome) p += 2;
        if (daGrade(r) === 'exato') p += 3; else p += 2;   // proporcional e a ribana 10x
        return p;
      };
      const ord = cands.map(r => ({ r, p: pontos(r) })).sort((a, b) => b.p - a.p);
      const topo = ord.filter(x => x.p === ord[0].p);
      if (topo.length > 1) empatadas++;
      fase.risco = topo[0].r.rel;
      achadas++;
      relatorio.push([g.nome, fase.nome, comp, larg,
        topo.length > 1 ? `ACHADA (${topo.length} empatadas)` : 'ACHADA', topo[0].r.rel]);
    });
  });

  console.log('');
  console.log('=== RESULTADO ===');
  console.log(`fases com medida cadastrada : ${comMedida}`);
  console.log(`ja tinham o PDF registrado  : ${jaTinha}`);
  console.log(`PROVA ACHADA agora          : ${achadas}   (${empatadas} decididas por desempate)`);
  console.log(`SEM PROVA (a trava vai pegar): ${semProva}`);
  const cobertura = comMedida ? Math.round((jaTinha + achadas) * 100 / comMedida) : 0;
  console.log(`cobertura depois de gravar  : ${cobertura}%`);

  // Quantas GRADES ficariam 100% provadas — e a conta que a trava da OS usa.
  let gOk = 0, gFalta = 0;
  grades.forEach(g => {
    const comMed = (g.fases || []).filter(f => f2(f.comp) > 0 && f2(f.larg) > 0);
    if (!comMed.length) return;
    if (comMed.every(f => (f.risco || '').trim())) gOk++; else gFalta++;
  });
  console.log(`grades com TODAS as fases provadas: ${gOk} | com alguma sem prova: ${gFalta}`);

  if (CSV) {
    const linhas = [['grade', 'fase', 'comp', 'larg', 'situacao', 'pdf']].concat(relatorio);
    fs.writeFileSync(CSV, '\uFEFF' + linhas.map(l => l.map(c => `"${String(c).replace(/"/g, '""')}"`).join(';')).join('\r\n'));
    console.log(`\nrelatorio: ${CSV}`);
  }

  if (!GRAVAR) {
    console.log('\n(nada foi gravado — rode com --gravar para valer)');
    return;
  }
  data.grades = JSON.stringify(grades);
  await gravarBlob(cab, data, updatedAt);
  console.log('\nGRAVADO no servidor.');
})().catch(e => { console.error('ERRO: ' + (e && e.message || e)); process.exit(1); });
