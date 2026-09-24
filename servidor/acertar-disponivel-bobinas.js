/*
 * ACERTA O DISPONIVEL DE UMA PRATELEIRA PARA N BOBINAS.
 *
 * Junior, 24/09/2026: "corrija o estoque disponivel de tecido cor grafite
 * para 23 bobinas disponiveis" e "... cor marrom para 9 bobinas".
 *
 * O disponivel da tela e kg, e as bobinas ao lado sao kg / peso da bobina
 * (pesoBobinaEstimado: o cadastro do tecido ou, zerado, a MEDIANA das entradas
 * com kg e fechados). O acerto e um lancamento MANUAL, igual ao da tela --
 * saida quando sobra, entrada quando falta -- que deixa o disponivel em
 * exatamente N x peso da bobina. Nenhum lancamento antigo e apagado.
 *
 *   node servidor/acertar-disponivel-bobinas.js "Malha Algodão|grafite=23"
 *   node servidor/acertar-disponivel-bobinas.js "Malha Algodão|grafite=23" --gravar
 *
 * Cada argumento e "tecido|pedaco da cor=bobinas". O pedaco da cor tem de
 * casar UMA prateleira so daquele tecido, senao o script para e lista.
 *
 * Depois de gravar: F5 no programa, senao aba aberta regrava o estoque velho.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const HOJE = new Date().toISOString().slice(0, 10);
const PEDIDOS = process.argv.slice(2).filter(a => !a.startsWith('--'));

/* ---- a conta vem do app.js, recortada: uma so versao dela ---- */
function contaDoApp(STATE, comprasCache) {
  const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
  const fn = n => {
    const i = src.indexOf('function ' + n + '(');
    if (i < 0) throw new Error('nao achei ' + n + ' no app.js');
    return src.slice(i, src.indexOf('\n}', i) + 2);
  };
  const cst = n => {
    const m = src.match(new RegExp('^const ' + n + ' = [^;]+;', 'm'));
    if (!m) throw new Error('nao achei a constante ' + n);
    return m[0];
  };
  const CONSTS = ['LIMITE_CAMADAS', 'MULTIPLICADOR_PECAS', 'LABEL_CATEGORIA',
                  'UNIDADES_PADRAO_FORRO', 'CEIL_BOBINA_EPS', '_EXC_LIGACAO', '_PAL_VIES'];
  const FNS = ['_normNome', '_sufixoTecidoNorm', 'corBaseNome', 'corCanonicaPorTecido',
    'categoriaEfetivaTecido', 'isTecidoRibana', 'calcularPapeisFases',
    '_tamanhoQueMandaNaGrade', 'camadasDaFaseRibana', '_ribanaEscalaComGrade',
    'camadasDaFaseForro', 'camadasPadraoDaFase', 'camadasCheiasDaFase',
    'multiplicadorPecaOS', '_faseNaoEnfestadaPorTom', 'unidadesPorCamadaTecido',
    'unidadesPorCamadaPrincipal', 'tecidosDaOS', 'gramaturaTecidoPorNome',
    'pesoBobinaPorNome', '_normFaseNome', '_faseSoDe', 'compraLimiteCamadasGrade',
    'consumoEnfestoOS', 'parseBobinas', 'bobinaInteira', 'ehFaseRibana',
    'bobinasEfetivasFase', 'consumoAgregadoPorTecidoCor',
    'comprasComoMovimentos', 'movimentacoesEstoque', 'calcularSaldosEstoque',
    'osComMaterialReservado', 'faltaDeTecidoParaOS', 'compraFaltasAbertas'];
  return new Function('STATE', 'comprasCache',
    CONSTS.map(cst).join('\n') + '\n' + FNS.map(fn).join('\n')
    + '\n' + ['larguraPadraoTecido', 'pesoBobinaEstimado'].map(fn).join('\n')
    + '\nreturn { calcularSaldosEstoque, compraFaltasAbertas, pesoBobinaEstimado };')(STATE, comprasCache);
}

/* ---- o servidor da fabrica ---- */
async function conectar() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  return { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
}

async function lerBlob(cab) {
  const l = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  return { data: l[0].data, updatedAt: l[0].updated_at };
}

async function lerCompras(cab) {
  try {
    const r = await fetch(`${SUPA}/rest/v1/compras_materiais?select=*`, { headers: cab });
    if (!r.ok) return [];
    const d = await r.json();
    return Array.isArray(d) ? d : [];
  } catch (e) { return []; }
}

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

let seq = 0;
const uid = () => 'aj' + Date.now().toString(36) + (seq++).toString(36) + Math.random().toString(36).slice(2, 6);
const kg3 = n => Math.round(n * 1000) / 1000;
const norm = s => String(s || '').trim().toLowerCase();

(async () => {
  if (!PEDIDOS.length) throw new Error('diga o que acertar: "tecido|cor=bobinas"');
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const compras = await lerCompras(cab);
  const base = {};
  Object.keys(data).forEach(k => {
    try { base[k] = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; } catch (e) { base[k] = data[k]; }
  });
  ['cores', 'tecidos', 'grades', 'ordens', 'desenhos', 'estoqueMov'].forEach(k => {
    if (!Array.isArray(base[k])) base[k] = [];
  });
  const conta = mov => contaDoApp({ ...base, estoqueMov: mov }, compras);
  const C = conta(base.estoqueMov);
  const detalhe = C.calcularSaldosEstoque().detalhe;

  const novos = [];
  const alvos = [];
  PEDIDOS.forEach(p => {
    const m = /^(.+)\|(.+)=(\d+)$/.exec(p);
    if (!m) throw new Error('argumento fora do formato "tecido|cor=bobinas": ' + p);
    const [, tec, cor, nTxt] = m;
    const n = parseInt(nTxt, 10);
    const achou = detalhe.filter(d => norm(d.tecidoNome) === norm(tec) && norm(d.corNome).includes(norm(cor)));
    console.log('\n' + tec + ' | ' + cor + ' -> ' + n + ' bobinas');
    if (achou.length !== 1) {
      console.log('  casou ' + achou.length + ' prateleiras. Com essa cor:');
      detalhe.filter(d => norm(d.corNome).includes(norm(cor))).forEach(d =>
        console.log('    ' + d.tecidoNome + ' :: ' + d.corNome + '   disponivel ' + d.disponivel.toFixed(3) + ' kg'));
      throw new Error('o pedido "' + p + '" tem de casar UMA prateleira');
    }
    const d = achou[0];
    const pb = C.pesoBobinaEstimado(d.tecidoNome);
    if (!pb || !(pb.kg > 0)) throw new Error('sem peso de bobina conhecido para ' + d.tecidoNome);
    const alvo = kg3(n * pb.kg);
    const dif = kg3(alvo - d.disponivel);
    console.log('  prateleira: ' + d.tecidoNome + ' :: ' + d.corNome);
    console.log('  bobina: ' + pb.kg + ' kg (' + (pb.origem === 'cadastro' ? 'cadastro' : 'mediana de ' + pb.n + ' entradas') + ')');
    console.log('  entradas ' + d.entrada.toFixed(3) + '  reservado ' + d.reservado.toFixed(3) + '  saidas ' + d.saida.toFixed(3));
    console.log('  disponivel hoje: ' + d.disponivel.toFixed(3) + ' kg = ' + Math.floor(d.disponivel / pb.kg + 1e-9) + ' bob');
    console.log('  alvo: ' + alvo.toFixed(3) + ' kg = ' + n + ' bob  ->  ' + (dif < 0 ? 'SAIDA' : 'ENTRADA') + ' de ' + Math.abs(dif).toFixed(3) + ' kg');
    if (Math.abs(dif) < 0.0005) { console.log('  ja esta certo.'); return; }
    novos.push({
      id: uid(), tipo: dif < 0 ? 'saida' : 'entrada', tecidoNome: d.tecidoNome, corNome: d.corNome,
      kg: Math.abs(dif), fechados: 0, abertos: 0, data: HOJE, origem: 'manual', osId: '', osNumero: '',
      obs: 'Ajuste: disponivel acertado para ' + n + ' bobinas (' + alvo.toFixed(3) + ' kg)'
    });
    alvos.push({ d, n, pb });
  });

  // Confere pela propria conta do app: a tela vai mostrar N bobinas.
  const movF = base.estoqueMov.concat(novos);
  const depois = conta(movF).calcularSaldosEstoque().detalhe;
  alvos.forEach(({ d, n, pb }) => {
    const x = depois.find(y => y.tecidoNome === d.tecidoNome && y.corNome === d.corNome);
    const bob = Math.floor((x.disponivel + 1e-9) / pb.kg);
    console.log('\nconferencia ' + d.corNome + ': ' + x.disponivel.toFixed(3) + ' kg = ' + bob + ' bob');
    if (bob !== n) throw new Error('conferencia falhou para ' + d.corNome);
  });

  if (!novos.length) { console.log('\nNada a fazer.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO -- nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-acertar-bobinas-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.estoqueMov = JSON.stringify(movF);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + novos.length + ' lancamento(s) manual(is).');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
