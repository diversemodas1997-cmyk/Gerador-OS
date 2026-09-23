/*
 * ZERA O ESTOQUE NEGATIVO E COBRE AS FALTAS DO QUADRO DE MATERIAL RESERVADO.
 *
 * Junior, 23/09/2026: "cadastre entrada dos tecidos em negativo para zerar o
 * estoque. Cadastre tambem os tecidos apontados faltantes na reserva de
 * materiais."
 *
 * Duas levas de ENTRADA MANUAL, gravadas juntas:
 *
 *   1. prateleira (tecido + cor) com saldo negativo recebe exatamente o kg que
 *      falta para parar em zero — a mesma regra de zerar-negativos-estoque.js;
 *   2. depois disso, as OS do quadro "material reservado" que ainda aparecem
 *      em vermelho (compraFaltasAbertas: faltaDeTecidoParaOS de cada OS
 *      reservada e nao consumida). A falta de uma OS e o que ela precisa menos o
 *      disponivel SEM a propria reserva — entao uma entrada de X kg numa
 *      prateleira reduz em X a falta de TODAS as OS daquela prateleira. Por
 *      isso a entrada e a MAIOR falta da prateleira, e nao a soma: a soma
 *      deixaria sobra de pano que nao existe.
 *
 * Nenhuma saida e apagada. A conta e toda do app.js, recortada na hora.
 *
 *   node servidor/zerar-faltas-reserva.js             so relata
 *   node servidor/zerar-faltas-reserva.js --gravar    grava (backup antes, carimbo conferido)
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
const OBS_NEG = 'Ajuste: entrada para zerar saldo negativo';
const OBS_FALTA = 'Ajuste: entrada para cobrir falta da reserva de material';

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
    + '\nreturn { calcularSaldosEstoque, compraFaltasAbertas };')(STATE, comprasCache);
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
const entrada = (tecidoNome, corNome, kg, obs) => ({
  id: uid(), tipo: 'entrada', tecidoNome, corNome, kg: kg3(kg), fechados: 0, abertos: 0,
  data: HOJE, origem: 'manual', osId: '', osNumero: '', obs
});

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const compras = await lerCompras(cab);
  // Cada chave do blob e TEXTO JSON. O STATE leva todas: a conta do consumo
  // le grades, ordens, tecidos, cores e as regras de configuracao.
  const base = {};
  Object.keys(data).forEach(k => {
    try { base[k] = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; } catch (e) { base[k] = data[k]; }
  });
  ['cores', 'tecidos', 'grades', 'ordens', 'desenhos', 'estoqueMov'].forEach(k => {
    if (!Array.isArray(base[k])) base[k] = [];
  });
  const conta = mov => contaDoApp({ ...base, estoqueMov: mov }, compras);

  // 1. prateleiras negativas
  const negativos = conta(base.estoqueMov).calcularSaldosEstoque().detalhe
    .filter(d => kg3(d.disponivel) < 0);
  const lev1 = negativos.map(d => entrada(d.tecidoNome || '', d.corNome || '', -d.disponivel, OBS_NEG));
  console.log('\n1. SALDOS NEGATIVOS (' + lev1.length + ')');
  lev1.forEach(m => console.log('  ' + m.kg.toFixed(3).padStart(10) + ' kg   ' + m.tecidoNome + ' · ' + m.corNome));

  // 2. faltas do quadro do reservado, ja contando a leva 1
  const mov1 = base.estoqueMov.concat(lev1);
  const faltas = conta(mov1).compraFaltasAbertas();
  const porPrat = new Map();
  faltas.forEach(f => f.tecidos.forEach(t => {
    const k = String(t.tecidoNome).toLowerCase() + '||' + String(t.corNome).toLowerCase();
    const cur = porPrat.get(k) || { tecidoNome: t.tecidoNome, corNome: t.corNome, kg: 0, os: [] };
    cur.kg = Math.max(cur.kg, t.falta);
    cur.os.push(f.osNumero + ' (' + t.falta.toFixed(3) + ')');
    porPrat.set(k, cur);
  }));
  const pr2 = [...porPrat.values()].filter(p => p.kg > 0.0005);
  const lev2 = pr2.map(p => entrada(p.tecidoNome, p.corNome, p.kg, OBS_FALTA));
  console.log('\n2. FALTAS NA RESERVA DE MATERIAL (' + faltas.length + ' OS, ' + lev2.length + ' prateleiras)');
  pr2.forEach((p, i) => console.log('  ' + lev2[i].kg.toFixed(3).padStart(10) + ' kg   ' +
    p.tecidoNome + ' · ' + p.corNome + '   OS ' + p.os.join(', ')));

  const novos = lev1.concat(lev2);
  const tot = novos.reduce((a, m) => a + m.kg, 0);
  console.log('\n  total: ' + kg3(tot).toFixed(3) + ' kg em ' + novos.length + ' entradas, data ' + HOJE);

  // Confere pela propria conta do app: nenhum negativo e nenhuma OS em falta.
  const movF = base.estoqueMov.concat(novos);
  const cF = conta(movF);
  const neg = cF.calcularSaldosEstoque().detalhe.filter(d => kg3(d.disponivel) < 0);
  const fal = cF.compraFaltasAbertas();
  console.log('  conferencia: ' + neg.length + ' negativos e ' + fal.length + ' OS em falta depois das entradas');
  if (neg.length || fal.length) throw new Error('conferencia falhou: ' + JSON.stringify({ neg, fal }).slice(0, 600));

  if (!novos.length) { console.log('Nada a fazer.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-zerar-faltas-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.estoqueMov = JSON.stringify(movF);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + novos.length + ' entradas manuais.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
