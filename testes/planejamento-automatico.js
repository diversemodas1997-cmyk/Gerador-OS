/* Rode com:  node testes/planejamento-automatico.js
   As contas do planejamento automatico (planejamento-automatico.js):
     - a demanda automatica dispara no minimo, pede ate o maximo e some quando
       a producao cobre;
     - o agrupamento so vira planejamento quando fecha 1 enfesto cheio;
     - o planejamento montado sem o formulario tem a grade, as camadas, as
       cores do desenho e o contexto da OS-modelo — e nada do que e da OS
       velha (numero, lote do pano, observacao). */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'planejamento-automatico.js'), 'utf8');

/* ------------------------------ stubs ------------------------------------ */
global.window = {};
global.STATE = {
  meta: {},
  cores: [{ id: 'c-pre', nome: 'Preto' }, { id: 'c-bra', nome: 'Branco' }],
  tecidos: [{ id: 't-mol', nome: 'Moletom' }, { id: 't-rib', nome: 'Ribana' }],
  modelos: [{ id: 'm-cam', nome: 'Camiseta Básica' }],
  componentes: [], materiais: [{ id: 'av-etq', desc: 'Etiqueta' }],
  desenhos: [], ordens: [], grades: []
};
global.ordenarCoresIdsPorDesc = ids => ids;
global.unidadesPorCamadaPrincipal = () => 1;
global._aplicarTonsPadrao = os => { os.progresso = { totalTamanhoTons: { 1: true } }; return os; };
// O que um enfesto de uma parte rende (a mesma conta do app.js).
global._demRende = p => {
  const r = {};
  Object.keys(p.tams).forEach(k => { if (p.tams[k] > 0) r[k] = p.tams[k] * (p.cam || 1) * (p.mult || 1); });
  return r;
};

eval(src);
const { calcularDemandas, aglomerou, montarPlano, idDoPlano, PREFIXO_DEM } = window._autoPlanInterno;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};
const AGORA = '2026-10-30T10:00:00.000Z';
const linha = o => Object.assign({ sku: 'CM.LISA-PRE', tam: 'g', desenhoId: 'd1', saldo: 140,
  min: 150, max: 1000, wip: 0, planejado: 0, manual: 0 }, o);

/* -------------------------------- demanda --------------------------------- */
let r = calcularDemandas([linha()], new Map(), AGORA);
ok('1. saldo no minimo dispara a demanda', r.gravar.length === 1);
ok('1b. pede ate o maximo (1000 - 140 = 860)', r.gravar[0] && r.gravar[0].qtd === 860, JSON.stringify(r.gravar[0]));
ok('1c. id fixo por SKU e tamanho', r.gravar[0] && r.gravar[0].id === PREFIXO_DEM + 'CM.LISA-PRE:g');
ok('1d. marcada como automatica', r.gravar[0] && r.gravar[0].auto && r.gravar[0].auto.max === 1000);

r = calcularDemandas([linha({ saldo: 151 })], new Map(), AGORA);
ok('2. acima do minimo nao dispara', r.gravar.length === 0 && r.apagar.length === 0);

r = calcularDemandas([linha({ max: null })], new Map(), AGORA);
ok('3. linha sem maximo nao entra', r.gravar.length === 0);

r = calcularDemandas([linha({ wip: 300 })], new Map(), AGORA);
ok('4. em producao NAO sai da quantidade gravada (o agrupamento abate)', r.gravar[0] && r.gravar[0].qtd === 860,
   JSON.stringify(r.gravar[0]));

r = calcularDemandas([linha({ wip: 900 })], new Map(), AGORA);
ok('5. producao que ja repoe tudo: sem demanda', r.gravar.length === 0);

r = calcularDemandas([linha({ planejado: 500, manual: 100 })], new Map(), AGORA);
ok('6. planejado e demanda a mao saem da quantidade (1000-140-500-100 = 260)', r.gravar[0] && r.gravar[0].qtd === 260);

const ja = { id: PREFIXO_DEM + 'CM.LISA-PRE:g', desenhoId: 'd1', tam: 'g', qtd: 860, criadoEm: '2026-10-29T08:00:00.000Z' };
r = calcularDemandas([linha({ saldo: 160 })], new Map([[ja.id, ja]]), AGORA);
ok('7. ja disparada segue acima do minimo, ate a falta acabar', r.gravar[0] && r.gravar[0].qtd === 840 && !r.apagar.length);
ok('7b. guarda quando nasceu', r.gravar[0] && r.gravar[0].criadoEm === ja.criadoEm);

r = calcularDemandas([linha()], new Map([[ja.id, ja]]), AGORA);
ok('8. nada mudou: nao regrava', r.gravar.length === 0 && r.apagar.length === 0);

r = calcularDemandas([linha({ saldo: 1000 })], new Map([[ja.id, ja]]), AGORA);
ok('9. estoque reposto: a demanda automatica some', r.apagar.length === 1 && r.apagar[0] === ja.id);

r = calcularDemandas([linha({ desenhoId: '' })], new Map(), AGORA);
ok('10. SKU sem desenho nao vira demanda', r.gravar.length === 0);

/* ------------------------------ aglomerar --------------------------------- */
// Grade 1P-2M-2G-1GG (6 pecas por camada), moletom de 36 camadas: 216 por enfesto.
const parte = { tipo: 'existe', gradeId: 'g1', tams: { p: 1, m: 2, g: 2, gg: 1 }, cam: 36, mult: 1, c: 1 };
ok('11. menos que 1 enfesto: ainda acumula', !aglomerou({ partes: [parte], novas: 0 }, { g: 100, m: 100 }));
ok('11b. fechou 1 enfesto cheio: planeja', aglomerou({ partes: [parte], novas: 0 }, { g: 120, m: 100 }));
ok('11c. solucao com grade nova nao planeja sozinha', !aglomerou({ partes: [parte], novas: 1 }, { g: 500 }));
ok('11d. tecido sem limite de camadas nao planeja sozinho',
   !aglomerou({ partes: [Object.assign({}, parte, { cam: 1 })], novas: 0 }, { g: 500 }));

/* ------------------------------- o id ------------------------------------- */
ok('12. mesmo conjunto de demandas = mesmo id (em qualquer ordem)',
   idDoPlano(['a', 'b'], 1) === idDoPlano(['b', 'a'], 1));
ok('12b. outra parte = outro id', idDoPlano(['a', 'b'], 1) !== idDoPlano(['a', 'b'], 2));

/* ------------------------------ montar ------------------------------------ */
const des = { id: 'd1', codigo: 'CM001', modeloId: 'm-cam', corPrincipalId: 'c-pre', corSecundariaId: 'c-bra',
  componentes: [], aviamentos: [{ materialId: 'av-etq', qtdPorPeca: 1, aplicacao: 'gola' }] };
const grade = { id: 'g1', nome: '1P-2M-2G-1GG | CM.LISA | 117cm', tamanhos: { p: 1, m: 2, g: 2, gg: 1 },
  fases: [{ ordem: 1, nome: 'Corpo', tecidoId: 't-mol', corId: 'x', comp: '2,5', larg: '1,17' },
          { ordem: 2, nome: 'Ribana', tecidoId: 't-rib', corId: 'x', comp: '0,5', larg: '1,17' }] };
const modelo = { id: 'os-velha', os: '0500', desenhoId: 'd1', colecaoId: 'col1', colecaoNome: 'Verão',
  griffeId: 'gr1', griffeNome: 'Diverse', etapas: ['Corte', 'Costura'], obs: 'urgente!',
  tecidos: [{ tecidoId: 't-mol', lote: 'L99' }],
  componentes: [{ nome: 'Bolso', material: 'T:t-mol', materialNome: 'Moletom', qtdPorPeca: 2 }],
  aviamentos: [{ material: 'av-etq', materialNome: 'Etiqueta', qtdPorPeca: 1 }] };
const pl = montarPlano(modelo, des, grade, 36, 3, AGORA);
ok('13. grade e camadas do enfesto cheio', pl.gradeId === 'g1' && pl.enfesto.camadas === 36 && pl.grade.total === 6);
ok('13b. repete a grade (3 OS na aprovacao)', pl.repeticoes === 3);
ok('13c. sem numero', pl.os === '');
ok('13d. cor de cada fase vem do desenho', pl.fases[0].corNome === 'Preto' && pl.fases[1].corNome === 'Branco',
   pl.fases.map(f => f.corNome).join(','));
ok('13e. uma linha de tecido por fase, sem lote', pl.tecidos.length === 2 && pl.tecidos.every(t => t.lote === ''));
ok('13f. contexto da OS-modelo (colecao, marca, etapas)',
   pl.colecaoNome === 'Verão' && pl.griffeNome === 'Diverse' && pl.etapas.join() === 'Corte,Costura');
ok('13g. NADA da OS velha que nao e contexto (obs)', pl.obs === '' && pl.id !== 'os-velha');
ok('13h. modelo e codigo do desenho', pl.modeloNome === 'Camiseta Básica' && pl.codigo === 'CM001');
ok('13i. componente da OS-modelo recalculado (2 por peca × 2 G × 36 camadas = 144 no G)',
   pl.componentes[0] && pl.componentes[0].qtdPorTamanho.g === 144, JSON.stringify(pl.componentes[0]));
ok('13j. producao puxada e tons padrao', pl.tipoProducao === 'puxada' && pl.progresso.totalTamanhoTons[1] === true);

const plSemModeloDoDesenho = montarPlano(Object.assign({}, modelo, { desenhoId: 'outro' }), des, grade, 36, 1, AGORA);
ok('14. OS-modelo de outro desenho: aviamentos vem do cadastro do desenho',
   plSemModeloDoDesenho.aviamentos[0] && plSemModeloDoDesenho.aviamentos[0].app === 'gola'
   && plSemModeloDoDesenho.aviamentos[0].qtdTotal === 6 * 36, JSON.stringify(plSemModeloDoDesenho.aviamentos[0]));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
