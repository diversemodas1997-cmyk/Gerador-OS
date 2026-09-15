/* Rode com:  node testes/dashboard-fluxo.js

   O DASHBOARD DO FLUXO, na tela de Início: quanto tem em cada campo por onde o
   produto passa, do corte ao estoque.

   O painel não pode inventar conta própria — ele tem que dizer o MESMO que as
   telas de cada campo, senão vira um segundo número para a mesma pergunta e a
   fábrica passa a ter duas verdades. Por isso o teste recorta do app.js as
   funções reais do modelo sobreposto (faseAtualOS, _transitoDaOS) e confere o
   resumo contra elas.

   O que é só do painel, e por isso é o miolo deste teste:
     - o TRÂNSITO se divide em manhã e tarde pela hora cadastrada na janela de
       expedição (a da PERNA: a ida tem a sua hora, a volta a dela), e a soma
       dos dois turnos tem que fechar com o total da perna;
     - "Recebido em São Carlos" repete, de propósito, o Estoque de corte de lá;
     - "Recebido em Descalvado" é a fatia da Retirada de fios que chegou pela
       caixa de recebimento, e não pela retirada em si.

   Como no teste do lote parcial, só _expPecasPacoteOS entra dublada: ela puxa a
   folha de OS inteira, e aqui o que importa é a divisão do lote em pacotes. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF, e a quebra
// depois do fecha-chaves e '\r\n'.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = [
  corta('function _normNome'),
  corta('function osEtapaMarcada'),
  corta('function componentesPorTecidoCorOS'),
  recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades'),
  cortaArr('const FASES_ESTOQUE'),
  corta('function _faseEntrouOS'),
  corta('function _nomeEtapaDaFase'),
  cortaLinha('function _faseIdxPorId'),
  cortaArr('const _TRANSITO_PERNAS'),
  corta('function _transitoDaOS'),
  cortaLinha('const TERMINAL_ETAPA_RE'),
  corta('function faseAtualOS'),
  corta('function _expCancelSet'),
  corta('function _expEmbarcadoOS'),
  corta('function _dashTurnoDaCarga'),
  corta('function _dashPesoTurnos'),
  corta('function _dashFluxoDados')
].join('\n');

// O resumo do painel com o STATE dado.
function dash(estado) {
  const fn = new Function('STATE', `
    const corCanonicaPorTecido = (cor) => cor || '';
    // Dublê: 4 vagas de tamanho (P, M, G, GG) em 1 tonalidade, 50 pç cada.
    // Total 200 pç — o mesmo total dos componentes da OS do teste.
    function _expPecasPacoteOS() {
      const mapa = new Map([['P|-', 50], ['M|-', 50], ['G|-', 50], ['GG|-', 50]]);
      return { mapa, total: 200, de: p => mapa.get(p.tam + '|' + (p.tom == null ? '-' : p.tom)) || 0 };
    }
    ${motor}
    return _dashFluxoDados();
  `);
  return fn(estado);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const CAMPOS = ['corte', 'corteSC', 'costurando', 'costurandoSC',
  'idaManha', 'idaTarde', 'voltaManha', 'voltaTarde',
  'recDesc', 'recSC', 'fios', 'estoque'];

// Confere os DOZE cartões, e não só os citados: cartão esquecido de fora do
// esperado tem que dar falha, senão peça que vazou para o campo errado passa.
const confere = (nome, got, esperado) => {
  const bate = CAMPOS.every(k => (esperado[k] || 0) === got[k].pecas);
  ok(nome, bate, Object.fromEntries(CAMPOS.filter(k => got[k].pecas).map(k => [k, got[k].pecas])));
};

/* ---------------------- o mundo do teste ---------------------- */

// Uma OS de 200 peças, todas do mesmo tecido+cor, com o checklist da fábrica.
// etapasSeq é o carimbo de QUANDO cada etapa foi marcada: é ele que decide a
// fase atual no modelo sobreposto.
const osBase = (check, seq) => ({
  id: 'os_1', os: '0501', modeloNome: 'Camiseta', data: '2026-09-14',
  gradeId: 'g1',
  etapas: ['Corte', 'Recebido em São Carlos', 'Costura', 'Recebido em Descalvado',
           'Retirada de fios', 'Ensaque', 'Expedição', 'Estoque'],
  progresso: { etapasCheck: check, etapasSeq: seq },
  componentes: [
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdTotal: 120 },
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdTotal: 80 }
  ]
});

// Uma janela que sai de manhã e volta à tarde — o desenho normal da viagem
// entre as duas unidades.
const JANELA_MANHA = { id: 'j1', nome: 'Terça', horaIda: '08:00', horaVolta: '17:00' };
// E uma que faz a ida depois do almoço.
const JANELA_TARDE = { id: 'j2', nome: 'Quinta', horaIda: '14:00', horaVolta: '19:00' };

const estado = (os, cargas, janelas, excecoes) => ({
  ordens: [os],
  expedicaoCargas: cargas || [],
  expedicaoJanelas: janelas || [JANELA_MANHA, JANELA_TARDE],
  expedicaoExcecoes: excecoes || [],
  corteMov: [], costurandoMov: [], corteScMov: [], costurandoScMov: [],
  fiosMov: [], expedicaoMov: []
});

const noCorte = () => osBase({ 'Corte': true }, { 'Corte': 1 });
const carga = (extra) => Object.assign({
  id: 'c1', osId: 'os_1', janelaId: 'j1', data: '2026-09-22', perna: 'ida',
  pacotes: [{ tam: 'P', tom: null }, { tam: 'M', tom: null }], volumes: 5
}, extra || {});

/* ---------- 1. o caminho, campo a campo ---------- */

confere('OS no corte: as 200 pç no Estoque de corte · Descalvado',
  dash(estado(noCorte(), [])),
  { corte: 200 });

confere('Costura marcada, sem passar por São Carlos: Costurando · Descalvado',
  dash(estado(osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }), [])),
  { costurando: 200 });

// "Recebido em São Carlos" É o Estoque de corte de lá — o cartão de Recebido
// repete o número de propósito, e é isso que o teste trava.
confere('Recebido em São Carlos: conta no corte de lá E no cartão de recebido',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2 }), [])),
  { corteSC: 200, recSC: 200 });

confere('Costura depois de São Carlos: Costurando · São Carlos',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura': 3 }), [])),
  { costurandoSC: 200 });

// A volta cai na Retirada de fios, e o cartão de Recebido em Descalvado é a
// fatia dela que chegou pela caixa de recebimento.
confere('Recebido em Descalvado: Retirada de fios E o cartão de recebido',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura': true, 'Recebido em Descalvado': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura': 3, 'Recebido em Descalvado': 4 }), [])),
  { fios: 200, recDesc: 200 });

confere('Retirada de fios marcada por último: fios sim, recebido não',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura': true, 'Recebido em Descalvado': true, 'Retirada de fios': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura': 3, 'Recebido em Descalvado': 4, 'Retirada de fios': 5 }), [])),
  { fios: 200 });

confere('Estoque marcado: sai do fluxo em processo e vai para o cartão final',
  dash(estado(osBase({ 'Corte': true, 'Retirada de fios': true, 'Estoque': true },
                     { 'Corte': 1, 'Retirada de fios': 2, 'Estoque': 3 }), [])),
  { estoque: 200 });

// Expedição não tem cartão neste painel (não está na lista pedida): a OS lá não
// pode vazar para nenhum outro cartão.
confere('OS em Expedição: não aparece em cartão nenhum',
  dash(estado(osBase({ 'Corte': true, 'Expedição': true }, { 'Corte': 1, 'Expedição': 2 }), [])),
  {});

// O cartão Estoque conta a CAIXA, não a última etapa: marcada uma vez, a OS
// conta ali mesmo que o fluxo tenha continuado depois. É a exceção pedida em
// 15/09/2026, e por isso a peça aparece nos dois cartões.
confere('Estoque marcado e o Corte marcado DEPOIS: conta nos dois cartões',
  dash(estado(osBase({ 'Estoque': true, 'Corte': true },
                     { 'Estoque': 1, 'Corte': 2 }), [])),
  { corte: 200, estoque: 200 });

confere('OS sem etapa nenhuma marcada: não conta em cartão nenhum',
  dash(estado(osBase({}, {}), [])),
  {});

/* ---------- 2. o trânsito, turno por turno ---------- */

confere('Ida na janela das 8h: 100 pç em Ida · manhã, 100 ficam no corte',
  dash(estado(noCorte(), [carga()])),
  { corte: 100, idaManha: 100 });

confere('Ida na janela das 14h: o mesmo lote cai em Ida · tarde',
  dash(estado(noCorte(), [carga({ janelaId: 'j2' })])),
  { corte: 100, idaTarde: 100 });

confere('Duas cargas de ida, uma de manhã e outra à tarde: 50 em cada turno',
  dash(estado(noCorte(), [
    carga({ id: 'c1', pacotes: [{ tam: 'P', tom: null }] }),
    carga({ id: 'c2', janelaId: 'j2', pacotes: [{ tam: 'M', tom: null }] })
  ])),
  { corte: 100, idaManha: 50, idaTarde: 50 });

// A hora da EXCEÇÃO manda: expedição remarcada para a tarde é caminhão da tarde.
confere('Ocorrência remarcada com hora nova: o turno segue a exceção',
  dash(estado(noCorte(), [carga()], undefined,
    [{ janelaId: 'j1', data: '2026-09-22', tipo: 'remarcada', novaData: '2026-09-23', horaIda: '15:30' }])),
  { corte: 100, idaTarde: 100 });

confere('Expedição cancelada: nada viaja, o lote inteiro fica no corte',
  dash(estado(noCorte(), [carga()], undefined,
    [{ janelaId: 'j1', data: '2026-09-22', tipo: 'cancelada' }])),
  { corte: 200 });

// A volta usa a hora da VOLTA da janela (17h), não a da ida (8h).
confere('Volta da janela das 17h: cai em Volta · tarde, saindo de São Carlos',
  dash(estado(
    osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura': true },
           { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura': 3 }),
    [carga({ perna: 'volta' })])),
  { costurandoSC: 100, voltaTarde: 100 });

// Carga cheia antiga (só o número de volumes, sem composição por pacote): o
// lote inteiro viaja, e o campo de origem zera.
confere('Carga antiga sem pacotes: as 200 pç inteiras no turno da janela',
  dash(estado(noCorte(), [carga({ pacotes: undefined, volumes: 5 })])),
  { idaManha: 200 });

/* ---------- 3. o que o painel promete: a soma fecha ---------- */

const d = dash(estado(noCorte(), [
  carga({ id: 'c1', pacotes: [{ tam: 'P', tom: null }] }),
  carga({ id: 'c2', janelaId: 'j2', pacotes: [{ tam: 'M', tom: null }, { tam: 'G', tom: null }] })
]));
const emProcesso = CAMPOS.filter(k => k !== 'recDesc' && k !== 'recSC')
  .reduce((s, k) => s + d[k].pecas, 0);
ok('a soma dos cartões (sem os de Recebido, que repetem) é o lote inteiro',
  emProcesso === 200, emProcesso);
ok('a OS em dois turnos conta uma vez em cada cartão de turno',
  d.idaManha.os === 1 && d.idaTarde.os === 1 && d.corte.os === 1,
  { idaManha: d.idaManha.os, idaTarde: d.idaTarde.os, corte: d.corte.os });

// Campo vazio é zero, e não some da leitura: o painel desenha os doze cartões
// sempre, e um `undefined` aqui viraria "—" na tela de quem confere a produção.
ok('todo cartão existe mesmo vazio, com peças e OS em zero',
  CAMPOS.every(k => d[k] && typeof d[k].pecas === 'number' && typeof d[k].os === 'number'),
  Object.keys(d));

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
