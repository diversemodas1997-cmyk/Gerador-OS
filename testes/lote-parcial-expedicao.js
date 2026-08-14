/* Rode com:  node testes/lote-parcial-expedicao.js

   LOTE PARCIAL: para onde vão as peças quando a OS é alocada no planejamento
   de expedição.

   A regra mudou em 14/08/2026. Antes, alocar uma carga de ida tirava os pacotes
   do Estoque de corte e punha em Costurando; agora põe direto em EXPEDIÇÃO —
   alocar no plano já é dizer que aquele pacote vai embarcar. E marcar a etapa
   Costura no checklist traz o lote inteiro de volta para Costurando.

   O teste existe porque esta conta não aparece na tela como conta: aparece como
   saldo. Um erro aqui não dá erro nenhum — dá pano que o programa jura estar
   num campo e está em outro, e some peça da fábrica sem ninguém ver.

   O teste recorta as funções do app.js de verdade. Só _expPecasPacoteOS entra
   dublada: ela puxa a folha de OS inteira (totais por tamanho × tom, vagas da
   grade) e aqui o que importa é a divisão do lote em pacotes, não como ela é
   calculada. As regras de ida/volta/cancelada/carga antiga são as reais. */
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
  cortaArr('const FASES_ESTOQUE'),
  corta('function _faseEntrouOS'),
  corta('function _nomeEtapaDaFase'),
  cortaLinha('function _faseIdxPorId'),
  corta('function _fracAlocadaExpedicaoOS'),
  corta('function calcularSaldosFase'),
  cortaLinha('const TERMINAL_ETAPA_RE'),
  corta('function faseAtualOS'),
  corta('function _expCancelSet'),
  corta('function _expEmbarcadoOS')
].join('\n');

// Saldo (em peças) de cada campo, com o STATE dado.
function saldos(estado) {
  const fn = new Function('STATE', `
    // A cor dos componentes ja vem no formato composto neste teste.
    const corCanonicaPorTecido = (cor) => cor || '';
    // Dublê: 4 vagas de tamanho (P, M, G, GG) em 1 tonalidade, 50 pç cada.
    // Total 200 pç — o mesmo total dos componentes da OS do teste.
    function _expPecasPacoteOS() {
      const mapa = new Map([['P|-', 50], ['M|-', 50], ['G|-', 50], ['GG|-', 50]]);
      return { mapa, total: 200, de: p => mapa.get(p.tam + '|' + (p.tom == null ? '-' : p.tom)) || 0 };
    }
    ${motor}
    const soma = id => {
      const i = FASES_ESTOQUE.findIndex(f => f.id === id);
      return calcularSaldosFase(i).detalhe.reduce((s, c) => s + c.estoque, 0);
    };
    const listaOS = id => {
      const i = FASES_ESTOQUE.findIndex(f => f.id === id);
      return calcularSaldosFase(i).detalhe.flatMap(c => c.osList);
    };
    return {
      corte: soma('corte'), costurando: soma('costurando'),
      fios: soma('fios'), expedicao: soma('expedicao'),
      osExpedicao: listaOS('expedicao')
    };
  `);
  return fn(estado);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const confere = (nome, got, esperado) => {
  const bate = ['corte', 'costurando', 'fios', 'expedicao']
    .every(k => (esperado[k] || 0) === got[k]);
  ok(nome, bate, got);
};

// Uma OS de 200 peças, todas do mesmo tecido+cor, com o checklist da fábrica.
// etapasSeq é o carimbo de QUANDO cada etapa foi marcada: é ele que decide a
// fase atual no modelo sobreposto.
const osBase = (check, seq) => ({
  id: 'os_1', os: '0501', modeloNome: 'Camiseta', data: '2026-08-14',
  gradeId: 'g1',
  etapas: ['Corte', 'Costura', 'Retirada de fios', 'Ensaque', 'Expedição', 'Estoque'],
  progresso: { etapasCheck: check, etapasSeq: seq },
  componentes: [
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdTotal: 120 },
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdTotal: 80 }
  ]
});

const estado = (os, cargas, excecoes) => ({
  ordens: [os],
  expedicaoCargas: cargas || [],
  expedicaoExcecoes: excecoes || [],
  corteMov: [], costurandoMov: [], fiosMov: [], expedicaoMov: []
});

const noCorte = () => osBase({ 'Corte': true }, { 'Corte': 1 });
const cargaIda = (extra) => Object.assign({
  id: 'c1', osId: 'os_1', janelaId: 'j1', data: '2026-08-20', perna: 'ida',
  pacotes: [{ tam: 'P', tom: null }, { tam: 'M', tom: null }], volumes: 5
}, extra || {});

/* ---------- 1. o caminho normal ---------- */

confere('OS no corte, sem nada alocado: as 200 pç ficam no corte',
  saldos(estado(noCorte(), [])),
  { corte: 200 });

confere('alocada METADE numa carga de ida: 100 pç vão para Expedição',
  saldos(estado(noCorte(), [cargaIda()])),
  { corte: 100, expedicao: 100 });

confere('alocada por INTEIRO: as 200 pç vão para Expedição',
  saldos(estado(noCorte(), [cargaIda({ pacotes: [
    { tam: 'P', tom: null }, { tam: 'M', tom: null },
    { tam: 'G', tom: null }, { tam: 'GG', tom: null }] })])),
  { corte: 0, expedicao: 200 });

const so = saldos(estado(noCorte(), [cargaIda()]));
ok('a OS aparece na coluna OS da Expedição mesmo sem a etapa marcada',
  so.osExpedicao.includes('0501'), so.osExpedicao);

/* ---------- 2. a etapa Costura desfaz a alocação ---------- */

confere('marcada a etapa Costura, o lote INTEIRO volta da Expedição para Costurando',
  saldos(estado(osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }),
    [cargaIda()])),
  { corte: 0, costurando: 200 });

confere('etapa Retirada de fios: o lote inteiro está lá, a alocação não conta',
  saldos(estado(osBase({ 'Corte': true, 'Costura': true, 'Retirada de fios': true },
    { 'Corte': 1, 'Costura': 2, 'Retirada de fios': 3 }), [cargaIda()])),
  { fios: 200 });

confere('etapa Expedição marcada: 200 pç lá, sem somar a fração alocada por cima',
  saldos(estado(osBase({ 'Corte': true, 'Expedição': true }, { 'Corte': 1, 'Expedição': 4 }),
    [cargaIda()])),
  { expedicao: 200 });

// O caso torto: a caixa Expedição está marcada, mas o Corte foi marcado DEPOIS
// (a OS voltou para o corte). Sem a guarda, o corte perderia as 100 pç para uma
// Expedição que já as tinha contado e devolvido — 100 peças sumiam da fábrica.
confere('Expedição marcada e Corte remarcado por cima: nada some, 200 pç no corte',
  saldos(estado(osBase({ 'Corte': true, 'Expedição': true }, { 'Corte': 5, 'Expedição': 4 }),
    [cargaIda()])),
  { corte: 200 });

confere('etapa Estoque (terminal): a OS sai de todos os campos',
  saldos(estado(osBase({ 'Corte': true, 'Estoque': true }, { 'Corte': 1, 'Estoque': 9 }),
    [cargaIda()])),
  {});

/* ---------- 3. quais cargas movem peça ---------- */

confere('carga de VOLTA não move peça (é o mesmo pacote voltando)',
  saldos(estado(noCorte(), [cargaIda({ perna: 'volta' })])),
  { corte: 200 });

confere('carga em data CANCELADA não move peça',
  saldos(estado(noCorte(), [cargaIda()],
    [{ janelaId: 'j1', data: '2026-08-20', tipo: 'cancelada' }])),
  { corte: 200 });

confere('carga remarcada (não cancelada) move normalmente',
  saldos(estado(noCorte(), [cargaIda()],
    [{ janelaId: 'j1', data: '2026-08-20', tipo: 'remarcada', novaData: '2026-08-27' }])),
  { corte: 100, expedicao: 100 });

confere('carga ANTIGA (só volumes, sem pacotes) leva o lote inteiro',
  saldos(estado(noCorte(), [{ id: 'c9', osId: 'os_1', janelaId: 'j1', data: '2026-08-20',
    perna: 'ida', volumes: 9 }])),
  { expedicao: 200 });

confere('duas cargas de ida somam os pacotes de cada uma',
  saldos(estado(noCorte(), [
    cargaIda(),
    cargaIda({ id: 'c2', data: '2026-08-27', pacotes: [{ tam: 'G', tom: null }] })
  ])),
  { corte: 50, expedicao: 150 });

confere('carga de outra OS não mexe nesta',
  saldos(estado(noCorte(), [cargaIda({ id: 'c3', osId: 'os_outra' })])),
  { corte: 200 });

/* ---------- 4. contagem manual continua ajustando ---------- */

const comAjuste = estado(noCorte(), [cargaIda()]);
comAjuste.expedicaoMov = [{ id: 'm1', tipo: 'saida', qtd: 30,
  tecidoNome: 'Malha Algodão', corNome: 'Preto Malha Algodão' }];
confere('lançamento manual da Expedição desconta do saldo alocado',
  saldos(comAjuste),
  { corte: 100, expedicao: 70 });

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
