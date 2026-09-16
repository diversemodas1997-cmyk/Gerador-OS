/* Rode com:  node testes/dashboard-historico.js

   O HISTÓRICO DE CADA CARTÃO do Início (16/09/2026): entrada e saída por semana,
   o que está agora, o residual (parado há mais de 7 dias), o total do período,
   o "desde quando" de cada OS e o tempo médio no quadro.

   O programa não guarda um diário de "a OS mudou de campo": o passado é
   RECONSTRUÍDO pelas horas em que cada etapa foi marcada (etapasSeq). Este
   teste guarda o que a reconstrução promete:

     · a entrada e a saída caem na semana certa;
     · a OS parada há mais de 7 dias conta como residual;
     · o tempo médio no quadro é a média do que já saiu;
     · a OS sem hora de verdade (carimbo sintético da migração: 1, 2, 3) conta
       onde está, mas fica SEM DATA, em vez de inventar uma;
     · o fim do passado reconstruído bate com o cartão de agora.

   Recorta do app.js as funções reais, como o teste do painel. */
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
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const motor = [
  corta('function _normNome'),
  corta('function osEtapaMarcada'),
  corta('function _produtosDosComponentes'),
  corta('function produtosOS'),
  recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades'),
  cortaArr('const FASES_ESTOQUE'),
  cortaArr('const STATUS_OS'),
  cortaLinha('const STATUS_FIM'),
  corta('function _marcasDoStatus'),
  corta('function _statusDoChecklistOS'),
  corta('function _ultimaMarcacaoChecklist'),
  corta('function _statusOS'),
  corta('function _faseEntrouOS'),
  corta('function _nomeEtapaDaFase'),
  cortaLinha('function _faseIdxPorId'),
  cortaArr('const _TRANSITO_PERNAS'),
  corta('function _fracoesMovidasOS'),
  corta('function _transitoDaOS'),
  corta('function _fracaoRecebida'),
  corta('function _fracaoPerdida'),
  corta('function _expIso'),
  corta('function _expHoje'),
  corta('function _expDataEfetivaCarga'),
  cortaLinha('const TERMINAL_ETAPA_RE'),
  corta('function faseAtualOS'),
  corta('function _expCancelSet'),
  corta('function _expEmbarcadoOS'),
  corta('function _dashTurnoDaCarga'),
  corta('function _dashPesoTurnos'),
  corta('function _dashCartoesDaOS'),
  corta('function _dashChavePorIdx'),
  corta('function _dashFluxoDados'),
  recorte('const DASH_RESIDUAL_DIAS', '// Os seis passos do caminho', 'o historico do painel')
].join('\n');

const DIA = 86400000;
const AGORA = Date.UTC(2026, 8, 16, 12, 0, 0);

function rodar(ordens) {
  const fn = new Function('STATE', 'AGORA', `
    const corCanonicaPorTecido = (cor) => cor || '';
    const totaisPorTamanhoTomOS = () => ({ totalGeral: 0 });
    function _expPecasPacoteOS() { return { mapa: new Map(), total: 200, de: () => 0 }; }
    ${motor}
    const d = _dashFluxoDados();
    return { d, h: _dashHistorico(d, AGORA), linha: STATE.ordens.map(_dashLinhaDoTempoOS) };
  `);
  return fn({
    ordens, expedicaoCargas: [], expedicaoJanelas: [], expedicaoExcecoes: [],
    corteMov: [], costurandoMov: [], corteScMov: [], costurandoScMov: [], fiosMov: [], expedicaoMov: []
  }, AGORA);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// Uma OS de 200 produtos; `marcas` = {etapa: há quantos dias foi marcada}.
const os = (num, marcas, seqCru) => {
  const check = {}, seq = {};
  Object.keys(marcas).forEach(n => { check[n] = true; seq[n] = seqCru ? marcas[n] : AGORA - marcas[n] * DIA; });
  return {
    id: 'id_' + num, os: num, modeloNome: 'Camiseta', data: '2026-08-20',
    etapas: ['Corte', 'Ensaque', 'Costura', 'Retirada de fios', 'Estoque'],
    progresso: { etapasCheck: check, etapasSeq: seq },
    componentes: [{ materialNome: 'Malha', corNome: 'Preto', qtdPorPeca: 1, qtdTotal: 200 }]
  };
};

const r = rodar([
  // A: cortada há 20 dias, ensacada há 10 — parada no estoque de corte (residual)
  os('0001', { 'Corte': 20, 'Ensaque': 10 }),
  // B: cortada há 9, ensacada há 8, foi para a costura há 3
  os('0002', { 'Corte': 9, 'Ensaque': 8, 'Costura': 3 }),
  // C: OS antiga, só com a ORDEM das etapas (1, 2, 3 milissegundos)
  os('0003', { 'Corte': 1, 'Ensaque': 2, 'Costura': 3 }, true)
]);
const { d, h } = r;

console.log('-- estoque de corte --');
ok('agora: só a OS 0001 está lá (200)', h.corte.agora === 200, h.corte.agora);
ok('entraram 0001 (há 10 dias) e 0002 (há 8): 400', h.corte.entrada === 400, h.corte.entrada);
ok('saiu a 0002 (há 3 dias): 200', h.corte.saida === 200, h.corte.saida);
ok('as entradas caem na 3ª semana, a saída na 4ª',
   h.corte.semanas.map(w => w.entrada + '/' + w.saida).join(' ') === '0/0 0/0 400/0 0/200',
   h.corte.semanas.map(w => w.entrada + '/' + w.saida));
ok('a 0001, parada há 10 dias, é residual', h.corte.residual === 200, h.corte.residual);
ok('total do período: nada estava lá há 4 semanas + 400 que entraram', h.corte.total === 400, h.corte.total);
ok('tempo médio no quadro: a 0002 ficou 5 dias', Math.abs(h.corte.tempoMedio - 5) < 1e-9, h.corte.tempoMedio);
ok('a mais antiga é a 0001, desde há 10 dias',
   h.corte.maisAntiga && h.corte.maisAntiga.os === '0001' && h.corte.maisAntiga.desde === AGORA - 10 * DIA, h.corte.maisAntiga);
ok('a última entrada foi há 8 dias', h.corte.ultimaEntrada === AGORA - 8 * DIA, h.corte.ultimaEntrada);
ok('a idade cai na faixa "8 a 14 dias"', h.corte.faixas[2].v === 200 && h.corte.faixas[0].v === 0, h.corte.faixas);

console.log('');
console.log('-- costurando, com uma OS sem data --');
ok('agora: 0002 e 0003 (400)', h.costurando.agora === 400, h.costurando.agora);
ok('só a 0002 tem data de entrada: 200 na 4ª semana', h.costurando.entrada === 200
   && h.costurando.semanas[3].entrada === 200, h.costurando.semanas);
ok('a 0003 conta, mas SEM DATA', h.costurando.semData === 200 && h.costurando.faixas[4].v === 200, h.costurando.faixas);
ok('e não vira residual inventado', h.costurando.residual === 0, h.costurando.residual);
ok('a OS sem hora de verdade não tem linha do tempo', r.linha[2] === null, r.linha[2]);
ok('e entra no total como "já estava"', h.costurando.total === 400, h.costurando.total);

console.log('');
console.log('-- o passado bate com o agora --');
const fim = l => l ? [...l[l.length - 1].cartoes.keys()].sort().join(',') : null;
ok('o último estado reconstruído da 0001 é o cartão de hoje', fim(r.linha[0]) === 'corte', fim(r.linha[0]));
ok('e o da 0002 também', fim(r.linha[1]) === 'costurando', fim(r.linha[1]));
ok('os números de agora são os mesmos do cartão',
   h.corte.agora === d.corte.pecas && h.costurando.agora === d.costurando.pecas, [d.corte.pecas, d.costurando.pecas]);
ok('o trânsito não tem histórico', h.idaManha.semHistorico === true && h.corte.semHistorico === false, '');

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
