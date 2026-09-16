/* Rode com:  node testes/dashboard-historico.js

   O HISTÓRICO DE CADA CARTÃO do Início (16/09/2026): entrada e saída por semana,
   o residual (o que está no quadro, no fim de cada período e agora), o total do período,
   o "desde quando" de cada OS e o tempo médio no quadro.

   O programa não guarda um diário de "a OS mudou de campo": o passado é
   RECONSTRUÍDO pelas horas em que cada etapa foi marcada (etapasSeq). Este
   teste guarda o que a reconstrução promete:

     · a entrada e a saída caem na semana certa — de SEGUNDA 00:00 a SEXTA 23:59,
       e o que acontece no sábado ou no domingo fica fora e é contado à parte;
     · o tempo médio no quadro é a média do que já saiu;
     · a OS sem hora de verdade (carimbo sintético da migração: 1, 2, 3) conta
       onde está, mas fica SEM DATA, em vez de inventar uma;
     · o volume no quadro ao fim de cada período, e o do período em curso é o agora;
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
  recorte('const DASH_DIA_MS', '// Os seis passos do caminho', 'o historico do painel')
].join('\n');

const DIA = 86400000;
// Datas em hora LOCAL, como a fábrica vive: (dia de setembro/2026, hora).
// A semana do painel vai de segunda 00:00 a sexta 23:59.
const em = (dia, hora, min) => new Date(2026, 8, dia, hora || 0, min || 0).getTime();
const emAgo = (dia, hora) => new Date(2026, 7, dia, hora || 0).getTime();
// Quarta-feira, 16/09/2026, 12:00. As 4 semanas: 24–28/08, 31/08–04/09,
// 07–11/09 e 14–18/09 (a atual, em curso).
const AGORA = em(16, 12);

function rodar(ordens, escala) {
  const fn = new Function('STATE', 'AGORA', 'ESCALA', `
    const window = {};
    const corCanonicaPorTecido = (cor) => cor || '';
    const totaisPorTamanhoTomOS = () => ({ totalGeral: 0 });
    function _expPecasPacoteOS() { return { mapa: new Map(), total: 200, de: () => 0 }; }
    ${motor}
    const d = _dashFluxoDados();
    return { d, h: _dashHistorico(d, AGORA, ESCALA), linha: STATE.ordens.map(_dashLinhaDoTempoOS) };
  `);
  return fn({
    ordens, expedicaoCargas: [], expedicaoJanelas: [], expedicaoExcecoes: [],
    corteMov: [], costurandoMov: [], corteScMov: [], costurandoScMov: [], fiosMov: [], expedicaoMov: []
  }, AGORA, escala);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// Uma OS de 200 produtos; `marcas` = {etapa: instante em ms}.
const os = (num, marcas) => {
  const check = {}, seq = {};
  Object.keys(marcas).forEach(n => { check[n] = true; seq[n] = marcas[n]; });
  return {
    id: 'id_' + num, os: num, modeloNome: 'Camiseta', data: '2026-08-20',
    etapas: ['Corte', 'Ensaque', 'Costura', 'Retirada de fios', 'Estoque'],
    progresso: { etapasCheck: check, etapasSeq: seq },
    componentes: [{ materialNome: 'Malha', corNome: 'Preto', qtdPorPeca: 1, qtdTotal: 200 }]
  };
};

const r = rodar([
  // A: cortada qua 26/08, ensacada seg 07/09 08:00 — segue no estoque de corte
  os('0001', { 'Corte': emAgo(26, 10), 'Ensaque': em(7, 8) }),
  // B: cortada ter 08/09, ensacada qua 09/09 10:00, costura seg 14/09 15:00
  os('0002', { 'Corte': em(8, 9), 'Ensaque': em(9, 10), 'Costura': em(14, 15) }),
  // C: OS antiga, só com a ORDEM das etapas (1, 2, 3 milissegundos)
  os('0003', { 'Corte': 1, 'Ensaque': 2, 'Costura': 3 })
]);
const { d, h } = r;

console.log('-- a semana é de segunda a sexta --');
ok('4 semanas, começando numa segunda 00:00',
   h.corte.periodos.every(w => new Date(w.de).getDay() === 1 && new Date(w.de).getHours() === 0), h.corte.periodos.map(w => new Date(w.de).toString()));
ok('e terminando no sábado 00:00 (a sexta inteira conta)',
   h.corte.periodos.every(w => new Date(w.ate).getDay() === 6 && w.ate - w.de === 5 * DIA), '');
ok('a 1ª semana começa em 24/08 e a última em 14/09 (a atual)',
   h.corte.periodos[0].de === emAgo(24) && h.corte.periodos[3].de === em(14), h.corte.periodos.map(w => new Date(w.de).toLocaleDateString('pt-BR')));

console.log('');
console.log('-- estoque de corte --');
ok('agora: só a OS 0001 está lá (200)', h.corte.agora === 200, h.corte.agora);
ok('entraram 0001 (seg 07/09) e 0002 (qua 09/09): 400', h.corte.entrada === 400, h.corte.entrada);
ok('saiu a 0002 (seg 14/09): 200', h.corte.saida === 200, h.corte.saida);
ok('as entradas caem na 3ª semana, a saída na 4ª',
   h.corte.periodos.map(w => w.entrada + '/' + w.saida).join(' ') === '0/0 0/0 400/0 0/200',
   h.corte.periodos.map(w => w.entrada + '/' + w.saida));
ok('total do período: nada estava lá em 24/08 + 400 que entraram', h.corte.total === 400, h.corte.total);
ok('tempo médio no quadro: a 0002 ficou 5 dias e 5 horas',
   Math.abs(h.corte.tempoMedio - (5 + 5 / 24)) < 1e-9, h.corte.tempoMedio);
ok('a mais antiga é a 0001, desde 07/09 08:00',
   h.corte.maisAntiga && h.corte.maisAntiga.os === '0001' && h.corte.maisAntiga.desde === em(7, 8), h.corte.maisAntiga);
ok('a última entrada foi qua 09/09 10:00', h.corte.ultimaEntrada === em(9, 10), h.corte.ultimaEntrada);
ok('a idade cai na faixa "8 a 14 dias"', h.corte.faixas[2].v === 200 && h.corte.faixas[0].v === 0, h.corte.faixas);
ok('nada aconteceu em fim de semana', h.corte.foraDoPeriodo === 0, h.corte.foraDoPeriodo);
ok('residual no fim de cada semana: 0, 0, 400 (sex 11/09) e 200 na semana em curso',
   h.corte.periodos.map(w => w.residual).join(' ') === '0 0 400 200', h.corte.periodos.map(w => w.residual));
ok('e o residual da semana em curso é o agora: o mesmo número do cartão', h.corte.periodos[3].residual === h.corte.agora, [h.corte.periodos[3].residual, h.corte.agora]);
ok('não existe mais residual separado do que está no quadro', !('residual' in h.corte) && !('estoque' in h.corte.periodos[0]), Object.keys(h.corte));
console.log('');
console.log('-- costurando, com uma OS sem data --');
ok('agora: 0002 e 0003 (400)', h.costurando.agora === 400, h.costurando.agora);
ok('só a 0002 tem data de entrada: 200 na 4ª semana', h.costurando.entrada === 200
   && h.costurando.periodos[3].entrada === 200, h.costurando.periodos);
ok('a 0003 conta, mas SEM DATA', h.costurando.semData === 200 && h.costurando.faixas[4].v === 200, h.costurando.faixas);
ok('a OS sem data entra no residual: ela está no quadro', h.costurando.periodos[3].residual === 400, h.costurando.periodos.map(w => w.residual));
ok('a OS sem hora de verdade não tem linha do tempo', r.linha[2] === null, r.linha[2]);
ok('e entra no total como "já estava"', h.costurando.total === 400, h.costurando.total);

console.log('');
console.log('-- sábado, domingo e as bordas da semana --');
const r2 = rodar([
  // ensacada no SÁBADO 12/09: fica fora das semanas, e a tela diz quanto
  os('0010', { 'Corte': em(10, 9), 'Ensaque': em(12, 10) }),
  // ensacada na SEGUNDA 14/09 00:00 em ponto: conta na semana atual
  os('0011', { 'Corte': em(11, 9), 'Ensaque': em(14, 0, 0) }),
  // ensacada na SEXTA 11/09 23:59: conta na 3ª semana
  os('0012', { 'Corte': em(10, 9), 'Ensaque': em(11, 23, 59) })
]);
const c2 = r2.h.corte;
ok('a ensacada no sábado não entra em coluna nenhuma',
   c2.periodos.map(w => w.entrada).join(' ') === '0 0 200 200', c2.periodos.map(w => w.entrada));
ok('e é contada como fora da semana (200)', c2.foraDoPeriodo === 200, c2.foraDoPeriodo);
ok('segunda 00:00 é da semana nova; sexta 23:59 é da semana que termina',
   c2.periodos[3].entrada === 200 && c2.periodos[2].entrada === 200, c2.periodos);
ok('o agora não muda por causa do fim de semana: as três estão lá', c2.agora === 600, c2.agora);

console.log('');
console.log('-- o período escolhido por quem olha: dia, mês e ano --');
const massa = [
  os('0001', { 'Corte': emAgo(26, 10), 'Ensaque': em(7, 8) }),
  os('0002', { 'Corte': em(8, 9), 'Ensaque': em(9, 10), 'Costura': em(14, 15) })
];
const porDia = rodar(massa, 'dia').h.corte;
ok('DIA: 10 colunas, todas de segunda a sexta',
   porDia.periodos.length === 10 && porDia.periodos.every(w => [1, 2, 3, 4, 5].includes(new Date(w.de).getDay())),
   porDia.periodos.map(w => w.nome));
ok('DIA: a última é hoje (qua 16/09) e a primeira, qui 03/09',
   porDia.periodos[9].de === em(16) && porDia.periodos[0].de === em(3), porDia.periodos.map(w => w.nome));
ok('DIA: a 0001 entrou na seg 07/09, a 0002 na qua 09/09, a saída na seg 14/09',
   porDia.periodos.find(w => w.de === em(7)).entrada === 200
   && porDia.periodos.find(w => w.de === em(9)).entrada === 200
   && porDia.periodos.find(w => w.de === em(14)).saida === 200, porDia.periodos.map(w => w.rot + ' ' + w.entrada + '/' + w.saida));
const porMes = rodar(massa, 'mes').h.corte;
ok('MÊS: 6 meses, de abr/26 a set/26',
   porMes.periodos.length === 6 && porMes.periodos[0].rot === 'abr/26' && porMes.periodos[5].rot === 'set/26',
   porMes.periodos.map(w => w.rot));
ok('MÊS: as duas entradas e a saída caem em setembro',
   porMes.periodos[5].entrada === 400 && porMes.periodos[5].saida === 200, porMes.periodos.map(w => w.rot + ' ' + w.entrada + '/' + w.saida));
ok('MÊS: no mês não há fim de semana para deixar de fora', porMes.foraDoPeriodo === 0, porMes.foraDoPeriodo);
const sabado = rodar([os('0010', { 'Corte': em(10, 9), 'Ensaque': em(12, 10) })], 'mes').h.corte;
ok('MÊS: a ensacada no sábado CONTA no mês', sabado.periodos[5].entrada === 200 && sabado.foraDoPeriodo === 0, sabado.periodos[5]);
const porAno = rodar(massa, 'ano').h.corte;
ok('ANO: 2024, 2025 e 2026, com tudo em 2026',
   porAno.periodos.map(w => w.rot).join(' ') === '2024 2025 2026' && porAno.periodos[2].entrada === 400,
   porAno.periodos.map(w => w.rot + ' ' + w.entrada));
ok('sem escala escolhida, vale a semana', rodar(massa).h.corte.periodos.length === 4, '');

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
