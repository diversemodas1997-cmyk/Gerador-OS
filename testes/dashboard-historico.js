/* Rode com:  node testes/dashboard-historico.js

   O HISTÓRICO DE CADA CARTÃO do Início (16/09/2026): entrada e saída por semana,
   o residual (o do período anterior + entrada − saída), o total do período,
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
  cortaLinha('const STATUS_FIM_RESERVA'),
  corta('function _dataFinalizacaoOS'),
  corta('function _marcasDoStatus'),
  corta('function _statusDoChecklistOS'),
  corta('function _ultimaMarcacaoChecklist'),
  corta('function _statusOS'),
  // O carimbo do status tambem decide o campo desde 18/09/2026, e _faseEntrouOS
  // passou a perguntar por ele.
  corta('function _faseCarimbadaOS'),
  corta('function _carimboTerminalOS'),
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
  corta('function _expAgora'),
  corta('function _expDataEfetivaCarga'),
  corta('function _expInstanteCarga'),
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
// A barra é das OS que se moveram NA SEMANA EM CURSO (07/10/2026): só a 0002,
// que saiu seg 14/09 depois de 5 dias e 5 horas. A 0001, parada desde 07/09,
// não se moveu na semana e fica fora da barra.
ok('a barra de tempo é só da 0002, que ficou 5 dias (faixa "3 a 7 dias")',
   h.corte.faixas[1].v === 200 && h.corte.faixas.reduce((s, f) => s + f.v, 0) === 200, h.corte.faixas);
ok('nada aconteceu em fim de semana', h.corte.foraDoPeriodo === 0, h.corte.foraDoPeriodo);
ok('residual de cada semana (anterior + entrada − saída): 0, 0, 400 e 200',
   h.corte.periodos.map(w => w.residual).join(' ') === '0 0 400 200', h.corte.periodos.map(w => w.residual));
ok('a conta fecha semana a semana: 400 + 0 − 200 = 200',
   h.corte.periodos.every((w, i) => w.residual === (i ? h.corte.periodos[i - 1].residual : h.corte.residualInicial) + w.entrada - w.saida), h.corte.periodos);
ok('e nunca fica negativo', h.corte.periodos.every(w => w.residual >= 0), '');
ok('o residual final é o que ficou: o que já estava + entrou − saiu', h.corte.residual === h.corte.residualInicial + h.corte.entrada - h.corte.saida && h.corte.residual === 200, h.corte.residual);
ok('não existe mais o "estoque" à parte do residual', !('estoque' in h.corte.periodos[0]), Object.keys(h.corte.periodos[0]));
console.log('');
console.log('-- costurando, com uma OS sem data --');
ok('agora: 0002 e 0003 (400)', h.costurando.agora === 400, h.costurando.agora);
ok('só a 0002 tem data de entrada: 200 na 4ª semana', h.costurando.entrada === 200
   && h.costurando.periodos[3].entrada === 200, h.costurando.periodos);
ok('a 0003 conta no cartão, mas SEM DATA', h.costurando.semData === 200, h.costurando.semData);
ok('a barra de tempo é só da 0002 (entrou seg 14/09, está há menos de 2 dias); a 0003, sem movimento, fica fora',
   h.costurando.faixas[0].v === 200 && h.costurando.faixas[4].v === 0, h.costurando.faixas);
ok('a OS sem data NÃO entra no residual (não tem entrada datada), mas está no cartão',
   h.costurando.periodos[3].residual === 200 && h.costurando.agora === 400, [h.costurando.periodos.map(w => w.residual), h.costurando.agora]);
const residualDia = rodar([os('0020', { 'Corte': em(3, 9), 'Ensaque': em(4, 10) }), os('0021', { 'Corte': em(4, 8) })], 'dia').h.cortando;
ok('DIA: sex 04/09 carrega o residual de qui 03/09 (200 + 200 − 200 = 200)',
   residualDia.periodos.find(w => w.de === em(4)).residual === 200 && residualDia.periodos.find(w => w.de === em(3)).residual === 200,
   residualDia.periodos.map(w => w.rot + ' ' + w.entrada + '-' + w.saida + '=' + w.residual));
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

const comAnterior = rodar([os('0030', { 'Corte': emAgo(20, 9) }), os('0031', { 'Corte': em(15, 9), 'Ensaque': em(16, 8) })], 'semana').h.cortando;
ok('o que já estava antes da 1ª coluna é o residual inicial (0030, cortada em 20/08)', comAnterior.residualInicial === 200, comAnterior.residualInicial);
ok('e ele carrega: 200, 200, 200 e 200 (a 0031 entrou e saiu na semana em curso)',
   comAnterior.periodos.map(w => w.residual).join(' ') === '200 200 200 200', comAnterior.periodos.map(w => w.residual));

console.log('');
console.log('-- o passado bate com o agora --');
const fim = l => l ? [...l[l.length - 1].cartoes.keys()].sort().join(',') : null;
ok('o último estado reconstruído da 0001 é o cartão de hoje', fim(r.linha[0]) === 'corte', fim(r.linha[0]));
ok('e o da 0002 também', fim(r.linha[1]) === 'costurando', fim(r.linha[1]));
ok('os números de agora são os mesmos do cartão',
   h.corte.agora === d.corte.pecas && h.costurando.agora === d.costurando.pecas, [d.corte.pecas, d.costurando.pecas]);
/* O TRÂNSITO TEM HISTÓRICO desde 07/10/2026 (Junior: "os quadros no Início
   devem mostrar o que foi movimentado apenas no dia atual, quando o filtro
   estiver selecionado Dia"). A caixa da viagem tem a hora da carga, e o
   trânsito segue o filtro como os outros. */
ok('o trânsito tem histórico, como os outros quadros', h.idaManha.semHistorico === false && h.corte.semHistorico === false, '');
const osViagem = (num, marcas) => Object.assign(os(num, marcas), {
  etapas: ['Corte', 'Ensaque', 'Expedição Desc X São Carlos', 'Recebido em São Carlos', 'Estoque'] });
const hv = rodar([
  // Saiu de Descalvado ONTEM (ter 15/09 14:00) e segue na estrada.
  osViagem('0070', { 'Corte': em(14, 8), 'Ensaque': em(14, 10), 'Expedição Desc X São Carlos': em(15, 14) }),
  // Saiu HOJE (qua 16/09 08:00).
  osViagem('0071', { 'Corte': em(15, 8), 'Ensaque': em(15, 10), 'Expedição Desc X São Carlos': em(16, 8) })
], 'dia').h;
const naLista = hv.idaManha.listaOS.filter(r => r.entrada > 0 || r.saida > 0).map(r => r.os);
ok('no Dia, o trânsito lista só a OS que entrou na estrada hoje', naLista.join(',') === '0071', hv.idaManha.listaOS);
ok('a que está na estrada desde ontem segue no número do cartão (corrente)',
   (hv.idaManha.listaOS.find(r => r.os === '0070') || {}).corrente === 200, hv.idaManha.listaOS);

console.log('');
console.log('-- a lista de OS do quadro (07/10/2026) --');
// A lista é do período EM CURSO: as colunas somam a última coluna do quadro.
const confere = (nome, x) => {
  const soma = c => x.listaOS.reduce((s, o) => s + o[c], 0);
  const w = x.periodos[x.periodos.length - 1];
  ok(nome + ': entrada, saída e residual da lista = os do período em curso',
     soma('entrada') === w.entrada && soma('saida') === w.saida && soma('residual') === w.residual,
     { lista: ['entrada', 'saida', 'residual'].map(soma), periodo: [w.entrada, w.saida, w.residual] });
  ok(nome + ': corrente da lista = cartão agora', soma('corrente') === x.agora, [soma('corrente'), x.agora]);
};
confere('corte', h.corte);
confere('costurando', h.costurando);
confere('cortando com anterior', comAnterior);
const l2 = h.corte.listaOS.find(o => o.os === '0002');
ok('a 0002 entrou no estoque de corte na semana passada e saiu nesta: entrada 0, saída 200, total 200',
   l2 && l2.entrada === 0 && l2.saida === 200 && l2.corrente === 0 && l2.residual === 0 && l2.total === 200, l2);
// No DIA, a entrada é só a de hoje (qua 16/09): quem entrou ontem não aparece em Entrada.
const hojeDia = rodar([os('0040', { 'Corte': em(15, 9) }), os('0041', { 'Corte': em(16, 9) })], 'dia').h.cortando;
const e40 = hojeDia.listaOS.find(o => o.os === '0040') || {}, e41 = hojeDia.listaOS.find(o => o.os === '0041') || {};
ok('DIA: só a OS cortada hoje tem entrada; a de ontem está em corrente, sem entrada',
   e41.entrada === 200 && e40.entrada === 0 && e40.corrente === 200, hojeDia.listaOS);
ok('a linha da OS sabe o id, para abrir a folha', l2 && l2.id === 'id_0002', l2);
const l3 = h.costurando.listaOS.find(o => o.os === '0003') || {};
ok('a OS antiga sem data aparece no corrente, não no residual', l3.corrente === 200 && l3.residual === 0, l3);

console.log('');
console.log('-- o diário de status: cada carimbo à mão vale desde a hora dele (07/10/2026) --');
// A 0050: Corte marcado seg 14/09 08:00; carimbada à mão Separando ter 15/09
// 09:00 e Ensacado qua 16/09 10:00. O carimbo de agora é o do Ensacado.
const comDiario = Object.assign(os('0050', { 'Corte': em(14, 8) }), {
  statusOS: 'ensacado', statusOSEm: new Date(em(16, 10)).toISOString(),
  // O carimbo Separando grava a segunda data (STATUS_FIM), e o Ensacado depois não a troca.
  finalizadaEm: new Date(em(15, 9)).toISOString(),
  statusHist: [{ k: 'cortando', em: em(14, 8) },
               { k: 'separando', em: em(15, 9), c: 'separando' },
               { k: 'ensacado', em: em(16, 10), c: 'ensacado' }]
});
const semDiario = Object.assign(os('0051', { 'Corte': em(14, 8) }), {
  statusOS: 'ensacado', statusOSEm: new Date(em(16, 10)).toISOString()
});
const hd = rodar([comDiario], 'dia').h;
const hs = rodar([semDiario], 'dia').h;
const ivDe = (hh, k) => (hh[k].listaOS[0] || {});
ok('com o diário, a 0050 saiu da mesa de corte ter 15/09 (o carimbo Separando), não qua 16/09',
   rodar([comDiario], 'semana').h.cortando.periodos[3].saida === 200
   && hd.cortando.listaOS.length === 0, [rodar([comDiario], 'semana').h.cortando.periodos, hd.cortando.listaOS]);
ok('e esteve em Separando de ter 15/09 09:00 a qua 16/09 10:00: no Dia de hoje, saiu',
   ivDe(hd, 'separando').saida === 200 && ivDe(hd, 'separando').entrada === 0, hd.separando.listaOS);
ok('o tempo médio na separação é o de verdade: 1 dia e 1 hora',
   Math.abs(hd.separando.tempoMedio - (1 + 1 / 24)) < 1e-9, hd.separando.tempoMedio);
ok('sem o diário (OS de antes), só o último carimbo é conhecido: a mesa vai até qua 16/09 10:00',
   ivDe(hs, 'cortando').saida === 200 && hs.separando.listaOS.length === 0, [hs.cortando.listaOS, hs.separando.listaOS]);
// O "Não iniciado" à mão apaga o carimbo: dali em diante vale a folha.
const limpa = Object.assign(os('0052', { 'Corte': em(14, 8) }), {
  statusHist: [{ k: 'separando', em: em(15, 9), c: 'separando' }, { k: 'cortando', em: em(16, 9), c: '' }]
});
console.log('');
console.log('-- a saída da mesa de corte é a SEGUNDA DATA (07/10/2026) --');
// A 0060 foi carimbada Ensacado ontem (ter 15/09 17:00: a segunda data). Hoje
// (qua 16/09) alguém pôs o checklist em dia: Corte 09:00 e Ensaque 10:00. Pela
// reconstrução, ela entrava e saía da mesa HOJE.
const emDia = Object.assign(os('0060', { 'Corte': em(16, 9), 'Ensaque': em(16, 10) }),
  { finalizadaEm: new Date(em(15, 17)).toISOString() });
const hsd = rodar([emDia], 'dia').h;
ok('a OS cortada ONTEM (segunda data) não aparece no Dia de hoje da mesa de corte',
   hsd.cortando.listaOS.every(r => !(r.saida > 0)), hsd.cortando.listaOS);
const hss = rodar([emDia], 'semana').h;
ok('na semana, a saída dela é contada uma vez, na terça (a segunda data)',
   hss.cortando.periodos[3].saida === 200 && hss.cortando.listaOS.filter(r => r.saida > 0).length === 1,
   hss.cortando.periodos[3]);
// A 0061 tem a segunda data HOJE: aparece no Dia.
const hoje61 = Object.assign(os('0061', { 'Corte': em(15, 9), 'Ensaque': em(16, 8) }),
  { finalizadaEm: new Date(em(16, 8)).toISOString() });
const h61 = rodar([hoje61], 'dia').h;
ok('a OS com a segunda data de hoje aparece no Dia, com a saída',
   (h61.cortando.listaOS[0] || {}).saida === 200, h61.cortando.listaOS);

// A 0062 foi carimbada Cortando à mão antes do diário existir, e Separando
// hoje às 11:00 por cima: sem a caixa Corte, a reconstrução nunca a vê na mesa.
// A segunda data diz que ela foi cortada hoje.
const soCarimbo = Object.assign(os('0062', {}), { statusOS: 'separando',
  statusOSEm: new Date(em(16, 11)).toISOString(), finalizadaEm: new Date(em(16, 11)).toISOString(),
  statusHist: [{ k: 'separando', em: em(16, 11), c: 'separando' }] });
const h62 = rodar([soCarimbo], 'dia').h;
ok('a OS cortada hoje que a reconstrução nunca viu na mesa sai nela pela segunda data',
   (h62.cortando.listaOS[0] || {}).saida === 200 && (h62.cortando.listaOS[0] || {}).entrada === 0, h62.cortando.listaOS);
ok('e sem hora de entrada, não pesa no residual nem no tempo médio',
   h62.cortando.residual === 0 && h62.cortando.tempoMedio == null, [h62.cortando.residual, h62.cortando.tempoMedio]);

console.log('');
console.log('-- a OS que passou por tudo no dia consta em todos os quadros, com a hora (07/10/2026) --');
// A 0070 andou a manhã inteira de hoje (qua 16/09): matéria-prima 07:00,
// enfesto 08:00, corte 09:00, Separando à mão 10:00, ensaque 10:30, costura
// 11:00. Ontem (15/09) a 0071 entrou e saiu do enfesto: não é do Dia de hoje.
const passou = Object.assign(os('0070', { 'Preparar matéria-prima': em(16, 7), 'Corte': em(16, 9),
  'Ensaque': em(16, 10, 30), 'Costura': em(16, 11) }), {
  statusHist: [{ k: 'separando', em: em(16, 10), c: 'separando' }],
  finalizadaEm: new Date(em(16, 10)).toISOString() });
passou.etapas = ['Preparar matéria-prima'].concat(passou.etapas);
passou.progresso.enfestosCheck = { F1: true }; passou.progresso.enfestosSeq = { F1: em(16, 8) };
const ontem = Object.assign(os('0071', { 'Preparar matéria-prima': em(15, 7), 'Corte': em(15, 9) }));
ontem.etapas = ['Preparar matéria-prima'].concat(ontem.etapas);
ontem.progresso.enfestosCheck = { F1: true }; ontem.progresso.enfestosSeq = { F1: em(15, 8) };
const hp = rodar([passou, ontem], 'dia').h;
const linha70 = k => (hp[k].listaOS.find(r => r.os === '0070') || {});
const hora = t => t ? new Date(t).getHours() + ':' + String(new Date(t).getMinutes()).padStart(2, '0') : null;
[['materiaPrima', '7:00', '8:00'], ['enfestando', '8:00', '9:00'], ['cortando', '9:00', '10:00'],
 ['separando', '10:00', '10:30'], ['corte', '10:30', '11:00'], ['costurando', '11:00', null]].forEach(([k, ent, sai]) => {
  const r = linha70(k);
  ok(`0070 consta em ${k}, entrada ${ent}${sai ? ', saída ' + sai : ' e segue lá'}`,
     r.entrada === 200 && hora(r.entrouEm) === ent && (sai ? r.saida === 200 && hora(r.saiuEm) === sai : !r.saida), r);
});
ok('a 0071, que andou ontem, não entra na lista de nenhum quadro do Dia de hoje',
   Object.keys(hp).every(k => hp[k].listaOS.every(r => r.os !== '0071' || !(r.entrada || r.saida))),
   Object.keys(hp).map(k => [k, hp[k].listaOS.filter(r => r.os === '0071')]));

console.log('');
console.log('-- o quadro Estoque lista as OS finalizadas no dia (07/10/2026) --');
// A 0080 foi carimbada Estoque À MÃO hoje às 11:00, sem a caixa; a 0081 teve
// a caixa Estoque marcada hoje às 10:00; a 0082 foi para o estoque ontem.
const est80 = Object.assign(os('0080', { 'Corte': em(15, 9), 'Ensaque': em(15, 15) }), {
  statusOS: 'estoque', statusOSEm: new Date(em(16, 11)).toISOString(),
  statusHist: [{ k: 'estoque', em: em(16, 11), c: 'estoque' }] });
const est81 = os('0081', { 'Corte': em(14, 9), 'Ensaque': em(14, 15), 'Estoque': em(16, 10) });
const est82 = os('0082', { 'Corte': em(14, 9), 'Ensaque': em(14, 15), 'Estoque': em(15, 10) });
const he = rodar([est80, est81, est82], 'dia');
const le = o => he.h.estoque.listaOS.find(r => r.os === o) || {};
ok('a OS carimbada Estoque à mão conta no cartão Estoque', he.d.estoque.pecas === 600, he.d.estoque);
ok('e aparece no Dia em que foi finalizada, com a hora do carimbo',
   le('0080').entrada === 200 && new Date(le('0080').entrouEm).getHours() === 11, le('0080'));
ok('a da caixa Estoque marcada hoje também', le('0081').entrada === 200 && new Date(le('0081').entrouEm).getHours() === 10, le('0081'));
ok('a finalizada ontem não entra na lista de hoje', !le('0082').entrada, le('0082'));

console.log('');
console.log('-- carimbo de ontem perdido não faz a OS sair hoje de um quadro anterior (07/10/2026) --');
// A 0090 (a 0628 real): Preparo marcado ontem 14:00 e carimbada Separando à mão
// ontem 17:00. O diário nasceu depois e anotou o Separando SEM o `c`; hoje às
// 11:00 ela foi carimbada Ensacado. Antes ela "saía" da matéria-prima hoje.
const est90 = Object.assign(os('0090', { 'Preparar matéria-prima': em(15, 14) }), {
  statusOS: 'ensacado', statusOSEm: new Date(em(16, 11)).toISOString(),
  statusHist: [{ k: 'separando', em: em(15, 17) }, { k: 'ensacado', em: em(16, 11), c: 'ensacado' }] });
est90.etapas = ['Preparar matéria-prima'].concat(est90.etapas);
// A 0091 (a 0621 real): enfesto ontem 11:00, carimbada Ensacado ontem 15:00 (a
// segunda data) e esse carimbo se perdeu; hoje a caixa Corte foi marcada às 09:00.
const est91 = Object.assign(os('0091', { 'Corte': em(16, 9) }), {
  finalizadaEm: new Date(em(15, 15)).toISOString() });
est91.progresso.enfestosCheck = { F1: true }; est91.progresso.enfestosSeq = { F1: em(15, 11) };
const h9 = rodar([est90, est91], 'dia').h;
const mov9 = (k, n) => h9[k].listaOS.filter(r => r.os === n && (r.entrada || r.saida));
ok('0090 não sai hoje da matéria-prima (o Separando anotado de ontem vale como carimbo)',
   !mov9('materiaPrima', '0090').length, h9.materiaPrima.listaOS);
ok('0090 sai hoje do Separando, que é o que o carimbo de hoje fez', (mov9('separando', '0090')[0] || {}).saida === 200, h9.separando.listaOS);
ok('0091 não sai hoje do enfesto (a segunda data de ontem vale como carimbo)',
   !mov9('enfestando', '0091').length, h9.enfestando.listaOS);

const hl = rodar([limpa], 'dia').h;
ok('o "Não iniciado" à mão devolve a OS à folha (Cortando) na hora em que foi dado',
   ivDe(hl, 'separando').saida === 200 && ivDe(hl, 'cortando').entrada === 200, [hl.separando.listaOS, hl.cortando.listaOS]);

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
