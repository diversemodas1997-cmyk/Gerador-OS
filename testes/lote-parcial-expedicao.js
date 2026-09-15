/* Rode com:  node testes/lote-parcial-expedicao.js

   LOTE PARCIAL: para onde vão as peças quando a OS é alocada no planejamento
   de expedição.

   A regra mudou DUAS vezes. Em 14/08/2026 alocar uma carga de ida passou a
   mandar os pacotes do Estoque de corte direto para EXPEDIÇÃO — alocar no plano
   já é dizer que aquele pacote vai embarcar.

   Em 14/09/2026, com as DUAS UNIDADES, mudou de novo: a OE é viagem INTERNA
   entre Descalvado e São Carlos, e Expedição é o fim do fluxo, não o meio. A
   fração alocada passou a ir para EM TRÂNSITO da perna (ida ou volta), saindo do
   campo em que a OS estiver — corte OU costurando da unidade de origem —, e o
   remanescente fica onde estava. Quem tira a OS do trânsito é a caixa de chegada
   do checklist: "Recebido em São Carlos" na ida, "Recebido em Descalvado" na
   volta. Expedição voltou a entrar só pela etapa marcada.

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
  // As constantes das duas unidades moram logo acima do array e ele as usa
  // (o `cond` das duas fases de costura sai daqui).
  recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades'),
  cortaArr('const FASES_ESTOQUE'),
  // O campo "Estoque de corte" so conta OS com o status Ensacado (a `cond` da
  // fase), entao o motor precisa saber ler o status.
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
      transitoIda: soma('transitoIda'), corteSC: soma('corteSC'),
      costurandoSC: soma('costurandoSC'), transitoVolta: soma('transitoVolta'),
      fios: soma('fios'), expedicao: soma('expedicao'),
      osTransitoIda: listaOS('transitoIda')
    };
  `);
  return fn(estado);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const CAMPOS = ['corte', 'costurando', 'transitoIda', 'corteSC', 'costurandoSC',
  'transitoVolta', 'fios', 'expedicao'];
// Confere os OITO campos, e não só os citados: campo esquecido de fora do
// esperado tem que dar falha, senão peça que vazou para o campo errado passa.
const confere = (nome, got, esperado) => {
  const bate = CAMPOS.every(k => (esperado[k] || 0) === got[k]);
  ok(nome, bate, Object.fromEntries(CAMPOS.filter(k => got[k]).map(k => [k, got[k]])));
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

// NO ESTOQUE DE CORTE = ENSACADA (15/09/2026). O campo deixou de entrar pela
// etapa Corte: peça que ainda está na mesa não é pano guardado. Entra quem tem
// o status ENSACADO, e por isso a OS deste teste marca o Ensaque depois do
// Corte — é o que a fábrica faz quando fecha o saco.
const noCorte = () => osBase({ 'Corte': true, 'Ensaque': true }, { 'Corte': 1, 'Ensaque': 2 });
// A DATA DA CARGA DECIDE O DESTINO (15/09/2026): ainda por sair -> Em trânsito;
// já saiu -> migra para o Estoque de corte da outra unidade. Datas relativas a
// HOJE, e não cravadas: cravadas, o teste mudaria de significado sozinho quando
// a data passasse.
const _dia = (n) => {
  const d = new Date(); d.setDate(d.getDate() + n);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')
    + '-' + String(d.getDate()).padStart(2, '0');
};
const FUTURO = _dia(7);    // a carga ainda vai sair
const PASSADO = _dia(-7);  // a carga já saiu
const cargaIda = (extra) => Object.assign({
  id: 'c1', osId: 'os_1', janelaId: 'j1', data: FUTURO, perna: 'ida',
  pacotes: [{ tam: 'P', tom: null }, { tam: 'M', tom: null }], volumes: 5
}, extra || {});

/* ---------- 1. o caminho normal: a ida sai de Descalvado ---------- */

confere('OS no corte, sem nada alocado: as 200 pç ficam no corte',
  saldos(estado(noCorte(), [])),
  { corte: 200 });

confere('alocada METADE numa carga de ida: 100 pç vão para Em trânsito · IDA',
  saldos(estado(noCorte(), [cargaIda()])),
  { corte: 100, transitoIda: 100 });

confere('alocada por INTEIRO: as 200 pç vão para Em trânsito · IDA',
  saldos(estado(noCorte(), [cargaIda({ pacotes: [
    { tam: 'P', tom: null }, { tam: 'M', tom: null },
    { tam: 'G', tom: null }, { tam: 'GG', tom: null }] })])),
  { corte: 0, transitoIda: 200 });

const so = saldos(estado(noCorte(), [cargaIda()]));
ok('a OS aparece na coluna OS do trânsito mesmo sem etapa nenhuma marcada lá',
  so.osTransitoIda.includes('0501'), so.osTransitoIda);

// A DECISÃO DE 14/09 (Junior): a fração viaja e o remanescente FICA ONDE ESTÁ.
// Antes a etapa Costura desmanchava a alocação inteira — a peça que já estava no
// caminhão voltava para o campo da costura. Agora a costura só manda no que
// ficou: quem está na estrada continua na estrada.
confere('marcada a etapa Costura, a ida continua valendo e sai de Costurando',
  saldos(estado(osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }),
    [cargaIda()])),
  { costurando: 100, transitoIda: 100 });

/* ---------- 2. quem tira a OS do trânsito é a caixa de chegada ---------- */

// O checklist da fábrica com as duas caixas das unidades (etapa 1, 14/09/2026).
// A UNIDADE DA COSTURA SAI DO NOME DA ETAPA (15/09/2026). Os 34 desenhos listam
// AS DUAS — "Costura CM.LISA | Descalvado" e "| São Carlos" —, e quem está no
// chão marca a que fez. A "Costura" pura é a etapa das OS antigas, de antes de
// a segunda unidade existir, e conta como Descalvado.
const TMPL_U = ['Corte', 'Recebido em São Carlos', 'Costura',
  'Costura CM.LISA | São Carlos',
  'Recebido em Descalvado', 'Retirada de fios', 'Ensaque', 'Expedição', 'Estoque'];
const osU = (check, seq) => Object.assign(osBase(check, seq), { etapas: TMPL_U.slice() });

// O Estoque de corte de São Carlos é o ENSACADO que chegou lá (15/09/2026):
// sem o ensaque a OS está sendo cortada, e cortar não é ter no estoque.
confere('a ida acaba quando "Recebido em São Carlos" é marcada: 200 pç no corte de SC',
  saldos(estado(osU({ 'Corte': true, 'Ensaque': true, 'Recebido em São Carlos': true },
    { 'Corte': 1, 'Ensaque': 2, 'Recebido em São Carlos': 3 }), [cargaIda()])),
  { corteSC: 200 });

confere('recebida em SC e costurando lá: o lote inteiro em Costurando · São Carlos',
  saldos(estado(osU({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
    { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3 }), [cargaIda()])),
  { costurandoSC: 200 });

/* ---------- 3. a volta é a mesma regra, do outro lado ---------- */

const cargaVolta = (extra) => cargaIda(Object.assign({ id: 'v1', perna: 'volta' }, extra || {}));
const emSC = () => osU({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
  { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3 });

confere('alocada METADE numa carga de volta: 100 pç saem de Costurando · SC',
  saldos(estado(emSC(), [cargaVolta()])),
  { costurandoSC: 100, transitoVolta: 100 });

confere('a volta não mexe na OS que ainda está em Descalvado',
  saldos(estado(noCorte(), [cargaVolta()])),
  { corte: 200 });

confere('a ida não mexe na OS que já está em São Carlos',
  saldos(estado(emSC(), [cargaIda()])),
  { costurandoSC: 200 });

confere('a volta acaba em "Recebido em Descalvado": 200 pç na Retirada de fios',
  saldos(estado(osU({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true,
    'Recebido em Descalvado': true },
    { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3, 'Recebido em Descalvado': 4 }),
    [cargaVolta()])),
  { fios: 200 });

/* ---------- 4. os campos que NÃO despacham ---------- */

confere('etapa Retirada de fios: o lote inteiro está lá, a alocação não conta',
  saldos(estado(osBase({ 'Corte': true, 'Costura': true, 'Retirada de fios': true },
    { 'Corte': 1, 'Costura': 2, 'Retirada de fios': 3 }), [cargaIda()])),
  { fios: 200 });

confere('etapa Expedição marcada: 200 pç lá, sem somar a fração alocada por cima',
  saldos(estado(osBase({ 'Corte': true, 'Expedição': true }, { 'Corte': 1, 'Expedição': 4 }),
    [cargaIda()])),
  { expedicao: 200 });

/* O caso torto: a caixa Expedição está marcada, mas o Corte foi marcado DEPOIS
   (a OS voltou para o corte). O que este caso protege é a SOMA: venha a peça de
   onde vier, as 200 têm que continuar existindo em algum campo. Sem a guarda do
   campo de origem, sumiam 100 no meio do caminho.

   O DESTINO mudou em 15/09/2026, e a soma não. Antes a OS voltava mesmo para o
   Estoque de corte e a ida seguia valendo, repartindo 100/100. Agora o Estoque
   de corte só aceita quem está ENSACADO, e a última etapa marcada aqui é o Corte
   — o status é "Cortando", peça na mesa. Sem campo de origem, não há fração que
   viaje: o lote inteiro conta na Expedição, que é a etapa que ela realmente tem
   marcada. As 200 continuam inteiras, que é o que este caso existe para provar. */
const torto = saldos(estado(osBase({ 'Corte': true, 'Expedição': true },
  { 'Corte': 5, 'Expedição': 4 }), [cargaIda()]));
confere('Expedição marcada e Corte remarcado por cima: o lote inteiro na Expedição',
  torto, { expedicao: 200 });
ok('  ... e as 200 pç continuam inteiras somando os campos',
  CAMPOS.reduce((t, k) => t + torto[k], 0) === 200,
  CAMPOS.reduce((t, k) => t + torto[k], 0));

confere('etapa Estoque (terminal): a OS sai de todos os campos',
  saldos(estado(osBase({ 'Corte': true, 'Estoque': true }, { 'Corte': 1, 'Estoque': 9 }),
    [cargaIda()])),
  {});

/* ---------- 5. quais cargas movem peça ---------- */

confere('carga em data CANCELADA não move peça',
  saldos(estado(noCorte(), [cargaIda()],
    [{ janelaId: 'j1', data: FUTURO, tipo: 'cancelada' }])),
  { corte: 200 });

confere('carga remarcada (não cancelada) move normalmente',
  saldos(estado(noCorte(), [cargaIda()],
    [{ janelaId: 'j1', data: FUTURO, tipo: 'remarcada', novaData: FUTURO }])),
  { corte: 100, transitoIda: 100 });

confere('carga ANTIGA (só volumes, sem pacotes) leva o lote inteiro',
  saldos(estado(noCorte(), [{ id: 'c9', osId: 'os_1', janelaId: 'j1', data: FUTURO,
    perna: 'ida', volumes: 9 }])),
  { transitoIda: 200 });

confere('duas cargas de ida, as duas por sair: somam no trânsito',
  saldos(estado(noCorte(), [
    cargaIda(),
    cargaIda({ id: 'c2', data: FUTURO, pacotes: [{ tam: 'G', tom: null }] })
  ])),
  { corte: 50, transitoIda: 150 });

/* ---------- 5b. a carga QUE JÁ SAIU migra para São Carlos ----------

   Junior, 15/09/2026: "migre de Estoque de corte | Descalvado para Estoque
   corte | São Carlos sempre que a OS for alocada no plano de expedição Desc x
   São Carlos". Quem dispara é a DATA: alocar é planejar, e uma carga marcada
   para a semana que vem não tirou pano nenhum da prateleira daqui. Chegado o
   dia, o caminhão saiu — e o estoque acompanha sem ninguém marcar nada. */

confere('carga de ida JÁ SAIU: a metade alocada migra para o corte de São Carlos',
  saldos(estado(noCorte(), [cargaIda({ data: PASSADO })])),
  { corte: 100, corteSC: 100 });

confere('carga inteira já saída: o lote todo migra, e Descalvado zera',
  saldos(estado(noCorte(), [cargaIda({ data: PASSADO, pacotes: [
    { tam: 'P', tom: null }, { tam: 'M', tom: null },
    { tam: 'G', tom: null }, { tam: 'GG', tom: null }] })])),
  { corteSC: 200 });

// As duas coisas convivem: parte foi na carga de ontem, parte vai na de amanhã.
// É por isso que a conta devolve uma LISTA de frações, e não uma só.
confere('uma carga já saída e outra por sair: os TRÊS campos ao mesmo tempo',
  saldos(estado(noCorte(), [
    cargaIda({ data: PASSADO, pacotes: [{ tam: 'P', tom: null }] }),
    cargaIda({ id: 'c2', data: FUTURO, pacotes: [{ tam: 'M', tom: null }, { tam: 'G', tom: null }] })
  ])),
  { corte: 50, corteSC: 50, transitoIda: 100 });

// A chegada encerra tudo: marcada a caixa, o lote inteiro é de São Carlos e não
// há mais fração nenhuma a mover.
confere('marcada a chegada, a migração para de valer: lote inteiro no corte de SC',
  saldos(estado(osU({ 'Corte': true, 'Ensaque': true, 'Recebido em São Carlos': true },
    { 'Corte': 1, 'Ensaque': 2, 'Recebido em São Carlos': 3 }), [cargaIda({ data: PASSADO })])),
  { corteSC: 200 });

/* A COSTURA TAMBÉM MIGRA (15/09/2026, Junior: "a migração sempre é Costurando
   Descalvado para Costurando São Carlos, mesmo que as etapas sejam
   fracionadas"). Este caso já esteve escrito ao contrário, com a premissa de
   que costurar de novo o que já foi costurado não acontece. Acontece: é
   justamente por isso que os 34 desenhos listam as DUAS costuras. A peça sai
   daqui meio-costurada e vai terminar lá. */
confere('da costura, a carga já saída migra para Costurando · São Carlos',
  saldos(estado(osU({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }),
    [cargaIda({ data: PASSADO })])),
  { costurando: 100, costurandoSC: 100 });

// E a carga que ainda vai sair continua indo para o trânsito, como sempre.
confere('da costura, a carga por sair continua indo para o trânsito',
  saldos(estado(osU({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }),
    [cargaIda({ data: FUTURO })])),
  { costurando: 100, transitoIda: 100 });

confere('carga de outra OS não mexe nesta',
  saldos(estado(noCorte(), [cargaIda({ id: 'c3', osId: 'os_outra' })])),
  { corte: 200 });

/* ---------- 6. contagem manual continua ajustando ----------
   O trânsito NÃO tem contagem manual (não tem movKey): ninguém conta prateleira
   de caminhão andando. O ajuste vive nos campos das unidades. */

const comAjuste = estado(noCorte(), [cargaIda()]);
comAjuste.corteMov = [{ id: 'm1', tipo: 'saida', qtd: 30,
  tecidoNome: 'Malha Algodão', corNome: 'Preto Malha Algodão' }];
confere('lançamento manual do corte desconta do que ficou, não do que viajou',
  saldos(comAjuste),
  { corte: 70, transitoIda: 100 });

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
