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
  corta('function _produtosDosComponentes'),
  corta('function produtosOS'),
  corta('function produtosPorTecidoCorOS'),
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
  cortaLinha('const TERMINAL_ETAPA_RE'),
  corta('function faseAtualOS'),
  corta('function _expCancelSet'),
  corta('function _expEmbarcadoOS'),
  corta('function _dashTurnoDaCarga'),
  corta('function _dashPesoTurnos'),
  corta('function _dashCartoesDaOS'),
  corta('function _dashChavePorIdx'),
  corta('function _dashFluxoDados')
].join('\n');

// O resumo do painel com o STATE dado.
function dash(estado) {
  const fn = new Function('STATE', `
    const corCanonicaPorTecido = (cor) => cor || '';
    // Sem grade no fixture a folha nao tem Total geral, e os produtos saem dos
    // componentes (qtdTotal / qtdPorPeca) — ver produtosOS.
    const totaisPorTamanhoTomOS = () => ({ totalGeral: 0 });
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

const CAMPOS = ['cortando', 'corte', 'corteSC', 'costurando', 'costurandoSC',
  'idaManha', 'idaTarde', 'voltaManha', 'voltaTarde',
  'recDesc', 'recSC', 'fios', 'estoque'];

// Confere os DOZE cartões, e não só os citados: cartão esquecido de fora do
// esperado tem que dar falha, senão peça que vazou para o campo errado passa.
const confere = (nome, got, esperado) => {
  const bate = CAMPOS.every(k => (esperado[k] || 0) === got[k].pecas);
  ok(nome, bate, Object.fromEntries(CAMPOS.filter(k => got[k].pecas).map(k => [k, got[k].pecas])));
};

/* ---------------------- o mundo do teste ---------------------- */

// Uma OS de 200 PRODUTOS, todos do mesmo tecido+cor, com o checklist da fábrica.
// São 600 peças cortadas (a frente e as duas mangas), e o que os campos contam
// desde 16/09/2026 é o produto: 200.
// etapasSeq é o carimbo de QUANDO cada etapa foi marcada: é ele que decide a
// fase atual no modelo sobreposto.
const osBase = (check, seq) => ({
  id: 'os_1', os: '0501', modeloNome: 'Camiseta', data: '2026-09-14',
  gradeId: 'g1',
  etapas: ['Corte', 'Recebido em São Carlos', 'Costura',
           'Costura CM.LISA | São Carlos',
           'Recebido em Descalvado', 'Retirada de fios', 'Ensaque', 'Expedição', 'Estoque'],
  progresso: { etapasCheck: check, etapasSeq: seq },
  componentes: [
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdPorPeca: 1, qtdTotal: 200 },
    { materialNome: 'Malha Algodão', corNome: 'Preto Malha Algodão', qtdPorPeca: 2, qtdTotal: 400 }
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

// NO ESTOQUE DE CORTE = ENSACADA (15/09/2026). O campo deixou de entrar pela
// etapa Corte: peça que ainda está na mesa não é pano guardado. Entra quem tem
// o status ENSACADO, e por isso a OS deste teste marca o Ensaque depois do
// Corte — é o que a fábrica faz quando fecha o saco.
const noCorte = () => osBase({ 'Corte': true, 'Ensaque': true }, { 'Corte': 1, 'Ensaque': 2 });
const carga = (extra) => Object.assign({
  id: 'c1', osId: 'os_1', janelaId: 'j1', data: '2026-09-22', perna: 'ida',
  pacotes: [{ tam: 'P', tom: null }, { tam: 'M', tom: null }], volumes: 5
}, extra || {});

/* ---------- 1. o caminho, campo a campo ---------- */

confere('OS no corte: as 200 pç no Estoque de corte · Descalvado',
  dash(estado(noCorte(), [])),
  { corte: 200 });

// A UNIDADE DA COSTURA SAI DO NOME DA ETAPA (15/09/2026). Os 34 desenhos listam
// AS DUAS — "Costura CM.LISA | Descalvado" e "| São Carlos" —, e quem está no
// chão marca a que fez. A "Costura" pura é a etapa das OS antigas, de antes de
// a segunda unidade existir, e conta como Descalvado.
confere('Costura marcada, sem passar por São Carlos: Costurando · Descalvado',
  dash(estado(osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }), [])),
  { costurando: 200 });

// RECEBIDO É CARIMBO DE PASSAGEM (15/09/2026): a caixa marcada conta no cartão
// de Recebido mesmo depois de a OS ter andado, porque ela FOI recebida — e a
// mesma OS conta também no campo em que está agora. Os dois somam; não dividem.
// E o Estoque de corte de São Carlos é o ENSACADO que já chegou lá: sem o
// ensaque a OS está sendo cortada, e cortar não é ter no estoque.
confere('ensacada e recebida em São Carlos: corte de lá E cartão de recebido',
  dash(estado(osBase({ 'Corte': true, 'Ensaque': true, 'Recebido em São Carlos': true },
                     { 'Corte': 1, 'Ensaque': 2, 'Recebido em São Carlos': 3 }), [])),
  { corteSC: 200, recSC: 200 });

/* RECEBIDA EM SÃO CARLOS NÃO ESTÁ MAIS NA MESA DAQUI. A caixa de recebimento
   acende o status Ensacado (15/09/2026): ela é o que diz que o pano está na
   prateleira DE LÁ, esperando a máquina — e isso é o estado Ensacado, venha ele
   da caixa de Ensaque ou da de chegada. */
confere('cortada e recebida em São Carlos: vai para o corte de lá, não fica na mesa',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2 }), [])),
  { corteSC: 200, recSC: 200 });

/* O CASO QUE A FÁBRICA ACHOU (OS 0519 e 0539, 15/09/2026): ensacada, costurada
   AQUI, expedida e recebida lá. Nem a caixa de expedição nem a de chegada
   acendiam status, então a última que acendia continuava sendo a costura DAQUI:
   a peça viajava, chegava, e a lista seguia dizendo "Costurando | Descalvado".
   Marcar a chegada tem de tirá-la de lá. */
confere('costurada aqui, expedida e recebida lá: vai para o corte de São Carlos',
  dash(estado(osBase({ 'Corte': true, 'Ensaque': true, 'Costura': true,
                       'Recebido em São Carlos': true },
                     { 'Corte': 1, 'Ensaque': 2, 'Costura': 3,
                       'Recebido em São Carlos': 4 }), [])),
  { corteSC: 200, recSC: 200 });

// E dali a costura de lá a leva para Costurando · São Carlos, que é o caminho
// que a fábrica descreve: corte SC -> costurando SC -> retirada de fios.
confere('e a costura de lá, marcada depois, leva para Costurando · São Carlos',
  dash(estado(osBase({ 'Corte': true, 'Ensaque': true, 'Costura': true,
                       'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
                     { 'Corte': 1, 'Ensaque': 2, 'Costura': 3,
                       'Recebido em São Carlos': 4, 'Costura CM.LISA | São Carlos': 5 }), [])),
  { costurandoSC: 200, recSC: 200 });

confere('Costura depois de São Carlos: Costurando · São Carlos',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3 }), [])),
  { costurandoSC: 200, recSC: 200 });

// A volta cai na Retirada de fios, e o cartão de Recebido em Descalvado é a
// fatia dela que chegou pela caixa de recebimento.
confere('Recebido em Descalvado: Retirada de fios E o cartão de recebido',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true, 'Recebido em Descalvado': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3, 'Recebido em Descalvado': 4 }), [])),
  { fios: 200, recDesc: 200, recSC: 200 });

// Os dois recebimentos continuam carimbados: a OS passou pelos dois.
confere('Retirada de fios marcada por último: fios, e os dois recebidos carimbados',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true, 'Recebido em Descalvado': true, 'Retirada de fios': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3, 'Recebido em Descalvado': 4, 'Retirada de fios': 5 }), [])),
  { fios: 200, recDesc: 200, recSC: 200 });

confere('Estoque marcado: sai do fluxo em processo e vai para o cartão final',
  dash(estado(osBase({ 'Corte': true, 'Retirada de fios': true, 'Estoque': true },
                     { 'Corte': 1, 'Retirada de fios': 2, 'Estoque': 3 }), [])),
  { estoque: 200 });

/* Expedição não tem cartão neste painel (não está na lista pedida) e também não
   acende status — a OS lá deriva "Cortando" do corte que ficou para trás. É o
   caso que a guarda `atual < 0` do cartão Cortando protege: ela está NUM campo,
   então não está na mesa, e não pode vazar para cartão nenhum. */
confere('OS em Expedição: não aparece em cartão nenhum, nem na mesa de corte',
  dash(estado(osBase({ 'Corte': true, 'Expedição': true }, { 'Corte': 1, 'Expedição': 2 }), [])),
  {});

// O cartão Estoque conta a CAIXA, não a última etapa: marcada uma vez, a OS
// conta ali mesmo que o fluxo tenha continuado depois. É a exceção pedida em
// 15/09/2026, e por isso a peça aparece nos dois cartões.
confere('Estoque marcado e outra etapa marcada DEPOIS: conta nos dois cartões',
  dash(estado(osBase({ 'Estoque': true, 'Corte': true, 'Ensaque': true },
                     { 'Estoque': 1, 'Corte': 2, 'Ensaque': 3 }), [])),
  { corte: 200, estoque: 200 });

confere('OS sem etapa nenhuma marcada: não conta em cartão nenhum',
  dash(estado(osBase({}, {}), [])),
  {});

/* A COSTURA FRACIONADA (15/09/2026, Junior: "alguns produtos são costurados em
   etapas fracionadas em diferentes unidades" e "quando os dois costurando estão
   preenchidos o volume migra para Costurando São Carlos").

   Com as duas caixas marcadas vence SÃO CARLOS, e não a última marcada: a
   costura fracionada começa aqui e termina lá, e a metade final é a de lá. Por
   isso o primeiro caso marca a de Descalvado DEPOIS — se a ordem mandasse, ele
   daria Descalvado. */
confere('as duas costuras: vence São Carlos, mesmo marcada por último a daqui',
  dash(estado(osBase({ 'Corte': true, 'Recebido em São Carlos': true,
                       'Costura CM.LISA | São Carlos': true, 'Costura': true },
                     { 'Corte': 1, 'Recebido em São Carlos': 2,
                       'Costura CM.LISA | São Carlos': 3, 'Costura': 4 }), [])),
  { costurandoSC: 200, recSC: 200 });

// E na ordem natural (daqui primeiro, lá depois) o resultado é o mesmo.
confere('as duas costuras, na ordem natural: também São Carlos',
  dash(estado(osBase({ 'Corte': true, 'Costura': true,
                       'Costura CM.LISA | São Carlos': true },
                     { 'Corte': 1, 'Costura': 2,
                       'Costura CM.LISA | São Carlos': 3 }), [])),
  { costurandoSC: 200 });

/* ---------- 1b. o Estoque de corte é só o que está ENSACADO ---------- */
/* Medido em 15/09/2026, o campo tinha 11 OS e só UMA estava ensacada: 3 ainda
   sendo cortadas, 4 já carimbadas como Estoque, 1 em Preparando matéria-prima e
   1 em Costurando. Ele entrava pela etapa CORTE, e corte é trabalho na mesa —
   não é pano guardado esperando a costura. */

confere('só o Corte marcado (status Cortando): NÃO é estoque de corte, é mesa',
  dash(estado(osBase({ 'Corte': true }, { 'Corte': 1 }), [])),
  { cortando: 200 });

confere('Ensaque marcado depois do Corte: aí sim entra no Estoque de corte',
  dash(estado(osBase({ 'Corte': true, 'Ensaque': true }, { 'Corte': 1, 'Ensaque': 2 }), [])),
  { corte: 200 });

// O carimbo à mão vale: na fábrica ninguém marca a caixa Ensaque, o ensaque é
// apontado carimbando o status na lista de OS. Exigir a caixa esvaziaria o campo.
const carimbadaEnsacada = () => {
  const o = osBase({ 'Corte': true }, { 'Corte': 1000 });
  o.statusOS = 'ensacado';
  o.statusOSPor = 'enfesto.corte@diverse.local';
  o.statusOSEm = new Date(5000).toISOString();
  return o;
};
confere('carimbada Ensacada à mão, sem a caixa: entra igual',
  dash(estado(carimbadaEnsacada(), [])),
  { corte: 200 });

// E o carimbo VELHO não segura a OS ali: marcada uma etapa depois dele, o
// checklist retoma e o status deixa de ser Ensacado.
const carimboVencido = () => {
  const o = osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1000, 'Costura': 9000 });
  o.statusOS = 'ensacado';
  o.statusOSPor = 'enfesto.corte@diverse.local';
  o.statusOSEm = new Date(5000).toISOString();
  return o;
};
confere('carimbo de Ensacado vencido por etapa nova: sai do corte, vai para a costura',
  dash(estado(carimboVencido(), [])),
  { costurando: 200 });

/* ---------- 1c. o cartão CORTANDO ---------- */
/* Cortar é trabalho em curso na mesa, e por isso nunca houve campo para ele.
   Quando o Estoque de corte passou a exigir o status Ensacado, a OS entre o
   corte e o ensaque deixou de aparecer em qualquer lugar do dashboard — e
   sumir da tela é pior do que aparecer no campo errado. */

confere('só o Corte marcado: aparece em Cortando, e em mais nenhum cartão',
  dash(estado(osBase({ 'Corte': true }, { 'Corte': 1 }), [])),
  { cortando: 200 });

confere('ensacada depois: sai de Cortando e entra no Estoque de corte',
  dash(estado(osBase({ 'Corte': true, 'Ensaque': true }, { 'Corte': 1, 'Ensaque': 2 }), [])),
  { corte: 200 });

// O cartão não repete ninguém: quem está cortando não está em campo nenhum.
confere('costurando: nem Cortando nem Estoque de corte',
  dash(estado(osBase({ 'Corte': true, 'Costura': true }, { 'Corte': 1, 'Costura': 2 }), [])),
  { costurando: 200 });

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
    osBase({ 'Corte': true, 'Recebido em São Carlos': true, 'Costura CM.LISA | São Carlos': true },
           { 'Corte': 1, 'Recebido em São Carlos': 2, 'Costura CM.LISA | São Carlos': 3 }),
    [carga({ perna: 'volta' })])),
  { costurandoSC: 100, voltaTarde: 100, recSC: 200 });

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
