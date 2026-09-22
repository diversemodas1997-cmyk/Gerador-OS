/* Rode com:  node testes/compra-falta-automatica.js

   O PANO QUE FALTA ENTRA SOZINHO NA LISTA DE COMPRA (22/09/2026, Junior:
   "insira automaticamente no quadro itens da compra todos os tecidos que sao
   apontados como faltantes no quadro reserva de material").

   O que este teste guarda:

     - uma linha por OS sem pano, com `origem:'falta'` e o `osId` — e NUNCA duas
       para a mesma OS: a lista e redesenhada a cada visita a tela, e um
       acrescimo sem chave viraria dezenas de linhas iguais no fim do dia;
     - SO AS FASES DO TECIDO QUE FALTA entram na conta. Uma tricolor em que so o
       preto falta nao pode mandar comprar o off-white que esta na prateleira;
     - a linha SE DESFAZ SOZINHA quando a falta acaba — e a mesma regra do
       vermelho no quadro do reservado: lancada a entrada, some;
     - a segunda passada nao mexe em nada (idempotencia), senao a tela gravaria
       o blob inteiro a cada desenho;
     - item somado a mao continua intocado.

   Recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
function pegaConst(nome) {
  const i = src.search(new RegExp('^const ' + nome + ' = ', 'm'));
  if (i < 0) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return src.slice(i, src.indexOf(';\n', i) + 1);
}

/* O consumo do enfesto e a falta de tecido entram DUBLADOS: os dois tem teste
   proprio (compra-necessidade-bruta.js e falta-de-tecido.js), e o que se prova
   aqui e o que a lista de compra faz com a resposta deles. */
function monta(mundo) {
  return new Function('M', `
    var STATE = M.STATE;
    var uid = () => 'novo' + (M.n++);
    function consumoEnfestoOS(o) { return M.consumo[o.id] || []; }
    function faltaDeTecidoParaOS(o) { return M.falta[o.id] || []; }
    function compraOsSimulada() { return null; }
    function calcularSaldosEstoque() { return { detalhe: M.saldos || [] }; }
    ${pegaFuncao('_normNome')}
    ${pegaConst('CEIL_BOBINA_EPS')}
    ${pegaFuncao('parseBobinas')}
    ${pegaFuncao('bobinaInteira')}
    ${pegaFuncao('ehFaseRibana')}
    ${pegaFuncao('bobinasEfetivasFase')}
    ${pegaFuncao('osComMaterialReservado')}
    ${pegaFuncao('compraConsumoItemDaFalta')}
    ${pegaFuncao('compraConsumoItem')}
    ${pegaFuncao('compraNecessidadeBruta')}
    ${pegaFuncao('compraFaltasAbertas')}
    ${pegaConst('_cpChaveFalta')}
    ${pegaFuncao('compraAplicarFaltas')}
    return { compraFaltasAbertas, compraAplicarFaltas, compraNecessidadeBruta,
             compraConsumoItem, plano: () => STATE.compraPlano };
  `)(mundo);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ----------------------
   A OS 0561 e uma tricolor: dois corpos de moletom (pano na prateleira) e uma
   ribana preta, que e a que falta. A grade preve bobina em todas as fases. */
const OS = { id: 'os561', os: '0561', gradeId: 'g1' };
const FASES_OS = [
  { ordem: 1, nomeEnf: 'Moletom',              tecidoReal: 'Moletom',              corReal: 'Preto Moletom',              kg: 40,   camadas: 36, camadasCheias: 36 },
  { ordem: 2, nomeEnf: 'Moletom',              tecidoReal: 'Moletom',              corReal: 'Off-White Moletom',          kg: 30,   camadas: 36, camadasCheias: 36 },
  { ordem: 3, nomeEnf: 'Ribana Malha Algodao', tecidoReal: 'Ribana Malha Algodao', corReal: 'Preto Ribana Malha Algodao', kg: 2.83, camadas: 18, camadasCheias: 18 }
];
const mundoBase = () => ({
  n: 1,
  STATE: {
    ordens: [OS],
    grades: [{ id: 'g1', nome: '2M-2G1-2G3 | CM.TRI | 116.5cm', fases: [
      { ordem: 1, bobinas: 3 }, { ordem: 2, bobinas: 2 }, { ordem: 3, bobinas: 1 }] }],
    desenhos: [],
    estoqueMov: [{ id: 'm1', tipo: 'saida', origem: 'os', osId: 'os561', osNumero: '0561',
                   kg: 72.83, status: 'reservado' }],
    compraPlano: []
  },
  consumo: { os561: FASES_OS },
  falta: { os561: [{ tecidoNome: 'Ribana Malha Algodao', corNome: 'Preto Ribana Malha Algodao',
                     precisa: 2.83, disponivel: -56.454, falta: 59.284 }] },
  saldos: [{ tecidoNome: 'Ribana Malha Algodao', corNome: 'Preto Ribana Malha Algodao', disponivel: -56.454 },
           { tecidoNome: 'Moletom', corNome: 'Preto Moletom', disponivel: 500 }]
});

/* ---------- 1. a OS sem pano entra como linha automatica ---------- */
let M = mundoBase();
let api = monta(M);
ok('1. a falta aberta e vista', api.compraFaltasAbertas().length === 1, api.compraFaltasAbertas());
ok('2. aplicar diz que mexeu', api.compraAplicarFaltas() === true);
let plano = api.plano();
ok('3. nasceu uma linha, marcada como automatica e presa ao osId',
   plano.length === 1 && plano[0].origem === 'falta' && plano[0].osId === 'os561'
   && plano[0].osNumero === '0561', plano);

/* ---------- 2. nao duplica, nao regrava a toa ---------- */
ok('4. a segunda passada nao mexe em nada (idempotencia)',
   api.compraAplicarFaltas() === false && api.plano().length === 1, api.plano());

/* ---------- 3. so as fases do tecido que falta entram na conta ---------- */
const bruta = api.compraNecessidadeBruta(api.plano());
ok('5. so o pano que falta entra na necessidade bruta',
   bruta.length === 1 && bruta[0].tecidoNome === 'Ribana Malha Algodao',
   bruta.map(b => b.tecidoNome + '/' + b.corNome));
ok('6. o quilo e o DA OS (2,830), nao o buraco da prateleira (59,284)',
   Math.abs(bruta[0].kg - 2.83) < 0.001, bruta[0].kg);
ok('7. a comprar nao desconta saldo negativo: compra-se o que a OS precisa',
   Math.abs(bruta[0].kgComprar - 2.83) < 0.001, bruta[0].kgComprar);
ok('8. a bobina vem do cadastro da grade daquela fase (1)',
   bruta[0].bobinas === 1, bruta[0].bobinas);

/* ---------- 4. lancada a entrada, a linha sai sozinha ---------- */
M.falta.os561 = [];
ok('9. sem falta, aplicar tira a linha',
   api.compraAplicarFaltas() === true && api.plano().length === 0, api.plano());

/* ---------- 5. o item somado a mao nao e tocado ---------- */
M = mundoBase();
api = monta(M);
M.STATE.compraPlano.push({ id: 'meu', gradeId: 'g1', camadas: 30, repeticoes: 1, criadoPor: 'eu' });
api.compraAplicarFaltas();
ok('10. o item manual continua la, ao lado do automatico',
   api.plano().length === 2 && api.plano().some(i => i.id === 'meu'), api.plano());
M.falta.os561 = [];
api.compraAplicarFaltas();
ok('11. e sobrevive a saida do automatico',
   api.plano().length === 1 && api.plano()[0].id === 'meu', api.plano());

/* ---------- 6. OS ja baixada nao entra ---------- */
M = mundoBase();
M.STATE.estoqueMov[0].status = 'consumido';
api = monta(M);
ok('12. OS que ja desceu da prateleira nao gera linha de compra',
   api.compraFaltasAbertas().length === 0 && api.compraAplicarFaltas() === false, api.plano());

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
