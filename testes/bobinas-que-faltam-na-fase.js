/* Rode com:  node testes/bobinas-que-faltam-na-fase.js

   "10/2" NA COLUNA DA FASE (22/09/2026, Junior: "na coluna que mostra
   quantidade de bobinas necessárias para serem reservadas em cada fase... 
   mostrar no mesmo número a quantidade necessárias/quantidade que falta").

   A falta é POR PRATELEIRA (tecido + cor) e a coluna é POR FASE. Descer uma na
   outra é onde a conta pode mentir, e é o que este teste guarda:

     · a falta desce pela MESMA proporção de faltaParaCompletarOS — 40% do quilo
       faltando = 40% das bobinas de cada fase daquele pano;
     · duas fases do mesmo pano dividem a falta; a fase de outro pano não é
       tocada;
     · arredonda PARA CIMA: meia bobina que falta obriga a comprar uma inteira;
     · a falta da fase PARA no que a fase precisa. Quando o saldo já está
       negativo — pano baixado sem entrada correspondente —, a falta em quilos é
       maior do que a OS inteira prevê, e a coluna diria "precisa de 1, faltam
       21". Essa dívida é da prateleira e mora no kg, não na coluna da fase;
     · fase sem bobina prevista (viés, ribana sem cadastro) não inventa falta.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome + ' no app.js'); process.exit(1); }
  const fim = src.indexOf('\n}', ini);
  return src.slice(ini, fim + 2);
}
function pegaConst(nome) {
  const m = src.match(new RegExp('^const ' + nome + ' = .*$', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
}

/* materialPorFaseOS entra DUBLADA: ela puxa o consumo de enfesto inteiro, e o
   que se prova aqui é a conta de cima — quanto da bobina de cada fase falta. */
function monta(fases) {
  return new Function('FASES', `
    function materialPorFaseOS() { return FASES; }
    ${pegaConst('CEIL_BOBINA_EPS')}
    ${pegaFuncao('_normNome')}
    ${pegaFuncao('faltaParaCompletarOS')}
    ${pegaFuncao('fatiaQueFaltaPorTecidoCor')}
    ${pegaFuncao('bobinasQueFaltamNaFase')}
    return { fatiaQueFaltaPorTecidoCor, bobinasQueFaltamNaFase };
  `)(fases);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ----------------------
   Uma camiseta tricolor: dois corpos do mesmo preto, um corpo branco, o viés
   (que não gasta bobina) e a ribana preta sem previsão de bobina. */
const FASES = [
  { ordem: 1, nome: 'Corpo 1', tecido: 'Malha Algodão', cor: 'Preto Malha Algodão',  kg: 120, bobinas: 6 },
  { ordem: 2, nome: 'Corpo 2', tecido: 'Malha Algodão', cor: 'Preto Malha Algodão',  kg: 100, bobinas: 4 },
  { ordem: 3, nome: 'Corpo 3', tecido: 'Malha Algodão', cor: 'Branco Malha Algodão', kg: 80,  bobinas: 4 },
  { ordem: 4, nome: 'Viés',    tecido: 'Malha Algodão', cor: 'Preto Malha Algodão',  kg: 3,   bobinas: 0 },
  { ordem: 5, nome: 'Ribana',  tecido: 'Ribana Malha Algodão', cor: 'Preto Ribana Malha Algodão', kg: 2.83, bobinas: null }
];
const api = monta(FASES);
const OS = { id: 'os1' };
const falta = (tecido, cor, n) => ({ tecidoNome: tecido, corNome: cor, precisa: 0, disponivel: 0, falta: n });
const porFase = (faltando) => {
  const fatias = api.fatiaQueFaltaPorTecidoCor(OS, faltando);
  return FASES.map(f => api.bobinasQueFaltamNaFase(f, fatias));
};

/* ---------- 1. a falta desce para as fases daquele pano ---------- */
// Preto: 220 kg previstos. Faltam 88 = 40% -> Corpo 1: 6x0,4 = 2,4 -> 3
//                                            Corpo 2: 4x0,4 = 1,6 -> 2
let r = porFase([falta('Malha Algodão', 'Preto Malha Algodão', 88)]);
ok('1. 40% do quilo faltando = 40% das bobinas de cada fase (3 e 2)',
   r[0] === 3 && r[1] === 2, r);
ok('2. a fase do pano que NAO falta fica em zero', r[2] === 0, r);
ok('3. o vies nao gasta bobina, entao nao tem bobina faltando', r[3] === 0, r);
ok('4. fase sem bobina prevista nao inventa falta', r[4] === 0, r);

/* ---------- 2. meia bobina arredonda para cima ---------- */
r = porFase([falta('Malha Algodão', 'Preto Malha Algodão', 11)]);  // 5% de 220
ok('5. 5% de 6 bobinas = 0,3 -> 1 (meia bobina que falta obriga a comprar uma)',
   r[0] === 1 && r[1] === 1, r);

/* ---------- 3. o saldo negativo nao estoura a coluna ---------- */
/* O caso de verdade: OS 0561, em 21/09/2026. A prateleira da Ribana Preta
   estava em -56,454 kg (saidas sem entrada correspondente), a OS precisava de
   2,830 kg, e a lista dizia "faltam 59,284 kg (1 bob)". A coluna da fase nao
   pode dizer que faltam 21 bobinas de uma fase que gasta 1: a divida e da
   prateleira, e o lugar dela e o kg. */
const FASES_RIB = [
  { ordem: 1, nome: 'Ribana', tecido: 'Ribana Malha Algodão', cor: 'Preto Ribana Malha Algodão', kg: 2.83, bobinas: 1 }
];
const api2 = monta(FASES_RIB);
const fatias2 = api2.fatiaQueFaltaPorTecidoCor(OS,
  [falta('Ribana Malha Algodão', 'Preto Ribana Malha Algodão', 59.284)]);
ok('6. falta maior que o previsto para na bobina da fase (1, nunca 21)',
   api2.bobinasQueFaltamNaFase(FASES_RIB[0], fatias2) === 1,
   api2.bobinasQueFaltamNaFase(FASES_RIB[0], fatias2));

/* ---------- 4. sem falta nenhuma, a coluna nao ganha barra ---------- */
ok('7. OS com pano: nenhuma fase tem bobina faltando',
   porFase([]).every(n => n === 0), porFase([]));
ok('8. falta de poucos gramas ainda e uma bobina (nao arredonda para zero)',
   porFase([falta('Malha Algodão', 'Preto Malha Algodão', 0.27)])[0] === 1,
   porFase([falta('Malha Algodão', 'Preto Malha Algodão', 0.27)]));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
