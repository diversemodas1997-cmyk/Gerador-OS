/* Rode com:  node testes/falta-para-completar-os.js

   QUANTO FALTA PARA COMPLETAR A OS, no aviso de material reservado.

   O aviso dizia quanto falta de cada tecido e parava aí. Quem vai comprar
   precisa do TAMANHO do buraco: falta pouco para fechar a OS, ou falta quase
   tudo? `faltaParaCompletarOS` fecha a conta em bobinas e em quilos.

   O que este teste guarda, e é onde a conta pode mentir:

     · a BOBINA que falta sai da bobina PREVISTA (cadastro da grade), pela
       proporção do quilo — e não de kg ÷ peso médio da bobina. Duas fontes para
       o mesmo número dariam a coluna da tabela dizendo 10 e o aviso dizendo 11;
     · arredonda PARA CIMA: meia bobina que falta obriga a comprar uma inteira;
     · tecido sem previsão de bobina entra só com o quilo — inventar bobina é
       pior do que não dizer;
     · o TOTAL é o previsto inteiro da OS, todas as fases, e não só as que
       faltam: é contra ele que "faltam 4" quer dizer alguma coisa. */
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

const motor = [
  corta('function _normNome'),
  corta('function faltaParaCompletarOS')
].join('\n');

/* materialPorFaseOS entra DUBLADA: ela puxa o consumo de enfesto inteiro
   (comprimento, largura, camadas, gramatura, bobinas da grade) e aqui o que
   importa é a conta de cima — o que a OS prevê por tecido+cor, e quanto disso
   falta. As fases vêm prontas pelo teste, no mesmo formato que ela devolve. */
function rodar(fases, faltando) {
  const fn = new Function('FASES', 'FALTANDO', `
    function materialPorFaseOS() { return FASES; }
    ${motor}
    return faltaParaCompletarOS({ id: 'os1' }, FALTANDO);
  `);
  return fn(fases, faltando);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ---------------------- */

// Uma OS de três fases: duas de malha preta (corpo) e uma de ribana preta.
const FASES = [
  { ordem: 1, nome: 'Corpo 1', tecido: 'Malha Algodão', cor: 'Preto Malha Algodão', kg: 120, bobinas: 6 },
  { ordem: 2, nome: 'Corpo 2', tecido: 'Malha Algodão', cor: 'Preto Malha Algodão', kg: 100, bobinas: 4 },
  { ordem: 3, nome: 'Ribana',  tecido: 'Ribana Bulk',   cor: 'Preto Ribana Bulk',   kg: 85,  bobinas: 4 }
];
const falta = (tecido, cor, n) => ({ tecidoNome: tecido, corNome: cor, precisa: 0, disponivel: 0, falta: n });

/* ---------- 1. a bobina que falta sai da bobina prevista ---------- */

// Malha preta: 220 kg previstos em 10 bobinas. Faltam 88 kg = 40% -> 4 bobinas.
let r = rodar(FASES, [falta('Malha Algodão', 'Preto Malha Algodão', 88)]);
ok('40% do quilo faltando = 40% das bobinas',
   r.itens[0].faltaBob === 4 && r.itens[0].previstoBob === 10,
   { faltaBob: r.itens[0].faltaBob, previstoBob: r.itens[0].previstoBob });

// Meia bobina que falta obriga a comprar uma inteira.
r = rodar(FASES, [falta('Malha Algodão', 'Preto Malha Algodão', 11)]);
ok('meia bobina arredonda PARA CIMA (5% de 10 = 0,5 -> 1)',
   r.itens[0].faltaBob === 1, r.itens[0].faltaBob);

// Falta tudo: nao pode passar do previsto.
r = rodar(FASES, [falta('Malha Algodão', 'Preto Malha Algodão', 500)]);
ok('faltando mais do que o previsto, a bobina para no previsto',
   r.itens[0].faltaBob === 10, r.itens[0].faltaBob);

/* ---------- 2. o total é o previsto INTEIRO da OS ---------- */

r = rodar(FASES, [falta('Malha Algodão', 'Preto Malha Algodão', 88)]);
ok('o previsto soma TODAS as fases, nao so as que faltam',
   r.previstoBob === 14 && r.previstoKg === 305, { bob: r.previstoBob, kg: r.previstoKg });
ok('e a falta soma so o que falta',
   r.faltaBob === 4 && r.faltaKg === 88, { bob: r.faltaBob, kg: r.faltaKg });

// Duas prateleiras faltando: as bobinas de cada tecido somam (nao da para
// somar quilo de tecidos diferentes e dividir por um peso so).
r = rodar(FASES, [falta('Malha Algodão', 'Preto Malha Algodão', 110),
                  falta('Ribana Bulk', 'Preto Ribana Bulk', 85)]);
ok('duas prateleiras: 5 bobinas de malha + 4 de ribana = 9',
   r.faltaBob === 9 && r.faltaKg === 195, { bob: r.faltaBob, kg: r.faltaKg });

/* ---------- 3. sem previsao de bobina, so o quilo ---------- */

const SEM_BOB = [
  { ordem: 1, nome: 'Corpo', tecido: 'Malha Algodão', cor: 'Preto Malha Algodão', kg: 220, bobinas: 6 },
  { ordem: 2, nome: 'Ribana', tecido: 'Ribana Bulk', cor: 'Preto Ribana Bulk', kg: 85, bobinas: null }
];
r = rodar(SEM_BOB, [falta('Ribana Bulk', 'Preto Ribana Bulk', 40)]);
ok('tecido sem bobina prevista: faltaBob fica em branco, o quilo vale',
   r.itens[0].faltaBob === null && r.faltaKg === 40,
   { faltaBob: r.itens[0].faltaBob, faltaKg: r.faltaKg });
ok('e o aviso sabe que nao pode falar em bobina',
   r.temBobina === false, r.temBobina);

// Mas se o que falta É um tecido com bobina, a bobina volta a valer.
r = rodar(SEM_BOB, [falta('Malha Algodão', 'Preto Malha Algodão', 110)]);
ok('a mesma OS fala em bobina quando o que falta tem bobina prevista',
   r.temBobina === true && r.faltaBob === 3, { temBobina: r.temBobina, faltaBob: r.faltaBob });

/* ---------- 4. os casos em que a conta nao existe ---------- */

ok('OS sem falta nenhuma devolve zerado', rodar(FASES, []).previstoKg === 0);
ok('OS sem fase nenhuma devolve zerado', rodar([], [falta('X', 'Y', 10)]).previstoKg === 0);
// Tecido que a OS nao usa em fase alguma: o quilo conta, a bobina nao se inventa.
r = rodar(FASES, [falta('Moletom Bulk', 'Bege Moletom Bulk', 30)]);
ok('tecido fora das fases da OS: quilo sim, bobina nao',
   r.itens[0].faltaBob === null && r.faltaKg === 30,
   { faltaBob: r.itens[0].faltaBob, faltaKg: r.faltaKg });

console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
