/* Rode com:  node testes/peso-da-bobina.js

   O KG ESTIMADO A PARTIR DAS BOBINAS (14/09/2026, Junior: "O usuário não tem
   como saber quantos kilos entra, ele sabe apenas quantas bobinas de tecido
   entra. A quantidade em kilos deve ser estimada pelo programa").

   Quem recebe a carga conta BOBINA. O quilo é o que o estoque precisa, e
   ninguém o sabe sem balança — então o programa faz a conta, e o kg entra no
   campo já preenchido, ainda editável.

   O que este teste protege:

     · a ordem das fontes: o CADASTRO do tecido ganha do histórico. Declarado
       vence deduzido, senão preencher o cadastro não teria efeito nenhum;

     · MEDIANA, e não média. Em 04/09 entraram 40 bobinas de 80 cm a 13 kg ao
       lado das normais de 19 kg. A média afundaria a estimativa de todas as
       cargas por causa de uma carga estreita;

     · entrada com ROLO ABERTO fica fora da amostra: o kg dela inclui pedaços de
       rolo, e dividir por bobinas cheias daria um peso por bobina inflado;

     · a LARGURA ajusta. Foi a própria casa que mostrou: 19 kg na bobina normal
       de Malha Algodão contra 13 kg na de 80 cm — 32% a menos. Estimar carga
       estreita pelo peso da normal erraria 45% para cima;

     · sem peso conhecido, NÃO INVENTA: devolve nada, e a tela pede o kg à mão.

   Recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  return src.slice(i, j + 2);
}
const monta = (STATE) => new Function('STATE', [
  recorte('function _normNome', 'a normalizacao de nome'),
  recorte('function pesoBobinaPorNome', 'o peso de bobina cadastrado'),
  recorte('function larguraPadraoTecido', 'a largura padrao do tecido'),
  recorte('function pesoBobinaEstimado', 'o peso de bobina estimado'),
  recorte('function estimativaKgEntrada', 'a estimativa de kg da entrada'),
  'return { pesoBobinaEstimado, estimativaKgEntrada, larguraPadraoTecido };'
].join('\n'))(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const MALHA = 'Malha Algodão';
const ent = (kg, fechados, abertos) => ({
  tipo: 'entrada', tecidoNome: MALHA, corNome: 'Preto', kg, fechados,
  abertos: abertos || 0, origem: 'manual'
});
const tecido = (extra) => Object.assign({ id: 't1', nome: MALHA }, extra || {});

console.log('-- de onde sai o peso da bobina --');
{
  // Nada cadastrado e nada no historico: nao inventa.
  let api = monta({ tecidos: [tecido()], estoqueMov: [] });
  ok('1. sem cadastro e sem historico, nao ha estimativa',
     api.pesoBobinaEstimado(MALHA) === null, api.pesoBobinaEstimado(MALHA));

  // So o cadastro.
  api = monta({ tecidos: [tecido({ pesoBobina: 22 })], estoqueMov: [] });
  ok('2. o cadastro do tecido responde', api.pesoBobinaEstimado(MALHA).kg === 22);
  ok('3. e diz que veio do cadastro', api.pesoBobinaEstimado(MALHA).origem === 'cadastro');

  // So o historico.
  api = monta({ tecidos: [tecido()], estoqueMov: [ent(190, 10), ent(325, 17)] });
  const h = api.pesoBobinaEstimado(MALHA);
  ok('4. sem cadastro, o historico responde', h && h.origem === 'historico', h);
  ok('5. e conta quantos lancamentos usou', h.n === 2, h);

  // O cadastro GANHA do historico: declarado vence deduzido.
  api = monta({ tecidos: [tecido({ pesoBobina: 22 })], estoqueMov: [ent(190, 10), ent(325, 17)] });
  ok('6. havendo cadastro, ele ganha do historico',
     api.pesoBobinaEstimado(MALHA).kg === 22, api.pesoBobinaEstimado(MALHA));
}

console.log('');
console.log('-- a carga fora de esquadro nao afunda a estimativa --');
{
  // Os numeros REAIS de 04/09 e 14/09: quatro cargas normais e uma de 80 cm.
  const reais = [ent(252, 14), ent(520, 40), ent(190, 10), ent(325, 17), ent(362, 19)];
  const api = monta({ tecidos: [tecido()], estoqueMov: reais });
  const p = api.pesoBobinaEstimado(MALHA);
  // pesos: 13,00 | 18,00 | 19,00 | 19,05 | 19,12  -> mediana 19,00
  ok('7. a mediana devolve a bobina NORMAL (19 kg), nao a media (17,6)',
     p.kg === 19, p);

  // Entrada com rolo aberto fica de fora: o kg dela tem pedaco de rolo dentro.
  const comAberto = monta({ tecidos: [tecido()], estoqueMov: [ent(190, 10), ent(684, 33, 5)] });
  ok('8. entrada com rolo ABERTO nao entra na amostra',
     comAberto.pesoBobinaEstimado(MALHA).n === 1, comAberto.pesoBobinaEstimado(MALHA));

  // Entrada sem contagem de bobinas tambem nao serve de amostra.
  const semCont = monta({ tecidos: [tecido()], estoqueMov: [ent(190, 10), ent(8846, 0)] });
  ok('9. entrada sem contagem de bobinas nao entra na amostra',
     semCont.pesoBobinaEstimado(MALHA).n === 1, semCont.pesoBobinaEstimado(MALHA));
}

console.log('');
console.log('-- o kg da carga --');
{
  const api = monta({ tecidos: [tecido({ pesoBobina: 19, largura: 117 })], estoqueMov: [] });
  ok('10. 10 bobinas de 19 kg dao 190 kg',
     api.estimativaKgEntrada(MALHA, 10, '').kg === 190, api.estimativaKgEntrada(MALHA, 10, ''));
  ok('11. sem bobina nenhuma nao ha o que estimar',
     api.estimativaKgEntrada(MALHA, 0, '') === null);

  // A largura de 80 cm contra a padrao de 117: 19 x (80/117) = 12,99 kg.
  const estreita = api.estimativaKgEntrada(MALHA, 40, 80);
  ok('12. bobina estreita pesa menos, na razao das larguras',
     Math.abs(estreita.kg - 40 * 19 * (80 / 117)) < 0.01, estreita);
  ok('13. ... e o resultado bate com os 520 kg que a casa lancou em 04/09',
     Math.abs(estreita.kg - 520) < 12, estreita.kg);
  ok('14. a largura igual a padrao nao ajusta nada',
     api.estimativaKgEntrada(MALHA, 10, 117).fator === 1);

  // Sem largura cadastrada no tecido nao ha com o que comparar: nao ajusta.
  const semLarg = monta({ tecidos: [tecido({ pesoBobina: 19 })], estoqueMov: [] });
  ok('15. sem largura de ficha tecnica, nao inventa ajuste',
     semLarg.estimativaKgEntrada(MALHA, 10, 80).fator === 1,
     semLarg.estimativaKgEntrada(MALHA, 10, 80));

  // Tecido sem peso conhecido: a tela pede o kg a mao.
  const semPeso = monta({ tecidos: [tecido()], estoqueMov: [] });
  ok('16. tecido sem peso conhecido nao produz estimativa',
     semPeso.estimativaKgEntrada(MALHA, 10, '') === null);
}

console.log('');
console.log(falhas ? falhas + ' FALHA(S)' : 'todos os testes passaram');
process.exit(falhas ? 1 : 0);
