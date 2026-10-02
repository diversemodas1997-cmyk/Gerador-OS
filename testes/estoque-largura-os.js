/* Rode com:  node testes/estoque-largura-os.js

   A RESERVA DA OS NA LARGURA DA GRADE (30/09/2026, Junior: "Coloque as reservas
   do algodão cru na largura de 80 cm, caso as OS sejam com grade de 80 cm").

   O que o teste guarda:
     · OS de grade 80 cm reservando algodão cru entra na linha de 80 cm;
     · a baixa dela também (a mesma regra de reserva/saída);
     · OS de 117 cm numa cor que não tem bobina de 117 fica nos 120 cm;
     · OS sem largura na fase usa a do nome da grade ("… | 80cm");
     · lançamento sem largura continua valendo 120 cm;
     · a soma das larguras é o total da cor.

   Recorta as funções do app.js de verdade. */
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
const cortaLinha = (nome) => recorte(nome, '\n', nome);

function porLargura(movs, ordens) {
  return new Function('MOVS', 'ORDENS', `
    const STATE = { ordens: ORDENS };
    const movimentacoesEstoque = () => MOVS;
    ${corta('function _normNome')}
    ${cortaLinha('const LARGURA_BOBINA_PADRAO_CM')}
    ${cortaLinha('const LARGURA_RIBANA_PADRAO_CM')}
    ${corta('function larguraPadraoDoTecido')}
    ${corta('function _larguraDaOSNoTecido')}
    ${corta('function estoquePorLargura')}
    return estoquePorLargura();
  `)(movs, ordens);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const CRU = { tecidoNome: 'Malha Algodão', corNome: 'Algodão cru' };
const PRETO = { tecidoNome: 'Malha Algodão', corNome: 'Preto' };
const ent = (base, largura, kg, bob) => Object.assign({ tipo: 'entrada', origem: 'manual', kg, fechados: bob, largura }, base);
const os = (base, osId, kg, status) => Object.assign({ tipo: 'saida', origem: 'os', osId, kg, status: status || 'reservado' }, base);

const movs = [
  ent(CRU, 119, 126, 7), ent(CRU, 80, 702, 54),
  os(CRU, 'os80', 30), os(CRU, 'os80b', 20, 'consumido'), os(CRU, 'osSemLargNaFase', 5),
  ent(PRETO, 120, 500, 25), os(PRETO, 'os117', 40),
  Object.assign({ tipo: 'entrada', origem: 'manual', kg: 10, fechados: 1 }, PRETO)   // sem largura
];
const ordens = [
  { id: 'os80', fases: [{ tecidoNome: 'Malha Algodão', larg: '0.80' }] },
  { id: 'os80b', fases: [{ tecidoNome: 'Malha Algodão', larg: '0,8' }] },
  { id: 'osSemLargNaFase', fases: [{ tecidoNome: 'Malha Algodão', larg: '' }], grade: { descricao: 'P ao G3 | CM.LISA | 80cm' } },
  { id: 'os117', fases: [{ tecidoNome: 'Malha Algodão', larg: '1.170' }] }
];
const m = porLargura(movs, ordens);
const cru = m.get('algodao cru') || m.get([...m.keys()].find(k => /cru/.test(k)));
const preto = m.get([...m.keys()].find(k => /preto/.test(k)));

ok('algodão cru: reserva da OS de 80 cm está na linha de 80 cm', cru.get(80).reservado === 35, cru.get(80));
ok('  ... e a baixa dela também', cru.get(80).saida === 20, cru.get(80));
ok('  ... OS sem largura na fase usa a do nome da grade (80cm)', cru.get(80).reservado === 35, cru.get(80));
ok('  ... nada foi para os 120 cm do algodão cru', !cru.has(120), [...cru.keys()]);
ok('  ... a de 119 cm não recebeu reserva', cru.get(119).reservado === 0, cru.get(119));
ok('preto: OS de 117 cm fica nos 120 (a cor não tem bobina de 117)', preto.get(120).reservado === 40 && !preto.has(117), [...preto.entries()]);
ok('preto: lançamento sem largura conta como 120 cm', preto.get(120).entrada === 510, preto.get(120));
const soma = (mp, k) => [...mp.values()].reduce((a, v) => a + v[k], 0);
ok('a soma das larguras é o total da cor (algodão cru)', soma(cru, 'entrada') === 828 && soma(cru, 'reservado') === 35 && soma(cru, 'saida') === 20);

// RIBANA É 60 CM (02/10/2026): sem largura, a ribana cai na linha de 60, não na de 120.
const RIB = { tecidoNome: 'Ribana Malha Algodão', corNome: 'Preto Ribana Malha Algodão' };
const mr = porLargura([
  ent(RIB, 60, 96, 12),
  Object.assign({ tipo: 'entrada', origem: 'manual', kg: 30 }, RIB),          // saldo inicial sem largura
  os(RIB, 'osRib', 7)
], [{ id: 'osRib', fases: [{ tecidoNome: 'Ribana Malha Algodão', larg: '' }] }]);
const rib = mr.get([...mr.keys()].find(k => /ribana/.test(k)));
ok('ribana: sem largura e OS caem na linha de 60 cm', rib.get(60).entrada === 126 && rib.get(60).reservado === 7, [...rib.entries()]);
ok('ribana: não abre linha de 120 cm', !rib.has(120), [...rib.keys()]);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
