/* O PAPEL E O FILME DO ENFESTO SAEM COM A OS (28/09/2026, Junior: "Faça os
   ajustes necessários para que o programa faça as baixas desses materiais
   junto com cada OS ... de acordo com as dimensões de comprimento do enfesto";
   "todas as fases, exceto viés. Mas, a baixa da fase gola só deve acontecer
   caso o checkbox na fase enfesto estiver preenchido"; "sempre uma camada de
   papel e filme"; "papel e filme em metros"; "igual ao tecido").

   As contas sao recortadas do app.js; o consumo do enfesto de cada OS e de
   mentira (o que importa aqui e o que se faz com ele). */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const linha = re => (src.match(re) || [''])[0];

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// O enfesto de mentira de cada OS: a mesma forma que consumoEnfestoOS devolve.
const ENFESTO = {
  a: [
    { ordem: 1, faseNome: 'Corpo', comp: 6.2, camadas: 36, viesPuro: false },
    { ordem: 2, faseNome: 'Ribana', comp: 2.5, camadas: 18, viesPuro: false },
    { ordem: 3, faseNome: 'Gola', comp: 1.3, camadas: 10, viesPuro: false },
    { ordem: 4, faseNome: 'Viés', comp: 3, camadas: 1, viesPuro: true }
  ],
  b: [
    { ordem: 1, faseNome: 'Corpo + Gola', comp: 4, camadas: 20, viesPuro: false },
    { ordem: 2, faseNome: 'Gola', comp: 1, camadas: 8, viesPuro: false },
    { ordem: 3, faseNome: 'Forro do capuz', comp: 2, camadas: 0, viesPuro: false }   // tom 0: nao enfestada
  ]
};
const STATE = {
  materiaisEstCad: [
    { id: 'kraft', nome: 'Papel kraft aerado', unidade: 'desc', medida: 'm', baixaOS: true, emEstoque: 1125 },
    { id: 'filme', nome: 'Filme plástico', unidade: 'desc', medida: 'm', baixaOS: true, emEstoque: 1600 },
    { id: 'fita', nome: 'Fita adesiva', unidade: 'desc', emEstoque: 30 }
  ],
  ordens: [
    // nao iniciada: reserva (a gola nao marcada fica de fora)
    { id: 'a', os: '0600', status: 'nao-iniciado' },
    // enfestando desde 29/09, gola marcada: baixa
    { id: 'b', os: '0601', status: 'enfestando',
      progresso: { enfestosCheck: { 2: true }, enfestosSeq: { 1: Date.parse('2026-09-29T10:00:00'), 2: Date.parse('2026-09-29T11:00:00') } } },
    // andou antes da contagem: fica de fora
    { id: 'c', os: '0590', status: 'cortando', progresso: { enfestosSeq: { 1: Date.parse('2026-09-20T10:00:00') } } },
    // cancelada: fica de fora
    { id: 'd', os: '0602', status: 'cancelado' },
    // conjugada passiva: e o mesmo enfesto da ativa
    { id: 'e', os: '0603', status: 'nao-iniciado', conjugadaPaiId: 'a' }
  ]
};
ENFESTO.c = ENFESTO.a; ENFESTO.d = ENFESTO.a; ENFESTO.e = ENFESTO.a;

const api = new Function('STATE', 'ENFESTO', `
  ${pegaFuncao('_normNome')}
  ${pegaFuncao('_normFaseNome')}
  ${linha(/const _EXC_LIGACAO = [^\n]*/)}
  ${pegaFuncao('_faseSoDe')}
  ${linha(/const _PAL_GOLA = [^\n]*/)}
  ${pegaFuncao('_aviDiaDoCarimbo')}
  ${linha(/const _aviUnidadeDe = [^\n]*/)}
  ${linha(/const _estItemMedida = [^\n]*/)}
  ${linha(/const MAT_BAIXA_DESDE = [^\n]*/)}
  const consumoEnfestoOS = o => ENFESTO[o.id] || [];
  const _statusOS = o => o.status;
  const _STATUS_QUE_BAIXAM = ['enfestando', 'cortando', 'separando', 'costurando'];
  ${pegaFuncao('_matFasesEnfestoOS')}
  ${pegaFuncao('_matDataBaixaOS')}
  ${pegaFuncao('_matDasOS')}
  ${pegaFuncao('_matFaltas')}
  return { _matFasesEnfestoOS, _matDasOS, _matFaltas };
`)(STATE, ENFESTO);

console.log('-- as fases que gastam --');
const fa = api._matFasesEnfestoOS(STATE.ordens[0]);
ok('1. todas as fases menos o vies; a gola sem o checkbox fica de fora', fa.map(f => f.fase).join(',') === 'Corpo,Ribana', fa);
const fb = api._matFasesEnfestoOS(STATE.ordens[1]);
ok('2. gola marcada entra; "Corpo + Gola" e corpo e entra sempre; fase com 0 camadas nao', fb.map(f => f.fase).join(',') === 'Corpo + Gola,Gola', fb);
ok('3. a conjugada passiva nao gasta', api._matFasesEnfestoOS(STATE.ordens[4]).length === 0);

console.log('-- reserva e baixa --');
const r = api._matDasOS();
const baixa = (item, os) => r.baixas.filter(b => b.itemId === item && b.osNumero === os).reduce((a, b) => a + b.qtd, 0);
const reserva = (item, os) => r.reservas.filter(b => b.itemId === item && b.osNumero === os).reduce((a, b) => a + b.qtd, 0);
ok('4. OS nao iniciada reserva 1x o comprimento (6,2 + 2,5 = 8,7 m) de papel e de filme',
   reserva('kraft', '0600') === 8.7 && reserva('filme', '0600') === 8.7, r.reservas);
ok('5. OS andando baixa (4 + 1 = 5 m) de cada, na data da primeira marca',
   baixa('kraft', '0601') === 5 && baixa('filme', '0601') === 5 && r.baixas.find(b => b.osNumero === '0601').data === '2026-09-29', r.baixas);
ok('6. OS que andou antes da contagem nao baixa', baixa('kraft', '0590') === 0);
ok('7. OS cancelada nao conta', baixa('kraft', '0602') === 0 && reserva('kraft', '0602') === 0);
ok('8. material sem "baixa automatica" nao e tocado', !r.baixas.concat(r.reservas).some(b => b.itemId === 'fita'));
ok('9. a baixa entra no historico como saida do estoque, lida da OS',
   r.baixas.every(b => b.auto && b.tipo === 'saida' && b.dEstoque === -b.qtd && b.unidade === 'desc'));

console.log('-- a falta prevista --');
ok('10. com estoque de sobra, nenhuma falta', api._matFaltas().length === 0, api._matFaltas());
STATE.materiaisEstCad[1].emEstoque = 10;   // filme: 10 - 5 baixados = 5 em estoque, 8,7 reservados
const fl = api._matFaltas();
ok('11. reservado alem do estoque vira falta (8,7 - 5 = 3,7 m), com as OS que reservam',
   fl.length === 1 && fl[0].itemId === 'filme' && fl[0].falta === 3.7 && fl[0].estoque === 5 && fl[0].os.join() === '0600', fl);
STATE.materiaisEstCad[1].emEstoque = 2;    // estoque negativo (-3): falta e o reservado inteiro
ok('12. estoque ja negativo: falta e o reservado inteiro, sem somar o negativo', api._matFaltas()[0].falta === 8.7, api._matFaltas());

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
