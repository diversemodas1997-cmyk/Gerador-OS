/* Rode com:  node testes/excedente-por-fase.js

   Excedente de enfesto: mora na FASE, não no tecido.

   O comprimento do relatório do CAD é a medida de CORTAR. O que se cadastra na
   grade é a de ENFESTAR, que é maior: sobra pano nas duas pontas para a
   enfestadeira segurar e para o corte não morrer na borda.

   Essa sobra morava no TECIDO e mudou para a FASE em 12/08/2026, porque não é do
   pano: depende do COMPRIMENTO que a fase estende. Um corpo de 8 m e um viés de
   1 m saem do mesmo rolo e não levam a mesma ponta.

   Mudar um campo de lugar é onde se perde dado calado. Por isso a migração é
   testada aqui: ela tem de copiar o que estava no tecido para as fases daquele
   tecido, sem pisar em quem já tem valor próprio, e sem mudar nenhuma medida.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}

// As três funções da conta: constante, resolução e soma. Não dependem de STATE.
const contas = recorte('// EXCEDENTE DE ENFESTO.', 'let _riscoLeituras', 'a conta do excedente');
const api = new Function(contas + `
  return { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoM, excedenteEnfestoCm, _riscoCompCadastro };
`)();
const { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoM, excedenteEnfestoCm, _riscoCompCadastro } = api;

// A migração, com os arredores dublados: ela mexe em STATE e chama saveState.
const migSrc = recorte('async function migrarExcedenteParaFases', '\nfunction uid()', 'a migracao');
function rodarMigracao(estado, papel = 'admin') {
  const salvos = [];
  const fn = new Function('STATE', 'currentRole', 'salvos', `
    const saveState = async (k) => { salvos.push(k); };
    const toast = () => {};
    ${migSrc}
    return migrarExcedenteParaFases();
  `);
  return fn(estado, papel, salvos).then(() => ({ estado, salvos }));
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(a - b) < 1e-9;

console.log('-- a conta --');
ok('1. o padrao da casa e 15 cm', EXCEDENTE_ENFESTO_PADRAO_CM === 15, EXCEDENTE_ENFESTO_PADRAO_CM);
ok('2. fase sem excedente cadastrado cai no padrao',
   perto(excedenteEnfestoM({ comp: '4' }), 0.15), excedenteEnfestoM({ comp: '4' }));
ok('3. fase com excedente proprio usa o dela',
   perto(excedenteEnfestoM({ excedente: 5 }), 0.05), excedenteEnfestoM({ excedente: 5 }));
ok('4. ZERO cadastrado e zero de verdade, nao o padrao',
   perto(excedenteEnfestoM({ excedente: 0 }), 0), excedenteEnfestoM({ excedente: 0 }));
ok('5. campo em branco volta ao padrao',
   perto(excedenteEnfestoM({ excedente: '' }), 0.15), excedenteEnfestoM({ excedente: '' }));
ok('6. sem fase nenhuma, o padrao',
   perto(excedenteEnfestoM(null), 0.15) && perto(excedenteEnfestoM(undefined), 0.15));
ok('7. numero negativo nao encurta o enfesto',
   perto(excedenteEnfestoM({ excedente: -20 }), 0), excedenteEnfestoM({ excedente: -20 }));
ok('8. o rotulo em cm acompanha',
   excedenteEnfestoCm({ excedente: 5 }) === 5 && excedenteEnfestoCm({}) === 15);

console.log('');
console.log('-- a soma que vai para o cadastro --');
// O caso que fixou os 15 cm: "2X P ao G3 | BM.TRICOLOR", Corpo Parte 1.
ok('9. 4,5493 m do risco + 15 cm = 4,6993 m para cadastrar',
   perto(_riscoCompCadastro(4.5493, {}), 4.6993), _riscoCompCadastro(4.5493, {}));
ok('10. a mesma fase com 5 cm proprios da outra medida',
   perto(_riscoCompCadastro(4.5493, { excedente: 5 }), 4.5993),
   _riscoCompCadastro(4.5493, { excedente: 5 }));
ok('11. duas fases do MESMO tecido podem somar diferente',
   !perto(_riscoCompCadastro(8, { excedente: 30 }), _riscoCompCadastro(8, { excedente: 5 })),
   [_riscoCompCadastro(8, { excedente: 30 }), _riscoCompCadastro(8, { excedente: 5 })]);
ok('12. sem medida no PDF nao inventa medida', _riscoCompCadastro(null, {}) === null);

console.log('');
console.log('-- a migracao: o que estava no tecido vai para as fases --');
const estadoBase = () => ({
  meta: {},
  tecidos: [
    { id: 't_malha', nome: 'Malha Algodão' },                    // nunca cadastrado
    { id: 't_ribana', nome: 'Ribana Malha Algodão', excedente: 5 },
    { id: 't_mol', nome: 'Moletom', excedente: 15 },
    { id: 't_vazio', nome: 'Tactel', excedente: '' }             // em branco = padrao
  ],
  grades: [
    { id: 'g1', nome: 'A', fases: [
      { ordem: 1, nome: 'Corpo', tecidoId: 't_malha' },
      { ordem: 2, nome: 'Gola', tecidoId: 't_ribana' },
      { ordem: 3, nome: 'Viés', tecidoId: 't_malha', excedente: 0 }   // ja tem o dela
    ] },
    { id: 'g2', nome: 'B', fases: [
      { ordem: 1, nome: 'Corpo', tecidoId: 't_mol' },
      { ordem: 2, nome: 'Forro', tecidoId: 't_vazio' }
    ] }
  ]
});

(async () => {
  const { estado, salvos } = await rodarMigracao(estadoBase());
  const fase = (g, o) => estado.grades.find(x => x.id === g).fases.find(f => f.ordem === o);

  ok('13. a fase de ribana herda os 5 cm que estavam no tecido',
     fase('g1', 2).excedente === 5, fase('g1', 2));
  ok('14. a fase de moletom herda os 15 cm explicitos',
     fase('g2', 1).excedente === 15, fase('g2', 1));
  ok('15. tecido nunca cadastrado nao escreve nada — a fase fica no padrao',
     fase('g1', 1).excedente === undefined, fase('g1', 1));
  ok('16. tecido com campo em branco tambem nao escreve nada',
     fase('g2', 2).excedente === undefined, fase('g2', 2));
  ok('17. NAO pisa na fase que ja tinha valor proprio (o zero continua zero)',
     fase('g1', 3).excedente === 0, fase('g1', 3));
  ok('18. gravou as grades e a marca de ja ter rodado',
     salvos.includes('grades') && salvos.includes('meta'), salvos);
  ok('19. deixou a marca em meta', estado.meta.excedentePorFaseV1 === true, estado.meta);

  // O ponto da migracao: a medida NAO pode mudar sozinha.
  const antes = 2.51 + 5 / 100;                    // ribana, como era pelo tecido
  ok('20. a medida da ribana continua a mesma depois de mudar de lugar',
     perto(_riscoCompCadastro(2.51, fase('g1', 2)), antes),
     _riscoCompCadastro(2.51, fase('g1', 2)));

  // Roda de novo: nao pode mexer em nada.
  const { estado: e2, salvos: s2 } = await rodarMigracao(estado);
  ok('21. rodando de novo nao mexe em nada (a marca segura)',
     s2.length === 0 && e2.grades[0].fases[1].excedente === 5, s2);

  // Fase criada DEPOIS da migracao nao herda nada — vai para o padrao.
  estado.grades[0].fases.push({ ordem: 4, nome: 'Nova', tecidoId: 't_ribana' });
  const { estado: e3 } = await rodarMigracao(estado);
  ok('22. fase criada depois nao herda do tecido: nasce no padrao',
     e3.grades[0].fases[3].excedente === undefined, e3.grades[0].fases[3]);

  const { salvos: s4 } = await rodarMigracao(estadoBase(), 'leitura');
  ok('23. quem nao e admin nao migra nada', s4.length === 0, s4);

  console.log('');
  if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
  console.log('todos os testes passaram');
})();
