/* Rode com:  node testes/compra-necessidade-bruta.js

   COMPRA: quanto pano uma produção pede ANTES de existir OS.

   A tela de Compra não tem conta própria. Ela monta, da GRADE e do DESENHO
   TÉCNICO, o mesmo objeto que uma OS salva teria, e entrega para as funções que
   já existiam — `consumoEnfestoOS` e `bobinasEfetivasFase`. É esse o ponto que
   este teste guarda: se um dia a conta da compra divergir da conta da folha de
   OS, é porque alguém escreveu uma segunda conta, e o pedido ao fornecedor vai
   sair de um número que a produção não reconhece.

   O que se verifica aqui:

     • a grade + o desenho resolvem tecido, cor e camadas de cada fase — a
       ribana enfestando pela regra dela, não pelo número do moletom;
     • a bobina SOMA ANTES DE ARREDONDAR. Na folha de OS cada fase arredonda
       para cima, porque quem separa material tira bobina inteira da prateleira.
       Na compra, duas produções de 6,3 bobinas são 13 e não 14 — arredondar
       cada pedaço compraria pano a mais a cada linha da lista;
     • o quilo que falta comprar desconta o que está disponível no estoque;
     • o kg/bobina é RESULTADO da conta, nunca premissa (mesmo raciocínio da
       folha de OS), e é por ele que o quilo a comprar volta a virar bobina;
     • fase sem medida ou sem gramatura sai ZERADA e não inventa número.

   O teste recorta as funções do app.js de verdade. Copiar as fórmulas para cá
   testaria a cópia, que é exatamente como um defeito sobrevive em dois lugares. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

// Recorta uma função de topo pelo nome: do `function nome(` até a chave que
// fecha na coluna zero. Vale porque o app.js escreve todas assim.
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome + ' no app.js'); process.exit(1); }
  const fim = src.indexOf('\n}', ini);
  if (fim < 0) { console.error('nao achei o fim de ' + nome); process.exit(1); }
  return src.slice(ini, fim + 2);
}

// Recorta uma constante de uma linha só.
function pegaConst(nome) {
  const re = new RegExp('^const ' + nome + ' = .*$', 'm');
  const m = src.match(re);
  if (!m) { console.error('nao achei a constante ' + nome + ' no app.js'); process.exit(1); }
  return m[0];
}

const CONSTS = ['LIMITE_CAMADAS', 'MULTIPLICADOR_PECAS', 'LABEL_CATEGORIA',
                'UNIDADES_PADRAO_FORRO', 'CAMADAS_REF_BOBINAS_CADASTRO',
                'CEIL_BOBINA_EPS', '_COMPRA_TAMS'];

const FUNCOES = [
  // nomes, cores e tecidos
  '_normNome', '_sufixoTecidoNorm', 'corBaseNome', 'corCanonicaPorTecido',
  'ordemCoresPorDesc', 'ordenarCoresIdsPorDesc',
  'categoriaEfetivaTecido', 'isTecidoRibana', 'calcularPapeisFases',
  'coresPorFaseDaGrade',
  // camadas
  '_tamanhoQueMandaNaGrade', 'camadasDaFaseRibana', '_ribanaEscalaComGrade',
  'camadasPadraoDaFase', 'multiplicadorPecaOS', '_faseNaoEnfestadaPorTom',
  // consumo
  'gramaturaTecidoPorNome', 'pesoBobinaPorNome', 'gramaturaCorPorNome',
  'consumoEnfestoOS',
  // bobinas
  'parseBobinas', 'bobinaInteira', 'ehFaseRibana', 'bobinasEfetivasFase',
  // estoque (o desconto do disponível)
  'comprasComoMovimentos', 'movimentacoesEstoque', 'calcularSaldosEstoque',
  // a compra
  'compraLimiteCamadasGrade', 'compraOsSimulada', '_compraPecasPorCamada',
  'compraPecasDeCamadas', 'compraCamadasDePecas', 'compraConsumoItem',
  'compraNecessidadeBruta'
];

// `new Function` em vez de `eval`: dentro de eval um `const` não vaza para o
// escopo de fora, e boa parte do que se recorta aqui é const.
const api = new Function(
  'var STATE = {}; var comprasCache = [];\n'
  + CONSTS.map(pegaConst).join('\n') + '\n'
  + FUNCOES.map(pegaFuncao).join('\n') + '\n'
  + 'return { setState: s => { STATE = s; }, '
  + FUNCOES.filter(n => n.indexOf('compra') === 0 || n === 'consumoEnfestoOS'
                        || n === 'compraNecessidadeBruta').join(', ')
  + ' };'
)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + extra));
  if (!cond) falhas++;
};
const perto = (a, b, tol) => Math.abs(Number(a) - Number(b)) <= (tol == null ? 0.001 : tol);

/* --------------------------------------------------------------------------
   O cadastro de mentira, com a forma do de verdade: moletom com ribana de
   moletom (a que escala com a grade), gramatura cadastrada POR COR e a
   previsão de bobinas no cadastro da grade.
   -------------------------------------------------------------------------- */
const T_MOL = 't-mol', T_RIB = 't-rib', T_MALHA = 't-malha';
const C_MOL = 'c-mol', C_RIB = 'c-rib', C_MALHA = 'c-malha';
const G1 = 'g-1', D1 = 'd-1';

function cadastroBase() {
  return {
    tecidos: [
      { id: T_MOL,   nome: 'Moletom',        categoria: 'moletom' },
      // Ribana não tem categoria no cadastro: é reconhecida pelo nome.
      { id: T_RIB,   nome: 'Ribana Moletom', categoria: '', pesoBobina: 20 },
      { id: T_MALHA, nome: 'Malha Algodão',  categoria: 'malha' }
    ],
    cores: [
      { id: C_MOL,   nome: 'Preto Moletom',        peso: 500 },
      { id: C_RIB,   nome: 'Preto Ribana Moletom', peso: 400 },
      { id: C_MALHA, nome: 'Branco Malha Algodão', peso: 452 }
    ],
    componentes: [],
    desenhos: [{
      id: D1, codigo: 'BM.TESTE', desc: 'Blusa moletom | Preto',
      corPrincipalId: C_MOL,
      componentes: [
        { nome: 'Frente', corId: C_MOL, tecidoId: T_MOL },
        { nome: 'Punho',  corId: C_RIB, tecidoId: T_RIB }
      ]
    }],
    grades: [{
      id: G1, nome: 'M-2G-GG | BM.TESTE',
      tamanhos: { p: 0, m: 1, g: 2, gg: 1, g1: 0, g2: 0, g3: 0 },
      fases: [
        { ordem: 1, nome: 'Moletom', tecidoId: T_MOL, comp: 8,   larg: 1.8, bobinas: '14' },
        { ordem: 2, nome: 'Punhos',  tecidoId: T_RIB, comp: 2,   larg: 1.0, unidades: 2, bobinas: '1' }
      ]
    }],
    ordens: [],
    estoqueMov: []
  };
}

api.setState(cadastroBase());

/* ---------- 1. a grade + o desenho montam a OS que ainda não existe ---------- */
console.log('\n-- a grade e o desenho resolvem as fases --');

const o = api.compraOsSimulada(G1, D1, 36);
ok('1. a grade vira duas fases', o && o.fases.length === 2, o && o.fases.length);
ok('2. o tecido de cada fase vem da grade',
  o.fases[0].tecidoNome === 'Moletom' && o.fases[1].tecidoNome === 'Ribana Moletom',
  o.fases.map(f => f.tecidoNome).join(' / '));
ok('3. a cor do corpo vem do desenho (cor principal)',
  o.fases[0].corNome === 'Preto Moletom', o.fases[0].corNome);
ok('4. a cor do punho vem do COMPONENTE do desenho, não da cor do corpo',
  o.fases[1].corNome === 'Preto Ribana Moletom', o.fases[1].corNome);
ok('5. os blocos do enfesto nascem SEM camadas — cada fase deriva a sua',
  o.enfesto.blocos.every(b => b.camadas === undefined),
  JSON.stringify(o.enfesto.blocos.map(b => b.camadas)));

/* ---------- 2. as camadas de cada fase ---------- */
console.log('\n-- cada fase enfesta o que é dela --');

const cons = api.consumoEnfestoOS(o);
ok('6. o corpo enfesta as camadas informadas', cons[0].camadas === 36, cons[0].camadas);
// Ribana de moletom escala com a grade: o tamanho que manda é o G (2 por grade),
// e a fase rende 2 unidades por camada -> 36 x 2 / 2 = 36.
ok('7. a ribana enfesta pela regra DELA, não pelo número do moletom',
  cons[1].camadas === 36, cons[1].camadas);

/* ---------- 3. o consumo de um item ---------- */
console.log('\n-- quilos e bobinas de uma produção --');

const item1 = { id: 'i1', gradeId: G1, desenhoId: D1, camadas: 36, repeticoes: 1 };
const c1 = api.compraConsumoItem(item1);

// 8 m x 1,80 m x 36 camadas x 500 g/m2 / 1000 = 259,2 kg
ok('8. o quilo do corpo é comprimento x largura x camadas x gramatura',
  perto(c1.linhas[0].kgTotal, 259.2), c1.linhas[0].kgTotal);
// 14 bobinas para o enfesto cheio de 80 camadas, e esta OS tem 36: 14 x 36/80 = 6,3
ok('9. a bobina do corpo vem do cadastro da grade, na proporção das camadas',
  perto(c1.linhas[0].bobinas, 6.3), c1.linhas[0].bobinas);
ok('10. e vem CRUA, sem arredondar (quem arredonda é o total)',
  c1.linhas[0].bobinas !== 7, c1.linhas[0].bobinas);
// 2 m x 1,00 m x 36 x 400 / 1000 = 28,8 kg ; bobina de ribana = 28,8 / 20 = 1,44
ok('11. a ribana pesa pela gramatura DELA', perto(c1.linhas[1].kgTotal, 28.8), c1.linhas[1].kgTotal);
ok('12. e a bobina dela sai do peso, não do cadastro da grade',
  perto(c1.linhas[1].bobinas, 1.44), c1.linhas[1].bobinas);
// grade M=1, G=2, GG=1 -> menor pedido 1 ; moletom não corta em camada dupla
ok('13. as peças saem do menor pedido da grade x camadas', c1.pecas === 36, c1.pecas);

/* ---------- 4. o total: soma antes de arredondar ---------- */
console.log('\n-- a bobina soma antes de arredondar --');

const t1 = api.compraNecessidadeBruta([item1]);
const mol1 = t1.find(x => x.tecidoNome === 'Moletom');
const rib1 = t1.find(x => x.tecidoNome === 'Ribana Moletom');
ok('14. uma produção: 6,3 bobinas viram 7', mol1.bobinas === 7, mol1.bobinas);
ok('15. e 1,44 bobina de ribana vira 2', rib1.bobinas === 2, rib1.bobinas);

// DUAS produções iguais: 6,3 + 6,3 = 12,6 -> 13. Arredondando cada uma antes de
// somar dariam 14, e a casa compraria uma bobina a mais por engano de conta.
const item2 = { id: 'i2', gradeId: G1, desenhoId: D1, camadas: 36, repeticoes: 1 };
const t2 = api.compraNecessidadeBruta([item1, item2]);
const mol2 = t2.find(x => x.tecidoNome === 'Moletom');
ok('16. duas produções de 6,3 bobinas são 13, e não 14', mol2.bobinas === 13, mol2.bobinas);
ok('17. o quilo dobra junto', perto(mol2.kg, 518.4), mol2.kg);

// O mesmo pelo campo "enfestos iguais".
const t3 = api.compraNecessidadeBruta([{ ...item1, repeticoes: 2 }]);
const mol3 = t3.find(x => x.tecidoNome === 'Moletom');
ok('18. dois enfestos iguais dão o mesmo que dois itens', mol3.bobinas === 13, mol3.bobinas);

/* ---------- 5. kg por bobina: resultado, nunca premissa ---------- */
console.log('\n-- o peso da bobina sai da conta --');

// 259,2 kg / 6,3 bobinas = 41,14 kg. Moletom é pano pesado; o número serve de
// conferência, e não entra em conta nenhuma.
ok('19. o kg/bobina é o quilo dividido pelas bobinas CRUAS',
  perto(mol1.kgPorBobina, 259.2 / 6.3, 0.01), mol1.kgPorBobina);

/* ---------- 6. o que falta comprar desconta o estoque ---------- */
console.log('\n-- a comprar = bruto - disponível --');

const comEstoque = cadastroBase();
comEstoque.estoqueMov = [
  { id: 'e1', tipo: 'entrada', tecidoNome: 'Moletom', corNome: 'Preto Moletom', kg: 100, data: '2026-08-01' }
];
api.setState(comEstoque);
const t4 = api.compraNecessidadeBruta([item1]);
const mol4 = t4.find(x => x.tecidoNome === 'Moletom');
ok('20. o disponível aparece', perto(mol4.disponivel, 100), mol4.disponivel);
ok('21. e sai do bruto', perto(mol4.kgComprar, 159.2), mol4.kgComprar);
// 159,2 kg a comprar / 41,14 kg por bobina = 3,87 -> 4 bobinas
ok('22. o quilo a comprar volta a virar bobina pelo kg/bobina da própria conta',
  mol4.bobinasComprar === 4, mol4.bobinasComprar);

// Estoque de sobra: não se compra negativo.
const sobrando = cadastroBase();
sobrando.estoqueMov = [
  { id: 'e2', tipo: 'entrada', tecidoNome: 'Moletom', corNome: 'Preto Moletom', kg: 9999, data: '2026-08-01' }
];
api.setState(sobrando);
const t5 = api.compraNecessidadeBruta([item1]);
const mol5 = t5.find(x => x.tecidoNome === 'Moletom');
ok('23. com estoque de sobra, não há o que comprar', mol5.kgComprar === 0 && mol5.bobinasComprar === 0,
  mol5.kgComprar + ' kg / ' + mol5.bobinasComprar + ' bob');
ok('24. mas o bruto continua sendo o bruto', perto(mol5.kg, 259.2), mol5.kg);

/* ---------- 7. cadastro incompleto não vira número inventado ---------- */
console.log('\n-- o que falta cadastrar aparece, e não vira chute --');

api.setState(cadastroBase());

const semMedida = cadastroBase();
semMedida.grades[0].fases[0].comp = '';
api.setState(semMedida);
const t6 = api.compraNecessidadeBruta([item1]);
const mol6 = t6.find(x => x.tecidoNome === 'Moletom');
ok('25. fase sem comprimento sai com zero quilo (e não com um quilo inventado)',
  !mol6 || mol6.kg === 0, mol6 && mol6.kg);

const semGramatura = cadastroBase();
semGramatura.cores[0].peso = 0;
api.setState(semGramatura);
const t7 = api.compraNecessidadeBruta([item1]);
const mol7 = t7.find(x => x.tecidoNome === 'Moletom');
ok('26. cor sem gramatura também sai zerada', !mol7 || mol7.kg === 0, mol7 && mol7.kg);

// Grade sem previsão de bobinas: o quilo continua, a bobina não se inventa.
const semBobinas = cadastroBase();
semBobinas.grades[0].fases[0].bobinas = '';
semBobinas.grades[0].fases[1].bobinas = '';
api.setState(semBobinas);
const t8 = api.compraNecessidadeBruta([item1]);
const mol8 = t8.find(x => x.tecidoNome === 'Moletom');
ok('27. sem previsão no cadastro da grade, a bobina fica em branco', mol8.bobinas === null, mol8.bobinas);
ok('28. mas o quilo continua lá', perto(mol8.kg, 259.2), mol8.kg);
ok('29. e a tela sabe DIZER qual fase ficou sem previsão',
  mol8.semPrevisao.length === 1, JSON.stringify(mol8.semPrevisao));

// Ribana sem o peso médio da bobina cadastrado: espera, não chuta.
const ribSemPeso = cadastroBase();
ribSemPeso.tecidos[1].pesoBobina = 0;
api.setState(ribSemPeso);
const t9 = api.compraNecessidadeBruta([item1]);
const rib9 = t9.find(x => x.tecidoNome === 'Ribana Moletom');
ok('30. ribana sem peso de bobina cadastrado não prevê bobina', rib9.bobinas === null, rib9.bobinas);
ok('31. e o quilo dela segue contando para a compra', perto(rib9.kg, 28.8), rib9.kg);

/* ---------- 8. camadas e peças, os dois lados da mesma medida ---------- */
console.log('\n-- camadas e peças --');

api.setState(cadastroBase());
const oo = api.compraOsSimulada(G1, D1, 1);
ok('32. o limite de camadas da grade é o do tecido mais restritivo (moletom 36)',
  api.compraLimiteCamadasGrade(cadastroBase().grades[0]) === 36,
  api.compraLimiteCamadasGrade(cadastroBase().grades[0]));
ok('33. camadas -> peças', api.compraPecasDeCamadas(oo, 36) === 36, api.compraPecasDeCamadas(oo, 36));
ok('34. peças -> camadas', api.compraCamadasDePecas(oo, 36) === 36, api.compraCamadasDePecas(oo, 36));
ok('35. peças que não fecham camada arredondam para CIMA (nunca produzir a menos)',
  api.compraCamadasDePecas(oo, 37) === 37, api.compraCamadasDePecas(oo, 37));

// Camiseta: malha sem moletom corta em camada dupla, então cada camada rende 2.
const camiseta = cadastroBase();
camiseta.grades.push({
  id: 'g-2', nome: 'P-M-G | CM.TESTE',
  tamanhos: { p: 1, m: 1, g: 1, gg: 0, g1: 0, g2: 0, g3: 0 },
  fases: [{ ordem: 1, nome: 'Malha', tecidoId: T_MALHA, comp: 6, larg: 1.17, bobinas: '10' }]
});
camiseta.desenhos.push({ id: 'd-2', codigo: 'CM.TESTE', desc: 'Camiseta | Branco',
                         corPrincipalId: C_MALHA, componentes: [] });
api.setState(camiseta);
const oc = api.compraOsSimulada('g-2', 'd-2', 80);
ok('36. camiseta corta em camada dupla: 80 camadas dão 160 peças',
  api.compraPecasDeCamadas(oc, 80) === 160, api.compraPecasDeCamadas(oc, 80));
ok('37. e o limite dela é o da malha (80)',
  api.compraLimiteCamadasGrade(camiseta.grades[1]) === 80,
  api.compraLimiteCamadasGrade(camiseta.grades[1]));

/* ---------- 9. grade que sumiu do cadastro ---------- */
console.log('\n-- item órfão não derruba a tela --');

api.setState(cadastroBase());
ok('38. item apontando para grade excluída não calcula nada (e não estoura)',
  api.compraConsumoItem({ id: 'x', gradeId: 'nao-existe', camadas: 10 }) === null);
ok('39. e o total ignora esse item em vez de quebrar',
  api.compraNecessidadeBruta([{ id: 'x', gradeId: 'nao-existe', camadas: 10 }]).length === 0);

console.log('');
if (falhas) { console.error(falhas + ' teste(s) falharam'); process.exit(1); }
console.log('todos os testes passaram');
