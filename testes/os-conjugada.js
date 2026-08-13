/* Rode com:  node testes/os-conjugada.js

   OS conjugada: duas peças que saem do mesmo enfesto e sempre andam juntas.
   Salvar a OS da primeira gera sozinha a OS da segunda.

   Isto era UM caso escrito no código — camiseta bicolor puxava a básica, casando
   por NOME de grade e de desenho. Em 13/08/2026 virou campo do cadastro da
   grade, e o teste existe por dois motivos:

   1. É a única regra do programa que CRIA uma ordem de serviço sozinha. Um erro
      aqui não aparece na tela: aparece no chão, dias depois, em pano cortado
      para uma peça que ninguém pediu — ou na peça que faltou porque a segunda OS
      não nasceu.
   2. A migração é o ponto de perda calada. Se ela não levar a regra antiga para
      dentro do cadastro, a bicolor para de puxar a básica sem um aviso sequer.

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

// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF, e a quebra
// depois do fecha-chaves e '\r\n'.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';

const motor = [
  'function gradeConjugadaDaGrade',
  'function deveGerarConjugada',
  'async function gerarConjugada(',
  'async function aplicarRegraConjugadaSeAplicavel',
  // A cor da peça conjugada sai do desenho dela, pela ordem canonica do desc —
  // as mesmas funcoes que o formulario usa.
  'function ordemCoresPorDesc',
  'function ordenarCoresIdsPorDesc',
  'function _normNome',
  'function _sufixoTecidoNorm',
  'function corBaseNome'
].map(corta).join('\n');

// O motor mexe em STATE e chama a vizinhanca inteira do app; os dublês abaixo
// devolvem o que interessa e guardam o que foi chamado.
function comMotor(estado, numeros = [100, 101, 102]) {
  const avisos = [];
  const fila = numeros.slice();
  const fn = new Function('STATE', 'avisos', 'fila', `
    let _seq = 0;
    const uid = () => 'novo_' + (++_seq);
    const proximoNumeroOS = () => fila.shift();
    const saveState = async () => {};
    const atualizarCounterOS = async () => {};
    const toast = (msg, tipo) => { avisos.push([tipo || '', msg]); };
    ${motor}
    return { gradeConjugadaDaGrade, deveGerarConjugada, gerarConjugada, aplicarRegraConjugadaSeAplicavel };
  `);
  return { api: fn(estado, avisos, fila), avisos, estado };
}

// A migracao, com os arredores dublados.
const migSrc = corta('async function migrarRegraConjugadaParaGrade');
const regraSrc = recorte('const REGRA_BICOLOR_BASICA', '};', 'a constante da regra') + '};';
const normSrc = corta('function _normNome');
function rodarMigracao(estado, papel = 'admin') {
  const salvos = [];
  const fn = new Function('STATE', 'currentRole', 'salvos', `
    const saveState = async (k) => { salvos.push(k); };
    const toast = () => {};
    ${normSrc}
    ${regraSrc}
    ${migSrc}
    return migrarRegraConjugadaParaGrade();
  `);
  return fn(estado, papel, salvos).then(() => ({ estado, salvos }));
}

// A regra de esconder do seletor de grades da OS mora em gradesParaDropdownOS;
// aqui interessa so o pedaco que decide quem some.
function ocultas(grades) {
  const alvos = new Set(grades.map(g => g.conjugadaGradeId).filter(Boolean));
  return grades.filter(g => alvos.has(g.id) || /conjug/i.test(g.nome || '')).map(g => g.id);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// Uma dupla qualquer: a grade ATIVA (6 tamanhos) conjuga a PASSIVA (2 tamanhos).
const estadoBase = () => ({
  meta: {},
  tecidos: [{ id: 't1', nome: 'Malha Algodão' }],
  cores: [
    { id: 'c_preto', nome: 'Preto Malha Algodão' },
    { id: 'c_branco', nome: 'Branco Malha Algodão' }
  ],
  materiais: [],
  componentes: [],
  modelos: [
    { id: 'm_bi', nome: 'Camiseta Bicolor' },
    { id: 'm_ba', nome: 'Camiseta Básica' }
  ],
  desenhos: [
    { id: 'd_bi', codigo: 'Dx100', desc: 'Camiseta Bicolor | Preto/Branco', modeloId: 'm_bi',
      corPrincipalId: 'c_preto', corSecundariaId: 'c_branco' },
    { id: 'd_ba', codigo: 'Dx200', desc: 'Camiseta Básica | Branco', modeloId: 'm_ba',
      corPrincipalId: 'c_branco' }
  ],
  grades: [
    { id: 'g_ativa', nome: 'P-M-G-G1-G2-G3 | CM.BICOLOR',
      tamanhos: { p: 1, m: 2, g: 2, gg: 1, g1: 1, g2: 1, g3: 1 },
      fases: [{ ordem: 1, nome: 'Corpo', tecidoId: 't1', comp: '8', larg: '1.8' }] },
    { id: 'g_passiva', nome: 'M-G (CONJUGADO) | CM.BÁSICA',
      tamanhos: { m: 3, g: 2 },
      fases: [
        { ordem: 1, nome: 'Corpo', tecidoId: 't1', comp: '5', larg: '1.8' },
        { ordem: 2, nome: 'Gola', tecidoId: 't1', comp: '1', larg: '1.8' }
      ] },
    { id: 'g_solta', nome: 'P-M-G | CM.LISA', tamanhos: { p: 2, m: 2, g: 2 }, fases: [] }
  ],
  ordens: []
});

const osAtiva = () => ({
  id: 'os_1', os: '0435', desenhoId: 'd_bi', gradeId: 'g_ativa',
  data: '2026-08-13', colecao: 'Inverno', marca: 'Dixie', responsavel: 'Junior',
  modeloId: 'm_bi', modeloNome: 'Camiseta Bicolor',
  grade: { descricao: 'P-M-G-G1-G2-G3 | CM.BICOLOR', p: 1, m: 2, g: 2, gg: 1, g1: 1, g2: 1, g3: 1, total: 9 },
  enfesto: { comprimento: 8, largura: 1.8, camadas: 20, target: 180, totalPecas: 180 },
  // A identidade da peça ATIVA: e exatamente isto que nao pode vazar para a
  // conjugada. A barra em caixa alta acima do desenho tecnico le `variantes`.
  variantes: [{ num: 1, obs: '', cor1: 'c_preto', cor2: 'c_branco', cor3: '',
                cor1Nome: 'Preto Malha Algodão', cor2Nome: 'Branco Malha Algodão', cor3Nome: '—' }],
  tecidos: [
    { c1: '', tecidoId: 't1', tecidoNome: 'Malha Algodão', corId: 'c_preto', corNome: 'Preto Malha Algodão' },
    { c1: '', tecidoId: 't1', tecidoNome: 'Malha Algodão', corId: 'c_branco', corNome: 'Branco Malha Algodão' }
  ]
});

console.log('-- quem gera, e quem nao gera --');
{
  const e = estadoBase();
  e.grades[0].conjugadaGradeId = 'g_passiva';
  const { api } = comMotor(e);
  ok('1. grade com conjugada cadastrada gera', api.deveGerarConjugada(osAtiva()) === true);

  const semConj = estadoBase();
  const m2 = comMotor(semConj);
  ok('2. grade sem conjugada nao gera nada',
     m2.api.deveGerarConjugada(osAtiva()) === false);

  const filha = Object.assign(osAtiva(), { id: 'os_2', conjugadaPaiId: 'os_1' });
  ok('3. a propria conjugada nao puxa uma terceira', api.deveGerarConjugada(filha) === false);

  const jaTem = Object.assign(osAtiva(), { conjugadaId: 'os_9' });
  e.ordens = [{ id: 'os_9' }];
  ok('4. nao duplica se a conjugada dela ja existe', api.deveGerarConjugada(jaTem) === false);

  e.ordens = [];
  ok('5. se a conjugada foi apagada, gera de novo', api.deveGerarConjugada(jaTem) === true);
}

{
  // O ciclo: A conjuga B e B conjuga A. Sem a trava do conjugadaPaiId isto viraria
  // OS gerando OS ate o navegador morrer.
  const e = estadoBase();
  e.grades[0].conjugadaGradeId = 'g_passiva';
  e.grades[1].conjugadaGradeId = 'g_ativa';
  const { api } = comMotor(e);
  ok('6. par cruzado nao entra em loop: a segunda OS nasce marcada e para ai',
     api.deveGerarConjugada(Object.assign(osAtiva(), { conjugadaPaiId: 'os_1' })) === false);
}

{
  const e = estadoBase();
  e.grades[0].conjugadaGradeId = 'g_ativa';   // ela mesma
  const { api } = comMotor(e);
  ok('7. grade conjugada com ela mesma nao gera', api.deveGerarConjugada(osAtiva()) === false);

  const e2 = estadoBase();
  e2.grades[0].conjugadaGradeId = 'g_apagada';
  const m2 = comMotor(e2);
  ok('8. conjugada apagada nao gera, e RECLAMA (nao pode falhar calado)',
     m2.api.deveGerarConjugada(osAtiva()) === false && m2.avisos.some(a => a[0] === 'err'),
     m2.avisos);
}

console.log('');
console.log('-- a OS que sai --');
(async () => {
  {
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    e.grades[0].conjugadaDesenhoId = 'd_ba';
    e.ordens = [osAtiva()];
    const { api, estado } = comMotor(e);
    const nova = await api.gerarConjugada(estado.ordens[0]);

    ok('9. sai com numero proprio', nova.os === 100 && nova.id !== 'os_1', [nova.os, nova.id]);
    ok('10. vem na grade conjugada', nova.gradeId === 'g_passiva'
       && nova.grade.descricao === 'M-G (CONJUGADO) | CM.BÁSICA', nova.grade);
    ok('11. o total da grade e o da conjugada, nao o da ativa', nova.grade.total === 5, nova.grade.total);
    // 180 pecas-alvo / 2 (o MENOR tamanho da grade, G) = 90 camadas.
    ok('12. as camadas sao recalculadas pela grade dela', nova.enfesto.camadas === 90, nova.enfesto);
    ok('13. o alvo de pecas e o mesmo da ativa', nova.enfesto.target === 180, nova.enfesto.target);
    ok('14. as fases vem do cadastro da conjugada',
       nova.fases.length === 2 && nova.fases[0].comp === '5', nova.fases);
    ok('15. o contexto da ativa vem junto (data, colecao, marca, equipe)',
       nova.data === '2026-08-13' && nova.colecao === 'Inverno' && nova.marca === 'Dixie'
       && nova.responsavel === 'Junior', nova);
    ok('16. o desenho escolhido no cadastro manda',
       nova.desenhoId === 'd_ba' && nova.codigo === 'Dx200', [nova.desenhoId, nova.codigo]);
    ok('17. nasce marcada como filha', nova.conjugadaPaiId === 'os_1' && !nova.conjugadaId, nova.conjugadaPaiId);
    ok('18. a ativa guarda o vinculo', estado.ordens[0].conjugadaId === nova.id, estado.ordens[0].conjugadaId);
    ok('19. a nova entrou na lista', estado.ordens.length === 2 && estado.ordens[1].id === nova.id);
  }

  {
    // O DEFEITO DE 13/08/2026, que foi impresso e quase foi para o corte: a
    // conjugada branca saiu com "PRETO" na barra acima do desenho tecnico e
    // preto nas fases, porque variantes/tecidos/modelo vinham inteiros da
    // bicolor. A peca muda; so o contexto e que se herda.
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    e.grades[0].conjugadaDesenhoId = 'd_ba';       // Camiseta Básica | Branco
    e.ordens = [osAtiva()];
    const { api, estado } = comMotor(e);
    const nova = await api.gerarConjugada(estado.ordens[0]);

    ok('9b. a barra acima do desenho le a cor da PECA conjugada, nao a da ativa',
       nova.variantes.length === 1 && nova.variantes[0].cor1Nome === 'Branco Malha Algodão',
       nova.variantes);
    ok('9c. nenhuma cor da ativa sobra na variante',
       !JSON.stringify(nova.variantes).includes('Preto'), nova.variantes);
    ok('9d. as fases do enfesto saem na cor da conjugada',
       nova.fases.every(f => f.corNome === 'Branco Malha Algodão'),
       nova.fases.map(f => f.corNome));
    ok('9e. as linhas de tecido acompanham, uma por fase',
       nova.tecidos.length === 2 && nova.tecidos.every(t => t.corNome === 'Branco Malha Algodão'),
       nova.tecidos);
    ok('9f. a DESCRICAO da folha e o modelo da conjugada',
       nova.modeloNome === 'Camiseta Básica' && nova.modeloId === 'm_ba', nova.modeloNome);
  }

  {
    // Desenho em branco no cadastro: a segunda OS e a MESMA peca, so em outra
    // grade. E o caso comum de quem so quer duplicar o lote.
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    e.ordens = [osAtiva()];
    const { api, estado } = comMotor(e);
    const nova = await api.gerarConjugada(estado.ordens[0]);
    ok('20. sem desenho no cadastro, sai com o mesmo desenho da ativa',
       nova.desenhoId === 'd_bi' && nova.codigo === 'Dx100', [nova.desenhoId, nova.codigo]);
  }

  {
    // Desenho cadastrado que foi apagado depois: PARA. Sair com a peca errada
    // na ordem e pior que nao sair.
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    e.grades[0].conjugadaDesenhoId = 'd_sumiu';
    e.ordens = [osAtiva()];
    const m = comMotor(e);
    const nova = await m.api.gerarConjugada(m.estado.ordens[0]);
    ok('21. desenho apagado no cadastro: nao gera, e reclama',
       nova === null && m.avisos.some(a => a[0] === 'err'), m.avisos);
    ok('22. e nao suja a lista de OS', m.estado.ordens.length === 1);
  }

  {
    // Sem pecas-alvo, as camadas da ativa sao herdadas (nao ha o que recalcular).
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    const os = osAtiva();
    os.enfesto.target = 0;
    e.ordens = [os];
    const { api, estado } = comMotor(e);
    const nova = await api.gerarConjugada(estado.ordens[0]);
    ok('23. sem alvo de pecas, herda as camadas da ativa', nova.enfesto.camadas === 20, nova.enfesto.camadas);
  }

  {
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    e.grades[0].conjugadaDesenhoId = 'd_ba';
    e.ordens = [osAtiva()];
    const m = comMotor(e);
    const nova = await m.api.aplicarRegraConjugadaSeAplicavel(m.estado.ordens[0]);
    ok('24. o aviso diz QUAL peca saiu, nao so que saiu',
       nova && m.avisos.some(a => a[0] === 'ok' && /Camiseta Básica/.test(a[1])), m.avisos);
  }

  console.log('');
  console.log('-- o seletor de grades da OS --');
  {
    const e = estadoBase();
    e.grades[0].conjugadaGradeId = 'g_passiva';
    const some = ocultas(e.grades);
    ok('25. a grade conjugada some do seletor', some.includes('g_passiva'));
    ok('26. a ativa e as soltas continuam a vista',
       !some.includes('g_ativa') && !some.includes('g_solta'), some);
  }
  {
    // Sem nenhum cadastro de conjugada, o nome ainda esconde — e o que segurava
    // essas grades antes de o campo existir.
    const e = estadoBase();
    ok('27. o casamento por nome continua valendo como rede',
       ocultas(e.grades).includes('g_passiva'));
  }

  console.log('');
  console.log('-- a migracao: a regra do codigo entra no cadastro --');
  {
    const { estado, salvos } = await rodarMigracao(estadoBase());
    const g = estado.grades.find(x => x.id === 'g_ativa');
    ok('28. a bicolor passa a apontar a basica', g.conjugadaGradeId === 'g_passiva', g.conjugadaGradeId);
    ok('29. com o desenho da basica junto', g.conjugadaDesenhoId === 'd_ba', g.conjugadaDesenhoId);
    ok('30. gravou grades e meta', salvos.includes('grades') && salvos.includes('meta'), salvos);
    ok('31. marcou que ja rodou', estado.meta.conjugadaPorGradeV1 === true);
  }
  {
    // Rodar duas vezes nao pode mexer de novo: quem editou o cadastro no meio
    // teria a edicao desfeita.
    const e = estadoBase();
    const r1 = await rodarMigracao(e);
    r1.estado.grades[0].conjugadaGradeId = 'g_solta';   // o usuario mudou de ideia
    const r2 = await rodarMigracao(r1.estado);
    ok('32. a segunda passagem nao mexe em nada',
       r2.estado.grades[0].conjugadaGradeId === 'g_solta', r2.estado.grades[0].conjugadaGradeId);
  }
  {
    // Instalacao que nunca teve a dupla bicolor/basica: nao inventa conjugacao.
    const e = estadoBase();
    e.grades = [e.grades[2]];
    const { estado } = await rodarMigracao(e);
    ok('33. sem a dupla antiga, nao inventa nada',
       !estado.grades.some(g => g.conjugadaGradeId), estado.grades);
  }
  {
    const e = estadoBase();
    const { estado } = await rodarMigracao(e, 'consulta');
    ok('34. quem nao e admin nao migra (nao grava por cima do blob)',
       !estado.grades[0].conjugadaGradeId && !estado.meta.conjugadaPorGradeV1);
  }

  console.log('');
  console.log(falhas === 0 ? 'tudo certo' : falhas + ' falha(s)');
  process.exit(falhas === 0 ? 0 : 1);
})();
