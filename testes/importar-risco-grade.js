/* Rode com:  node testes/importar-risco-grade.js

   Importar risco (PDF do encaixe): o seletor de GRADE.

   O programa descobre a grade pela tabela de tamanhos do relatório do CAD. Isso
   acerta na maioria das vezes e ERRA quando o encaixe não é uma grade inteira:

     CM.REC - CORPO 2 - M-G-GG-G1-G2-G3
       Encaixados: 14/30 · Modelos completos: 2 · Modelos pedidos: 6
       M  completos 1  moldes 5      GG  completos 0  moldes 1
       G  completos 1  moldes 5      G1..G3 idem

   O encaixe leva os 5 moldes de M e de G e só 1 molde de cada tamanho maior. O
   CAD conta 2 modelos completos e a tabela sai "M 1 · G 1" — correto do ponto de
   vista dele, e não é a grade: a grade é M ao G3. Filtrando o seletor pelos
   tamanhos lidos, a única candidata era "M-G (CONJUGADO) | CM.LISA", que nada
   tem a ver, e a certa não aparecia em lugar nenhum.

   Regra que este teste protege: adivinhar a grade é papel do programa; IMPEDIR a
   correção não é. O seletor traz sempre o cadastro inteiro.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}

// Uma arrow de varias linhas: `const X = (...) => { ... };`
const arrow = comeco => {
  const i = src.indexOf(comeco);
  if (i < 0) { console.error('nao achei ' + comeco + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n};', i) + 3);
};

// Um bloco `const X = [ ... ];` inteiro.
const bloco = comeco => {
  const i = src.indexOf(comeco);
  if (i < 0) { console.error('nao achei ' + comeco + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n];', i) + 3);
};

// Uma linha inteira do app.js, para os `const X = ...` de uma linha so.
const linha = comeco => {
  const i = src.indexOf(comeco);
  if (i < 0) { console.error('nao achei ' + comeco + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n', i) + 1);
};

const api = new Function('STATE', `
  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  ${recorte('function _riscoGradesQueCasam', 'a busca por tamanhos')}
  ${recorte('function _riscoGradesProporcionais', 'a busca proporcional')}
  ${recorte('function _riscoCandidatas', 'a montagem das candidatas')}
  ${recorte('function _riscoFator', 'o fator do risco')}
  ${recorte('function _riscoMultiplo', 'o multiplo com o sentido')}
  ${recorte('function _normNome(', '_normNome')}
  ${recorte('function _riscoFormasDoNome', 'as formas do nome')}
  ${linha('const _chaveTam =')}
  ${recorte('function _riscoNomeTamanhos', 'o nome pelos tamanhos')}
  ${recorte('function _riscoNomeDivergente', 'o nome que discorda das quantidades')}
  ${recorte('function _riscoTamanhosTexto', 'o texto dos tamanhos')}
  ${recorte('function _riscoCelulaGrade', 'a celula da grade')}
  ${recorte('function _riscoFaseEhAgregadora', 'a fase agregadora')}
  ${bloco('const _RISCO_PECAS = [')}
  ${recorte('function _riscoFaseDoNomeArquivo', 'a fase pelo nome do arquivo')}
  ${recorte('function _riscoAssinatura', 'a assinatura dos tamanhos')}
  ${recorte('function _riscoDivisorDoGrupo', 'o divisor do grupo')}
  ${arrow('const _riscoTamanhosDivididos =')}
  ${recorte('function _riscoGruposNovos', 'os grupos de grade nova')}
  let _riscoLeituras = [];
  return { _riscoGruposNovos, _riscoLeituras, _riscoGradesQueCasam, _riscoGradesProporcionais, _riscoCandidatas, _riscoFator,
           _riscoMultiplo, _riscoTamanhosTexto, _riscoCelulaGrade, _riscoNomeDivergente,
           _riscoNomeTamanhos };
`);

// O cadastro, reduzido ao que importa: a grade CERTA e a que casa por engano.
const T = (o) => Object.assign({ p: 0, m: 0, g: 0, gg: 0, g1: 0, g2: 0, g3: 0 }, o);
const STATE = { grades: [
  { id: 'g_conj', nome: 'M-G (CONJUGADO) | CM.LISA | 117cm', tamanhos: T({ m: 1, g: 1 }) },
  { id: 'g_certa', nome: 'M ao G3 | CM.REC | 117cm', tamanhos: T({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }) },
  { id: 'g_outra', nome: '2P-2GG | CM.REC | 117cm', tamanhos: T({ p: 2, gg: 2 }) },
  // O caso da ribana: a grade do corpo, com 1 de cada, e o risco que a corta 10x.
  { id: 'g_ribana', nome: 'M-GG-G3 | CM.LISA | 117cm', tamanhos: T({ m: 1, gg: 1, g3: 1 }) },
  { id: 'g_dobrada', nome: 'M-2G-GG | CM.LISA | 117cm', tamanhos: T({ m: 1, g: 2, gg: 1 }) },
  // Grades de verdade cujos números PARECEM múltiplos: elas casam exato e a
  // proporção nao pode passar na frente.
  { id: 'g_4m4g', nome: '4M-4G | BM.LISA | 177cm', tamanhos: T({ m: 4, g: 4 }) },
  { id: 'g_8g', nome: '8G | BM.TRI | 179cm', tamanhos: T({ g: 8 }) },
  // A GRADE DOBRADA: o nome diz a metade, os tamanhos dizem o dobro, e o
  // relatorio dela e o da metade (caso real da "M-2G-GG | CM.TRI").
  { id: 'g_dobro', nome: 'M-2G-GG | CM.TRI | 116.5cm', tamanhos: T({ m: 2, g: 4, gg: 2 }) },
  // A mesma coisa, mas sem nenhuma grade com a distribuicao do proprio risco.
  { id: 'g_pg1', nome: '2P-2G1 | CM.TRI | 116.5cm', tamanhos: T({ p: 2, g1: 2 }) }
] };
const { _riscoGradesQueCasam, _riscoGradesProporcionais, _riscoCandidatas, _riscoFator,
        _riscoMultiplo, _riscoTamanhosTexto, _riscoCelulaGrade, _riscoNomeDivergente,
        _riscoNomeTamanhos, _riscoGruposNovos, _riscoLeituras } = api(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

console.log('-- os tamanhos que o PDF declarou --');
ok('1. lista so os que tem quantidade', _riscoTamanhosTexto({ m: 1, g: 1 }) === 'M 1 · G 1',
   _riscoTamanhosTexto({ m: 1, g: 1 }));
ok('2. mantem a ordem P..G3',
   _riscoTamanhosTexto({ g3: 2, p: 1, gg: 4 }) === 'P 1 · GG 4 · G3 2',
   _riscoTamanhosTexto({ g3: 2, p: 1, gg: 4 }));
ok('3. tabela vazia diz "nenhum"', _riscoTamanhosTexto({}) === 'nenhum' && _riscoTamanhosTexto(null) === 'nenhum');

console.log('');
console.log('-- o caso real do CORPO 2: a tabela nao descreve a grade --');
const lidos = { m: 1, g: 1 };                       // o que o PDF do CORPO 2 declara
const cands = _riscoGradesQueCasam(lidos);
ok('4. pelos tamanhos, a unica candidata e a CONJUGADO (a errada)',
   cands.length === 1 && cands[0].id === 'g_conj', cands.map(g => g.nome).join(', '));

const L = { tamanhos: lidos, grades: cands, grade: null };
const html = _riscoCelulaGrade(L, 0);
ok('5. a candidata aparece no grupo dos tamanhos do PDF',
   html.includes('com os tamanhos do PDF') && html.includes('corrigir: M-G (CONJUGADO)'), html);
ok('6. A GRADE CERTA tambem e oferecida, no grupo das outras',
   html.includes('todas as outras grades') && html.includes('value="g_certa"'), html);
ok('7. nao repete a candidata no grupo das outras',
   (html.match(/value="g_conj"/g) || []).length === 1, html);
ok('8. criar grade nova continua no fim', html.includes('__nova__'), html);
ok('9. havendo candidata, nao acusa "nenhuma grade"',
   !html.includes('nenhuma grade com estes tamanhos'), html);

console.log('');
console.log('-- escolhida a mao, fora das candidatas --');
const L2 = { tamanhos: lidos, grades: cands, grade: STATE.grades[1] };   // M ao G3
const html2 = _riscoCelulaGrade(L2, 0);
ok('10. a escolhida fica marcada no seletor',
   /value="g_certa" selected/.test(html2), html2);
ok('11. e a linha avisa que foi a mao, dizendo o que o PDF trazia',
   html2.includes('escolhida à mão') && html2.includes('M 1 · G 1'), html2);
const L3 = { tamanhos: lidos, grades: cands, grade: STATE.grades[0] };   // a candidata
ok('12. escolher a propria candidata NAO gera o aviso',
   !_riscoCelulaGrade(L3, 0).includes('escolhida à mão'), _riscoCelulaGrade(L3, 0));

console.log('');
console.log('-- nenhuma candidata (o CORPO 2 anterior, sem modelo completo) --');
const L4 = { tamanhos: {}, grades: [], grade: null };
const html4 = _riscoCelulaGrade(L4, 0);
ok('13. acusa que nada casou', html4.includes('nenhuma grade com estes tamanhos'), html4);
ok('14. e mesmo assim oferece o cadastro inteiro',
   html4.includes('value="g_certa"') && html4.includes('value="g_conj"')
   && html4.includes('value="g_outra"'), html4);
ok('15. sem candidatas nao ha o grupo do PDF',
   !html4.includes('com os tamanhos do PDF'), html4);

console.log('');
console.log('-- a busca por tamanhos continua exata --');
ok('16. casamento e por distribuicao inteira, nao por subconjunto',
   _riscoGradesQueCasam({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 })
     .map(g => g.id).join() === 'g_certa',
   _riscoGradesQueCasam({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }).map(g => g.nome).join());
ok('17. quantidade diferente nao casa', _riscoGradesQueCasam({ m: 2, g: 2 }).length === 0);
ok('18. tabela vazia nao casa com nada', _riscoGradesQueCasam({}).length === 0);

console.log('');
console.log('-- o risco que corta a grade N vezes (a RIBANA) --');
// CM.LISA RIBANA M-GG-G3.pdf: "RIBANA 57CM - 10x M GG G3", 30 modelos completos.
const ribana = { m: 10, gg: 10, g3: 10 };
ok('19. exato nao acha nada (era o "nenhuma grade com estes tamanhos")',
   _riscoGradesQueCasam(ribana).length === 0);
const prop = _riscoGradesProporcionais(ribana);
ok('20. a proporcional acha a grade do corpo, com o fator',
   prop.length === 1 && prop[0].grade.id === 'g_ribana' && prop[0].fator === 10,
   JSON.stringify(prop.map(p => [p.grade.nome, p.fator])));
const Lr = { tamanhos: ribana };
_riscoCandidatas(Lr);
Lr.grade = Lr.grades[0];
ok('21. vira candidata da leitura', Lr.grades.length === 1 && Lr.grades[0].id === 'g_ribana');
ok('22. e o fator fica disponivel para as unidades da fase', _riscoFator(Lr) === 10);
const htmlR = _riscoCelulaGrade(Lr, 0);
ok('23. o seletor escreve o fator', htmlR.includes('10× M-GG-G3'), htmlR);
ok('24. e nao acusa mais "nenhuma grade"',
   !htmlR.includes('nenhuma grade com estes tamanhos'), htmlR);
ok('25. a linha explica o que o multiplo significa',
   htmlR.includes('corta <b>10×</b> a grade'), htmlR);
// 10M-20G-10GG e 10x a "M-2G-GG | CM.LISA" (1/2/1) e 5x a "M-2G-GG | CM.TRI"
// (2/4/2). As duas sao candidatas de verdade, com fatores diferentes — quem
// escolhe e quem cadastra, e o seletor mostra os dois numeros.
const dez = _riscoGradesProporcionais({ m: 10, g: 20, gg: 10 });
ok('26. a grade dobrada tambem, e com o fator de cada uma',
   dez.map(p => p.grade.id + ':' + p.fator + ':' + p.sentido).sort().join(' ')
     === 'g_dobrada:10:risco g_dobro:5:risco',
   JSON.stringify(dez.map(p => [p.grade.nome, p.fator, p.sentido])));

console.log('');
console.log('-- o sentido inverso: a GRADE DOBRADA --');
// Caso real: "M-2G-GG | CM.TRI" esta cadastrada com {m:2,g:4,gg:2} (o nome diz
// a metade, os tamanhos dizem o dobro) e o relatorio dela e o de M-2G-GG. A
// grade se corta DUAS VEZES sobre esse encaixe.
const metade = _riscoGradesProporcionais({ m: 1, g: 2, gg: 1 });
const doDobro = metade.find(p => p.grade.id === 'g_dobro');
ok('26b. o risco da metade acha a grade dobrada',
   !!doDobro && doDobro.fator === 2 && doDobro.sentido === 'grade',
   JSON.stringify(metade.map(p => [p.grade.nome, p.fator, p.sentido])));
// EXATO CONTINUA VINDO PRIMEIRO, e e por isso que a tela so oferece a dobrada
// quando nao ha grade com a distribuicao do proprio risco: a "M-2G-GG |
// CM.LISA" (1/2/1) casa exato com este risco e ganha a vez. Quem quiser a
// dobrada assim mesmo a acha em "todas as outras grades".
const Lx = { tamanhos: { m: 1, g: 2, gg: 1 } };
_riscoCandidatas(Lx);
ok('26c. havendo grade exata, a dobrada NAO passa na frente',
   Lx.grades.length === 1 && Lx.grades[0].id === 'g_dobrada',
   Lx.grades.map(g => g.nome).join());

// O caso de verdade: risco de P-G1 e a grade cadastrada com o DOBRO (2/2), sem
// nenhuma grade com a distribuicao do risco.
const Ld = { tamanhos: { p: 1, g1: 1 } };
_riscoCandidatas(Ld);
Ld.grade = Ld.grades.find(g => g.id === 'g_pg1');
ok('26c2. sem exata, a dobrada e oferecida',
   !!Ld.grade, Ld.grades.map(g => g.nome).join() || '(nenhuma)');
ok('26d. o seletor escreve "1/2 de", e nao "2x"',
   _riscoCelulaGrade(Ld, 0).includes('1/2 de 2P-2G1 | CM.TRI'), _riscoCelulaGrade(Ld, 0));
ok('26d2. e a linha diz que a grade se corta 2 vezes sobre o risco',
   _riscoCelulaGrade(Ld, 0).includes('<b>1/2</b> da grade'), _riscoCelulaGrade(Ld, 0));
// A TRAVA QUE IMPORTA AQUI: no sentido inverso o numero NAO vira unidades da
// fase. Uma camada nao rende 2 grades — ela rende metade de uma.
ok('26e. o fator do sentido inverso NAO vale como unidades da fase',
   _riscoFator(Ld) === 1, _riscoFator(Ld));
ok('26f. mas o multiplo continua sabido, com o sentido',
   JSON.stringify(_riscoMultiplo(Ld)) === '{"fator":2,"sentido":"grade"}',
   JSON.stringify(_riscoMultiplo(Ld)));

console.log('');
console.log('-- o NOME da grade que discorda das quantidades dela --');
// 4 das 133 grades reais estao assim. "M-2G-GG | CM.TRI" esta cadastrada com
// M=2 G=4 GG=2, e "G | CM.TRI" com G=1 G1=2 — o nome esconde um tamanho
// inteiro. Quem confere o risco le o NOME; quem calcula pano usa a QUANTIDADE.
ok('26g. acusa o nome que diz a metade do que a grade tem',
   _riscoNomeDivergente({ nome: 'M-2G-GG | CM.TRI | 116.5cm', tamanhos: T({ m: 2, g: 4, gg: 2 }) }) === '2M-4G-2GG',
   _riscoNomeDivergente({ nome: 'M-2G-GG | CM.TRI | 116.5cm', tamanhos: T({ m: 2, g: 4, gg: 2 }) }));
ok('26h. e o nome que esconde um tamanho',
   _riscoNomeDivergente({ nome: 'G | CM.TRI | 116.5cm', tamanhos: T({ g: 1, g1: 2 }) }) === 'G-2G1',
   _riscoNomeDivergente({ nome: 'G | CM.TRI | 116.5cm', tamanhos: T({ g: 1, g1: 2 }) }));
ok('26i. nome que bate nao acusa nada',
   _riscoNomeDivergente({ nome: '2M-4G-2GG | CM.TRI | 116.5cm', tamanhos: T({ m: 2, g: 4, gg: 2 }) }) === '');
// A FAIXA nao e divergencia: "P ao G3" e "P-M-G-GG-G1-G2-G3" sao a mesma
// distribuicao escrita de dois jeitos, e a casa usa as duas.
ok('26j. faixa e forma extensa sao a mesma coisa',
   _riscoNomeDivergente({ nome: 'P ao G3 | CM.LISA | 117cm',
     tamanhos: T({ p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }) }) === '',
   _riscoNomeDivergente({ nome: 'P ao G3 | CM.LISA | 117cm',
     tamanhos: T({ p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }) }));
ok('26k. e a tela mostra o aviso quando ha divergencia',
   _riscoCelulaGrade({ tamanhos: { m: 1, g: 2, gg: 1 }, grades: [], fatores: {},
     grade: STATE.grades.find(g => g.id === 'g_dobro') }, 0).includes('as quantidades cadastradas dizem'));

console.log('');
console.log('-- o nome da GRADE NOVA nao carrega o 10x da ribana --');
/* "Nao deve existir uma grade com nome 10x no inicio, pois o nome da grade tem
   correspondencia com a fase CORPO, nao tem relacao com a fase gola" (Junior,
   20/08/2026). O encaixe da ribana leva dez grades no mesmo pano, entao a tela
   propunha criar "10M-10GG-10G3" — nome que nao descreve peca nenhuma. O dez
   pertence as UNIDADES da fase.

   A divisao e NA BASE DE DEZ, e nao pelo maior divisor comum, e as duas travas
   abaixo (casos 41 e 42) sao o motivo. */
const grupoDe = (arquivo, tamanhos) => {
  _riscoLeituras.length = 0;
  _riscoLeituras.push({ arquivo, tamanhos, grades: [] });
  return _riscoGruposNovos()[0];
};
const nomeDe = G => _riscoNomeTamanhos(G.tamanhosGrade);

let G10 = grupoDe('CM.LISA RIBANA M-GG-G3.pdf', { m: 10, gg: 10, g3: 10 });
ok('35. a ribana 10x nasce como a grade do corpo',
  nomeDe(G10) === 'M-GG-G3' && G10.fator === 10, nomeDe(G10) + ' fator ' + G10.fator);
ok('36. e a faixa tambem: "10X P ao G3" vira "P ao G3"',
  nomeDe(grupoDe('RIBANA 10X P M G GG G1 G2 G3.pdf',
    { p: 10, m: 10, g: 10, gg: 10, g1: 10, g2: 10, g3: 10 })) === 'P ao G3');
ok('37. gola tambem e agregadora',
  nomeDe(grupoDe('CM.LISA - GOLA 10P-10M.pdf', { p: 10, m: 10 })) === 'P-M');
ok('38. o que ja era 2G continua 2G depois de dividir',
  nomeDe(grupoDe('CM.LISA - RIBANA M-2G-GG.pdf', { m: 10, g: 20, gg: 10 })) === 'M-2G-GG');
// O caso que obriga a base 10: este arquivo mora na pasta "2M-2G". Pelo maior
// divisor comum (20) o nome sairia "M-G", que e OUTRA grade.
const G20 = grupoDe('CM.LISA - RIBANA 20M-20G.pdf', { m: 20, g: 20 });
ok('39. 20M-20G vira 2M-2G (a pasta dele), e nao M-G',
  nomeDe(G20) === '2M-2G' && G20.fator === 10, nomeDe(G20) + ' fator ' + G20.fator);
ok('40. o CORPO nunca e dividido — "4M-4G" e grade de verdade',
  nomeDe(grupoDe('BM.LISA - CORPO 4M-4G.pdf', { m: 4, g: 4 })) === '4M-4G');
// FORRO e agregadora, mas 2/4/2 nao e multiplo de dez: aquele encaixe e da
// grade 2M-4G-2GG mesmo. Pelo mdc viraria "M-2G-GG", que nao e aquela grade.
ok('41. forro 2M-4G-2GG fica como esta (mdc 2 nao e base dez)',
  nomeDe(grupoDe('BM.TRI - FORRO 2M-4G-2GG.pdf', { m: 2, g: 4, gg: 2 })) === '2M-4G-2GG');
ok('42. e o corpo 2/2/2 do tricolor tambem',
  nomeDe(grupoDe('CM.TRI - CORPO 2 - M-G-G1.pdf', { m: 2, g: 2, g1: 2 })) === '2M-2G-2G1');

console.log('');
console.log('-- as travas --');
const L4m = { tamanhos: { m: 4, g: 4 } };
_riscoCandidatas(L4m);
ok('27. EXATO PRIMEIRO: 4M-4G e uma grade de verdade, nao 4x a M-G',
   L4m.grades.length === 1 && L4m.grades[0].id === 'g_4m4g' && _riscoFator(L4m, L4m.grades[0]) === 1,
   L4m.grades.map(g => g.nome).join());
const L8g = { tamanhos: { g: 8 } };
_riscoCandidatas(L8g);
ok('28. o mesmo com a 8G da BM.TRI', L8g.grades.length === 1 && L8g.grades[0].id === 'g_8g',
   L8g.grades.map(g => g.nome).join());
ok('29. fator tem de fechar em TODOS os tamanhos',
   _riscoGradesProporcionais({ m: 10, gg: 15, g3: 10 }).length === 0);
ok('30. fator tem de ser inteiro',
   _riscoGradesProporcionais({ m: 3, gg: 3, g3: 3 }).length === 1
   && _riscoGradesProporcionais({ m: 2, g: 5, gg: 2 }).length === 0);
ok('31. os tamanhos tem de ser os MESMOS (nem sobra nem falta)',
   _riscoGradesProporcionais({ m: 10, gg: 10 }).length === 0
   && _riscoGradesProporcionais({ m: 10, g: 10, gg: 10, g3: 10 }).length === 0);
ok('32. fator 1 nao e proporcao (isso e o casamento exato)',
   _riscoGradesProporcionais({ m: 1, gg: 1, g3: 1 }).length === 0);
ok('33. tabela vazia nao casa por proporcao', _riscoGradesProporcionais({}).length === 0);
const Lmao = { tamanhos: ribana };
_riscoCandidatas(Lmao);
Lmao.grade = STATE.grades[1];                       // escolhida a mao, fora das candidatas
ok('34. grade escolhida a mao nao herda fator nenhum', _riscoFator(Lmao) === 1);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
