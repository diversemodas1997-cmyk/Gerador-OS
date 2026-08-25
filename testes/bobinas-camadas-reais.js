/* Rode com:  node testes/bobinas-camadas-reais.js

   Bobinas: quem manda é o cadastro da grade; o peso da bobina é consequência.

   Duas coisas a casa mede e sabe: quantas BOBINAS uma grade gasta num enfesto
   cheio (o campo do cadastro) e a GRAMATURA de cada tecido e cor. O que ninguém
   mede é quanto pesa uma bobina — elas não vêm todas iguais, ficam entre 18 e
   24 kg. Por isso o peso não entra na conta: sai dela.

       bobinas     = cadastro × (camadas desta OS ÷ camadas do enfesto CHEIO)
       kg da fase  = comp × larg × camadas × gramatura ÷ 1000
       ---------------------------------------------------------------
       metros/bob. = comp × camadas ÷ bobinas     <- resultado
       peso/bobina = kg ÷ bobinas                 <- resultado

   A versão anterior fazia o contrário: fixava a bobina em 20 kg e tirava as
   bobinas de kg ÷ 20. Duas medidas boas ficavam reféns de um chute, e na OS 0461
   (12/08/2026) a folha mostrou 4 bobinas onde o cadastro dizia 8 — porque o Azul
   estava cadastrado com a gramatura do pano simples, não a do enfesto.

   O peso que sai da conta vale como CONFERÊNCIA: fora de 18-24 kg, o errado é a
   gramatura daquela cor, e agora isso aparece em vez de se esconder no número.

   O ENFESTO CHEIO é o da própria grade, fase por fase (`L.camadasCheias`, que
   `consumoEnfestoOS` monta com `camadasCheiasDaFase`). Até 18/08/2026 era a
   constante 80, o limite da malha algodão — e como moletom para nas 36 camadas,
   toda grade de moletom aparecia na folha com 45% do que estava cadastrado, com
   o enfesto CHEIO. Foi assim na OS 0485 (BM.TRI 177cm): 6 bobinas cadastradas no
   Corpo Parte 1 saíam 3, e as 12 do Corpo Parte 3 saíam 6.

   O teste recorta as funções do app.js de verdade. Copiar a fórmula para cá
   testaria a cópia, que é exatamente como um defeito sobrevive em dois lugares. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('/* ===== INICIO BOBINAS POR CAMADAS REAIS');
const fim = src.indexOf('/* ===== FIM BOBINAS POR CAMADAS REAIS');
if (ini < 0 || fim < 0) {
  console.error('nao achei o trecho de bobinas por camadas reais no app.js');
  process.exit(1);
}

// `bobinasEfetivasFase` pergunta a esta função se a fase foi declarada não
// enfestada. Ela mora fora do trecho acima, então é recortada pelo nome — de
// novo, a de verdade: um stub aqui testaria o stub.
const iniF = src.indexOf('function _faseNaoEnfestadaPorTom');
if (iniF < 0) { console.error('nao achei _faseNaoEnfestadaPorTom no app.js'); process.exit(1); }
const fimF = src.indexOf('\n}', iniF);
if (fimF < 0) { console.error('nao achei o fim de _faseNaoEnfestadaPorTom'); process.exit(1); }

// `new Function` em vez de `eval`: dentro de eval, um `const` NÃO vaza para o
// escopo de fora (só `function` vaza), e as constantes daqui são const.
const api = new Function(
  src.slice(ini, fim) + '\n' + src.slice(iniF, fimF + 2) + '\n'
  + 'return { PESO_BOBINA_MIN_KG, PESO_BOBINA_MAX_KG,'
  + ' bobinaInteira, bobinasEfetivasFase, pesoPorBobinaFase, camadasPorBobina,'
  + ' metrosPorBobinaFase, ehFaseRibana, _tituloBobinas, _bobinasCelula };'
)();
const { PESO_BOBINA_MIN_KG, PESO_BOBINA_MAX_KG,
        bobinaInteira, bobinasEfetivasFase, pesoPorBobinaFase, camadasPorBobina,
        metrosPorBobinaFase, ehFaseRibana, _tituloBobinas, _bobinasCelula } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + extra));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(a - b) < 0.01;

// Uma OS reduzida ao que importa aqui. `tons` é opcional: { ordem: { tom: valor } }.
const os = tons => ({ enfesto: {}, progresso: tons ? { enfestosTons: tons } : {} });
// Uma linha de consumo como `consumoEnfestoOS` devolve. `camadasCheias` é o
// enfesto CHEIO desta fase nesta grade — a referência contra a qual o cadastro
// de bobinas foi preenchido. Sem passar, vale 80: o caso da malha algodão, que
// é o das grades abaixo.
const linha = (comp, larg, peso, camadas, corReal, camadasCheias) => ({
  comp, larg, peso, camadas, corReal: corReal || 'Preto Malha Algodão',
  camadasCheias: camadasCheias == null ? 80 : camadasCheias,
  kg: (comp * larg * camadas * peso) / 1000
});

// O caso real que levantou a regra: grade P-M-G-GG | CM.LISA | 117cm, fase
// Corpo, cadastrada com 8 bobinas para um enfesto cheio de 80 camadas.
const COMP = 4.25, LARG = 1.17, GRAM = 452.5, CADASTRO = 8;
const cheio = linha(COMP, LARG, GRAM, 80);
const os461 = linha(COMP, LARG, GRAM, 71);
// A mesma OS com a gramatura errada que estava no Azul (pano simples, 182).
const os461Errada = linha(COMP, LARG, 182, 71, 'Azul Malha Algodão');

console.log('-- as constantes da casa --');
ok('1. uma bobina de verdade pesa entre 18 e 24 kg',
   PESO_BOBINA_MIN_KG === 18 && PESO_BOBINA_MAX_KG === 24,
   PESO_BOBINA_MIN_KG + '-' + PESO_BOBINA_MAX_KG);

console.log('');
console.log('-- as bobinas saem do CADASTRO, na proporcao das camadas --');
ok('3. enfesto cheio (80 camadas) -> o cadastro intacto',
   bobinasEfetivasFase(os(), CADASTRO, 1, cheio) === 8,
   bobinasEfetivasFase(os(), CADASTRO, 1, cheio));
ok('4. 40 camadas -> metade do cadastro',
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, 40)) === 4,
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, 40)));
ok('5. 160 camadas -> o dobro (a referencia e absoluta, nao um teto)',
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, 160)) === 16,
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, 160)));
ok('6. sem camadas nao ha o que proporcionar -> devolve o cadastro',
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, 0)) === 8);
ok('7. sem cadastro nao ha previsao: devolve o que veio (a folha mostra tracinho)',
   bobinasEfetivasFase(os(), null, 1, cheio) === null,
   bobinasEfetivasFase(os(), null, 1, cheio));

console.log('');
console.log('-- o PESO do enfesto nao manda nas bobinas --');
ok('8. dobrar a gramatura NAO muda as bobinas: o cadastro e que sabe',
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM * 2, 71))
   === bobinasEfetivasFase(os(), CADASTRO, 1, os461),
   bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM * 2, 71)));
ok('9. a OS 0461 (71 camadas) -> 7,1 arredondado para 8 bobinas',
   bobinasEfetivasFase(os(), CADASTRO, 1, os461) === 8,
   bobinasEfetivasFase(os(), CADASTRO, 1, os461));
ok('10. e da o MESMO numero com a gramatura errada (o erro aparece no peso, nao aqui)',
   bobinasEfetivasFase(os(), CADASTRO, 1, os461Errada) === 8,
   bobinasEfetivasFase(os(), CADASTRO, 1, os461Errada));

console.log('');
console.log('-- bobina nao se abre pela metade: sempre para CIMA --');
ok('11. 7,1 bobinas viram 8', bobinaInteira(7.1) === 8, bobinaInteira(7.1));
ok('12. um resto minimo ja conta como bobina', bobinaInteira(0.05) === 1, bobinaInteira(0.05));
ok('13. o exato NAO sobe', bobinaInteira(7) === 7, bobinaInteira(7));
ok('14. nunca sai fracionado',
   [10, 20, 35, 71, 79, 80, 81].every(
     c => Number.isInteger(bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, c)))),
   [10, 20, 35, 71, 79, 80, 81].map(
     c => bobinasEfetivasFase(os(), CADASTRO, 1, linha(COMP, LARG, GRAM, c))).join(' '));

console.log('');
console.log('-- fase declarada NAO ENFESTADA na folha: so ela zera --');
const osMista = os({ 2: { 1: '0' } });
ok('15. a fase com tom 0 nao gasta bobina, mesmo com cadastro',
   bobinasEfetivasFase(osMista, CADASTRO, 2, cheio) === 0,
   bobinasEfetivasFase(osMista, CADASTRO, 2, cheio));
ok('16. a fase VIZINHA, enfestada, segue com o consumo dela',
   bobinasEfetivasFase(osMista, CADASTRO, 1, cheio) === 8,
   bobinasEfetivasFase(osMista, CADASTRO, 1, cheio));
ok('17. tom positivo em outro campo desmente o zero (a fase aconteceu)',
   bobinasEfetivasFase(os({ 2: { 1: '0', 2: '30' } }), CADASTRO, 2, cheio) === 8);
ok('18. tom em branco nao e resposta: nao zera nada',
   bobinasEfetivasFase(os({ 2: { 1: '' } }), CADASTRO, 2, cheio) === 8);

console.log('');
console.log('-- metro e peso da bobina sao CONSEQUENCIA --');
// O divisor e o valor CRU (cadastro x camadas/cheio), nao o arredondado: e
// assim que o metro por bobina fica igual em qualquer OS desta grade. Aqui o
// cheio e 80, que e o da malha algodao desta grade.
const cru = cam => CADASTRO * (cam / 80);
ok('19. o enfesto cheio da 42,5 m por bobina (4,25 m x 80 camadas / 8)',
   perto(metrosPorBobinaFase(cheio, cru(80)), 42.5), metrosPorBobinaFase(cheio, cru(80)));
ok('20. e 22,5 kg por bobina — dentro da faixa da prateleira',
   perto(pesoPorBobinaFase(cheio, cru(80)), 22.5)
   && pesoPorBobinaFase(cheio, cru(80)) >= PESO_BOBINA_MIN_KG
   && pesoPorBobinaFase(cheio, cru(80)) <= PESO_BOBINA_MAX_KG,
   pesoPorBobinaFase(cheio, cru(80)));
ok('21. em QUALQUER numero de camadas da o mesmo metro e o mesmo peso por bobina',
   perto(metrosPorBobinaFase(os461, cru(71)), 42.5)
   && perto(pesoPorBobinaFase(os461, cru(71)), 22.5),
   metrosPorBobinaFase(os461, cru(71)) + ' m / ' + pesoPorBobinaFase(os461, cru(71)) + ' kg');
ok('22. com a gramatura ERRADA a bobina daria 9 kg — abaixo da faixa, e denuncia o cadastro',
   perto(pesoPorBobinaFase(os461Errada, cru(71)), 9.05)
   && pesoPorBobinaFase(os461Errada, cru(71)) < PESO_BOBINA_MIN_KG,
   pesoPorBobinaFase(os461Errada, cru(71)));
ok('23. pano mais pesado -> bobina mais pesada, mesmos metros',
   pesoPorBobinaFase(linha(COMP, LARG, GRAM * 2, 80), cru(80)) > pesoPorBobinaFase(cheio, cru(80))
   && perto(metrosPorBobinaFase(linha(COMP, LARG, GRAM * 2, 80), cru(80)),
            metrosPorBobinaFase(cheio, cru(80))),
   pesoPorBobinaFase(linha(COMP, LARG, GRAM * 2, 80), cru(80)));
ok('24. camadas por bobina = camadas / bobinas',
   perto(camadasPorBobina(80, 8), 10), camadasPorBobina(80, 8));
ok('25. sem bobinas nao ha divisao -> null',
   pesoPorBobinaFase(cheio, 0) === null && metrosPorBobinaFase(cheio, 0) === null
   && camadasPorBobina(80, 0) === null);

console.log('');
console.log('-- a celula da folha --');
const fmt = n => String(n);
ok('26. com cadastro, mostra as bobinas inteiras',
   _bobinasCelula(os(), os461, CADASTRO, 1, true, fmt) === '8',
   _bobinasCelula(os(), os461, CADASTRO, 1, true, fmt));
ok('27. sem previsao cadastrada, mostra tracinho mesmo tendo peso',
   _bobinasCelula(os(), cheio, null, 1, false, fmt) === '—',
   _bobinasCelula(os(), cheio, null, 1, false, fmt));
ok('28. fase nao enfestada mostra zero, nao tracinho',
   _bobinasCelula(osMista, cheio, CADASTRO, 2, true, fmt) === '0',
   _bobinasCelula(osMista, cheio, CADASTRO, 2, true, fmt));

console.log('');
console.log('-- a dica explica de onde saiu o numero --');
const t = _tituloBobinas(os(), os461, CADASTRO, 1);
ok('29. cita o cadastro, o enfesto cheio de referencia e as camadas desta OS',
   t.includes('8 bobinas') && t.includes('80 camadas') && t.includes('71'), t);
ok('30. mostra o metro e o peso da bobina como resultado',
   t.includes('42,5 m de pano') && t.includes('22,5 kg'), t);
ok('31. dentro da faixa, nao alarma', !t.includes('ATENCAO'), t);
const tErr = _tituloBobinas(os(), os461Errada, CADASTRO, 1);
ok('32. bobina leve demais manda conferir a gramatura da cor',
   tErr.includes('ATENCAO') && tErr.includes('Azul Malha Algodão') && tErr.includes('182'),
   tErr);

console.log('');
console.log('-- o alarme nao pode tocar a toa --');
ok('33. pano pesado de verdade nao alarma (so o leve demais alarma)',
   !_tituloBobinas(os(), linha(COMP, LARG, GRAM * 3, 80), CADASTRO, 1).includes('ATENCAO'),
   _tituloBobinas(os(), linha(COMP, LARG, GRAM * 3, 80), CADASTRO, 1));
ok('34. fase nao enfestada diz isso, e nao um numero',
   _tituloBobinas(osMista, cheio, CADASTRO, 2).toLowerCase().includes('nao enfestada'),
   _tituloBobinas(osMista, cheio, CADASTRO, 2));

console.log('');
console.log('-- RIBANA: conta propria, e o programa espera os dados --');
// A gola da propria OS 0461. O "1 bobina" cadastrado na grade descreve o corpo,
// nao a ribana — e por isso NAO pode virar previsao de ribana.
const golaBase = { comp: 0.8, larg: 0.65, camadas: 14,
                   corReal: 'Azul Ribana Malha Algodão', tecidoReal: 'Ribana Malha Algodão' };
// `pesoDesteTecido` diz se a gramatura achada e da PROPRIA ribana; a fase
// costuma ficar com a cor do tecido principal, e aquela gramatura nao serve.
const ribana = (gram, pesoBobina) => ({ ...golaBase, peso: gram, pesoBobina,
                                        pesoDesteTecido: gram > 0,
                                        kg: (0.8 * 0.65 * 14 * gram) / 1000 });
ok('35. reconhece a fase de ribana pelo nome do tecido',
   ehFaseRibana(golaBase) === true && ehFaseRibana(cheio) === false);
ok('36. sem gramatura e sem peso de bobina -> nao prevê nada (folha mostra tracinho)',
   bobinasEfetivasFase(os(), null, 2, ribana(0, 0)) === null,
   bobinasEfetivasFase(os(), null, 2, ribana(0, 0)));
ok('37. com gramatura mas SEM o peso da bobina -> continua esperando',
   bobinasEfetivasFase(os(), null, 2, ribana(220, 0)) === null,
   bobinasEfetivasFase(os(), null, 2, ribana(220, 0)));
ok('38. com o peso da bobina mas SEM gramatura -> continua esperando',
   bobinasEfetivasFase(os(), null, 2, ribana(0, 8)) === null,
   bobinasEfetivasFase(os(), null, 2, ribana(0, 8)));
// 18/08/2026: o cadastro da grade passou a valer para a ribana tambem. Quem
// escreve um numero naquela fase contou AQUELE enfesto — recusa-lo era o que
// fazia o Barra/Punhos da BM.TRI sair com tracinho tendo "1" cadastrado, e o
// tecido inteiro sumir da tela de Compra.
ok('39. o cadastro de bobinas da grade AGORA vale para a ribana',
   bobinasEfetivasFase(os(), 99, 2, ribana(0, 0)) === 99,
   bobinasEfetivasFase(os(), 99, 2, ribana(0, 0)));
ok('39b. e ganha do peso quando os dois existem (quem contou olhou o enfesto)',
   bobinasEfetivasFase(os(), 3, 2, ribana(220, 8)) === 3,
   bobinasEfetivasFase(os(), 3, 2, ribana(220, 8)));
// Zero e resposta, como nas camadas: nao desce para a conta de peso atras de um
// numero que a casa ja disse nao existir.
ok('39c. zero cadastrado na grade e resposta, e nao volta para a conta de peso',
   bobinasEfetivasFase(os(), 0, 2, ribana(220, 8)) === 0,
   bobinasEfetivasFase(os(), 0, 2, ribana(220, 8)));
// 0,8 x 0,65 x 14 camadas x 220 g/m2 = 1,601 kg; bobina de 8 kg -> 0,2 -> 1.
ok('40. com as duas pontas cadastradas, prevê pelo peso: 1,6 kg / 8 kg -> 1 bobina',
   bobinasEfetivasFase(os(), null, 2, ribana(220, 8)) === 1,
   bobinasEfetivasFase(os(), null, 2, ribana(220, 8)));
ok('41. e acompanha o tamanho do enfesto (10 camadas de sobra pedem outra bobina)',
   bobinasEfetivasFase(os(), null, 2, { ...ribana(220, 8), camadas: 100,
     kg: (0.8 * 0.65 * 100 * 220) / 1000 }) === 2,
   bobinasEfetivasFase(os(), null, 2, { ...ribana(220, 8), camadas: 100,
     kg: (0.8 * 0.65 * 100 * 220) / 1000 }));
ok('42. ribana declarada nao enfestada segue zerando',
   bobinasEfetivasFase(osMista, null, 2, ribana(220, 8)) === 0,
   bobinasEfetivasFase(osMista, null, 2, ribana(220, 8)));
ok('43. a folha mostra tracinho enquanto falta dado',
   _bobinasCelula(os(), ribana(220, 0), null, 2, true, fmt) === '—',
   _bobinasCelula(os(), ribana(220, 0), null, 2, true, fmt));
const tRib = _tituloBobinas(os(), ribana(220, 0), null, 2);
ok('44. e a dica diz exatamente o que falta cadastrar',
   tRib.includes('peso medio da bobina') && tRib.includes('Ribana Malha Algodão')
   && !tRib.includes('gramatura de Azul'), tRib);
ok('45. faltando os dois, cobra os dois',
   _tituloBobinas(os(), ribana(0, 0), null, 2).includes('gramatura')
   && _tituloBobinas(os(), ribana(0, 0), null, 2).includes('peso medio'),
   _tituloBobinas(os(), ribana(0, 0), null, 2));
ok('46. com tudo cadastrado, a dica mostra a conta da ribana',
   _tituloBobinas(os(), ribana(220, 8), null, 2).startsWith('Ribana:'),
   _tituloBobinas(os(), ribana(220, 8), null, 2));

// O caso real da OS 0461: a fase de ribana esta com a cor do tecido principal
// ("Azul Malha Algodao", 182 g/m2). Ha gramatura, mas ela e de outro pano.
const golaEmprestada = { ...golaBase, corReal: 'Azul Malha Algodão', peso: 182,
                         pesoDesteTecido: false, pesoBobina: 8,
                         kg: (0.8 * 0.65 * 14 * 182) / 1000 };
ok('47. gramatura EMPRESTADA do tecido principal nao serve, mesmo com peso de bobina',
   bobinasEfetivasFase(os(), null, 2, golaEmprestada) === null,
   bobinasEfetivasFase(os(), null, 2, golaEmprestada));
ok('48. e a dica explica que a cor da fase e a do tecido principal',
   _tituloBobinas(os(), golaEmprestada, null, 2).includes('tecido principal'),
   _tituloBobinas(os(), golaEmprestada, null, 2));

console.log('');
console.log('-- MOLETOM: o enfesto cheio e o da GRADE, nao o da malha --');

/* A OS 0485 (18/08/2026), grade 2X P ao G3 | BM.TRI | 177cm. Moletom nao passa
   de 36 camadas (LIMITE_CAMADAS.moletom), e o enfesto estava CHEIO nas 36. A
   folha mostrava 3 bobinas onde o cadastro dizia 6, porque a referencia era a
   constante 80 — o limite da MALHA. Todo enfesto de moletom entrava valendo
   36/80 = 45%, e nenhuma grade de moletom conseguia mostrar o proprio cadastro. */
const MOL_GRAM = 300;
// comp, larg, gramatura, camadas desta OS, cor, camadas do enfesto CHEIO
const corpo485 = (cam) => linha(4.70, 1.77, MOL_GRAM, cam, 'Verde Moletom', 36);
// O forro de capuz e "2x" E a grade e 2-por-tamanho (2X P ao G3): os dois x2 se
// cancelam, e o cheio do forro passa a ser o do corpo, 36 (confirmado por Junior
// em 25/08/2026). Com forro 4x voltaria a 18.
const forro485 = (cam) => linha(2.78, 1.165, 452, cam, 'Preto Malha Algodão', 36);

ok('49. moletom cheio (36 de 36) devolve o cadastro intacto: 6 bobinas, nao 3',
   bobinasEfetivasFase(os(), 6, 1, corpo485(36)) === 6,
   bobinasEfetivasFase(os(), 6, 1, corpo485(36)));
ok('50. e as 12 do Corpo Parte 3 tambem saem inteiras, nao 6',
   bobinasEfetivasFase(os(), 12, 3, corpo485(36)) === 12,
   bobinasEfetivasFase(os(), 12, 3, corpo485(36)));
ok('51. meio enfesto de moletom (18 de 36) segue gastando meia previsao',
   bobinasEfetivasFase(os(), 6, 1, corpo485(18)) === 3,
   bobinasEfetivasFase(os(), 6, 1, corpo485(18)));
ok('52. o forro 2x numa grade 2x, cheio (36 de 36) -> 6 bobinas',
   bobinasEfetivasFase(os(), 6, 4, forro485(36)) === 6,
   bobinasEfetivasFase(os(), 6, 4, forro485(36)));
ok('53. e meio forro (18 de 36) volta a metade -> 3 bobinas',
   bobinasEfetivasFase(os(), 6, 4, forro485(18)) === 3,
   bobinasEfetivasFase(os(), 6, 4, forro485(18)));
ok('54. a mesma grade em malha (cheio 80) nao mudou de comportamento',
   bobinasEfetivasFase(os(), CADASTRO, 1, cheio) === 8,
   bobinasEfetivasFase(os(), CADASTRO, 1, cheio));

// Sem enfesto cheio conhecido (grade apagada) nao ha proporcao a fazer: o
// cadastro e a medida que a casa tirou a olho, e sai inteiro.
const semRef = { ...corpo485(36), camadasCheias: 0 };
ok('55. sem saber o enfesto cheio, devolve o cadastro em vez de inventar',
   bobinasEfetivasFase(os(), 6, 1, semRef) === 6,
   bobinasEfetivasFase(os(), 6, 1, semRef));

const tCheio = _tituloBobinas(os(), corpo485(36), 6, 1);
ok('56. a dica do enfesto cheio diz que ele esta cheio, e nao encolhe o numero',
   tCheio.includes('36 camadas') && tCheio.includes('cheio')
   && !tCheio.includes('arredondado'), tCheio);
const tMeio = _tituloBobinas(os(), corpo485(18), 6, 1);
ok('57. e a do enfesto parcial mostra a conta: 6 x 18/36 = 3',
   tMeio.includes('36 camadas') && tMeio.includes('18') && tMeio.includes('3'), tMeio);

console.log('');
console.log('-- a previsao escrita NA OS manda no que a grade preve --');

/* O campo "Bobinas previstas" de cada enfesto, na janela da OS. Em branco a
   folha segue a grade; escrito, ela obedece o que esta ali. Serve para o enfesto
   que saiu diferente do de sempre — e para a ribana, que sem gramatura e sem
   peso de bobina cadastrados nao tem como ser prevista por conta nenhuma. */
const comOverride = (L, n) => ({ ...L, bobinasOS: n });

ok('58. o numero escrito na OS vale como esta, sem proporcao de camadas',
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(18), 5)) === 5,
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(18), 5)));
ok('59. e vale mesmo quando a grade nao preve nada nesta fase',
   bobinasEfetivasFase(os(), null, 1, comOverride(corpo485(36), 4)) === 4,
   bobinasEfetivasFase(os(), null, 1, comOverride(corpo485(36), 4)));
ok('60. zero escrito na OS e resposta: nao gasta bobina',
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), 0)) === 0,
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), 0)));
ok('61. fracao escrita a mao sobe para a bobina inteira',
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), 0.5)) === 1,
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), 0.5)));
ok('62. em branco (null) nao e override: volta a seguir a grade',
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), null)) === 6,
   bobinasEfetivasFase(os(), 6, 1, comOverride(corpo485(36), null)));
ok('63. a ribana sem cadastro nenhum finalmente tem previsao quando escrita a mao',
   bobinasEfetivasFase(os(), 1, 2, comOverride(ribana(0, 0), 2)) === 2,
   bobinasEfetivasFase(os(), 1, 2, comOverride(ribana(0, 0), 2)));
ok('64. mas a fase declarada nao enfestada continua zerando por cima de tudo',
   bobinasEfetivasFase(os({ 2: { 1: '0' } }), 6, 2, comOverride(corpo485(36), 9)) === 0,
   bobinasEfetivasFase(os({ 2: { 1: '0' } }), 6, 2, comOverride(corpo485(36), 9)));
ok('65. a celula da folha mostra o numero escrito mesmo sem previsao na grade',
   _bobinasCelula(os(), comOverride(corpo485(36), 4), null, 1, false, fmt) === '4',
   _bobinasCelula(os(), comOverride(corpo485(36), 4), null, 1, false, fmt));
const tOS = _tituloBobinas(os(), comOverride(corpo485(36), 4), 6, 1);
ok('66. e a dica diz que o numero foi escrito na OS, e como desfazer',
   tOS.includes('escrita nesta OS') && tOS.includes('Apague'), tOS);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
