/* Rode com:  node testes/bobinas-camadas-reais.js

   Bobinas: quem manda é o cadastro da grade; o peso da bobina é consequência.

   Duas coisas a casa mede e sabe: quantas BOBINAS uma grade gasta num enfesto
   cheio (o campo do cadastro) e a GRAMATURA de cada tecido e cor. O que ninguém
   mede é quanto pesa uma bobina — elas não vêm todas iguais, ficam entre 18 e
   24 kg. Por isso o peso não entra na conta: sai dela.

       bobinas     = cadastro × (camadas desta OS ÷ 80), sempre para cima
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
  + 'return { CAMADAS_REF_BOBINAS_CADASTRO, PESO_BOBINA_MIN_KG, PESO_BOBINA_MAX_KG,'
  + ' bobinaInteira, bobinasEfetivasFase, pesoPorBobinaFase, camadasPorBobina,'
  + ' metrosPorBobinaFase, _tituloBobinas, _bobinasCelula };'
)();
const { CAMADAS_REF_BOBINAS_CADASTRO, PESO_BOBINA_MIN_KG, PESO_BOBINA_MAX_KG,
        bobinaInteira, bobinasEfetivasFase, pesoPorBobinaFase, camadasPorBobina,
        metrosPorBobinaFase, _tituloBobinas, _bobinasCelula } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + extra));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(a - b) < 0.01;

// Uma OS reduzida ao que importa aqui. `tons` é opcional: { ordem: { tom: valor } }.
const os = tons => ({ enfesto: {}, progresso: tons ? { enfestosTons: tons } : {} });
// Uma linha de consumo como `consumoEnfestoOS` devolve.
const linha = (comp, larg, peso, camadas, corReal) => ({
  comp, larg, peso, camadas, corReal: corReal || 'Preto Malha Algodão',
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
ok('1. o cadastro da grade descreve um enfesto de 80 camadas',
   CAMADAS_REF_BOBINAS_CADASTRO === 80, CAMADAS_REF_BOBINAS_CADASTRO);
ok('2. uma bobina de verdade pesa entre 18 e 24 kg',
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
// O divisor e o valor CRU (cadastro x camadas/80), nao o arredondado: e assim
// que o metro por bobina fica igual em qualquer OS desta grade.
const cru = cam => CADASTRO * (cam / CAMADAS_REF_BOBINAS_CADASTRO);
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
ok('29. cita o cadastro, as 80 camadas de referencia e as desta OS',
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
// A gola da propria OS 0461: 1 bobina cadastrada, 14 camadas de ribana. Se a
// conta dividisse pelo numero arredondado, diria que a bobina pesa 1,3 kg.
const gola = { comp: 0.8, larg: 0.65, peso: 182, camadas: 14, corReal: 'Azul Ribana Malha Algodão',
               tecidoReal: 'Ribana Malha Algodão', kg: (0.8 * 0.65 * 14 * 182) / 1000 };
const tGola = _tituloBobinas(os(), gola, 1, 2);
ok('33. a gola nao inventa bobina de 1 kg (divide pelo cru, nao pelo arredondado)',
   !tGola.includes('1,33 kg') && !tGola.includes('1,32 kg'), tGola);
ok('34. e nao alarma: ribana vem em rolo pequeno mesmo',
   !tGola.includes('ATENCAO'), tGola);
ok('35. pano pesado de verdade tambem nao alarma (so o leve demais alarma)',
   !_tituloBobinas(os(), linha(COMP, LARG, GRAM * 3, 80), CADASTRO, 1).includes('ATENCAO'),
   _tituloBobinas(os(), linha(COMP, LARG, GRAM * 3, 80), CADASTRO, 1));
ok('36. fase nao enfestada diz isso, e nao um numero',
   _tituloBobinas(osMista, cheio, CADASTRO, 2).toLowerCase().includes('nao enfestada'),
   _tituloBobinas(osMista, cheio, CADASTRO, 2));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
