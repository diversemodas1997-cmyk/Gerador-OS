/* Rode com:  node testes/bobinas-camadas-reais.js

   Bobinas de tecido a partir do PESO do enfesto.

   Uma bobina pesa cerca de 20 kg. O peso de um enfesto já era calculado para a
   baixa de estoque — comp × larg × camadas × gramatura ÷ 1000 —, então quantas
   bobinas ele consome é esse peso dividido por 20. Lida ao contrário, a mesma
   conta diz quantas camadas uma bobina rende naquele enfesto.

   O que isso conserta: a coluna "Consumo" mostrava um número fixo, digitado no
   cadastro da grade, que não se mexia por mais que as camadas mudassem. Quando a
   produção estendia menos camadas do que o planejado — e é isso que os tons da
   folha registram —, o estoque continuava com tecido reservado para um enfesto
   que não aconteceu, e a compra seguinte saía maior do que precisava.

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
// Devolvendo os símbolos explicitamente, o teste continua rodando o código do
// app.js e não depende de qual palavra-chave cada trecho usa para declarar.
const api = new Function(
  src.slice(ini, fim) + '\n' + src.slice(iniF, fimF + 2) + '\n'
  + 'return { PESO_BOBINA_KG, CAMADAS_REF_BOBINAS_CADASTRO, camadasPorBobina,'
  + ' bobinasEfetivasFase, _tituloBobinas, _bobinasCelula };'
)();
const { PESO_BOBINA_KG, CAMADAS_REF_BOBINAS_CADASTRO, camadasPorBobina,
        bobinasEfetivasFase, _tituloBobinas, _bobinasCelula } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + extra));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(a - b) < 1e-6;
// Bobina não se abre pela metade: o consumo previsto arredonda SEMPRE para cima
// e nunca sai fracionado. A expectativa passa pelo mesmo arredondamento — e o
// teste documenta que ele existe.
const bobDe = kg => Math.ceil(kg / 20 - 1e-9);

// Uma OS reduzida ao que importa aqui. `tons` é opcional: { ordem: { tom: valor } }.
const os = tons => ({ enfesto: {}, progresso: tons ? { enfestosTons: tons } : {} });
// Uma linha de consumo como `consumoEnfestoOS` devolve.
const linha = (comp, larg, peso, camadas) => ({
  comp, larg, peso, camadas, kg: (comp * larg * camadas * peso) / 1000
});
// Linha sem gramatura: não dá para pesar, cai na reserva do cadastro.
const semPeso = camadas => linha(6, 1.8, 0, camadas);

// Enfesto de referência: 6 m × 1,80 m de malha 180 g/m².
// Uma camada pesa 6 × 1,8 × 180 ÷ 1000 = 1,944 kg.
const COMP = 6, LARG = 1.8, GRAM = 180;
const L80 = linha(COMP, LARG, GRAM, 80);
const L40 = linha(COMP, LARG, GRAM, 40);

console.log('-- as constantes da casa --');
ok('1. uma bobina pesa 20 kg', PESO_BOBINA_KG === 20, PESO_BOBINA_KG);
ok('2. o cadastro da grade descreve um enfesto de 80 camadas',
   CAMADAS_REF_BOBINAS_CADASTRO === 80, CAMADAS_REF_BOBINAS_CADASTRO);

console.log('');
console.log('-- quantas camadas uma bobina rende --');
ok('3. 6 m x 1,80 m x 180 g/m2 -> ~10,29 camadas por bobina',
   perto(camadasPorBobina(COMP, LARG, GRAM), 20 / 1.944), camadasPorBobina(COMP, LARG, GRAM));
ok('4. pano mais pesado rende MENOS camadas',
   camadasPorBobina(COMP, LARG, 360) < camadasPorBobina(COMP, LARG, GRAM));
ok('5. enfesto mais curto rende MAIS camadas',
   camadasPorBobina(3, LARG, GRAM) > camadasPorBobina(COMP, LARG, GRAM));
ok('6. sem gramatura nao ha conta -> null', camadasPorBobina(COMP, LARG, 0) === null);
ok('7. sem comprimento nao ha conta -> null', camadasPorBobina(0, LARG, GRAM) === null);
ok('8. texto invalido nao quebra', camadasPorBobina('x', LARG, GRAM) === null);

console.log('');
console.log('-- as bobinas saem do peso do enfesto --');
const kgFixo = kg => ({ comp: COMP, larg: LARG, peso: GRAM, camadas: 10, kg });
ok('9. 40 kg -> 2 bobinas', bobinasEfetivasFase(os(), null, 1, kgFixo(40)) === 2,
   bobinasEfetivasFase(os(), null, 1, kgFixo(40)));
ok('10. 20 kg -> 1 bobina', bobinasEfetivasFase(os(), null, 1, kgFixo(20)) === 1);

console.log('');
console.log('-- bobina nao se abre pela metade: sempre para CIMA --');
ok('11. 10 kg (meia bobina) -> 1 bobina inteira',
   bobinasEfetivasFase(os(), null, 1, kgFixo(10)) === 1,
   bobinasEfetivasFase(os(), null, 1, kgFixo(10)));
ok('12. qualquer sobra abre mais uma bobina (38,88 kg = 1,944)',
   bobinasEfetivasFase(os(), null, 1, kgFixo(38.88)) === 2,
   bobinasEfetivasFase(os(), null, 1, kgFixo(38.88)));
ok('12a. um resto minimo ja conta como bobina (1,06 kg = 0,05)',
   bobinasEfetivasFase(os(), null, 1, kgFixo(1.06)) === 1,
   bobinasEfetivasFase(os(), null, 1, kgFixo(1.06)));
ok('12b. o exato NAO sobe: 40 kg sao 2 bobinas, nao 3',
   bobinasEfetivasFase(os(), null, 1, kgFixo(40)) === 2,
   bobinasEfetivasFase(os(), null, 1, kgFixo(40)));
ok('12c. nunca sai fracionado',
   [1.06, 10, 20, 38.88, 40, 155.52, 64.25].every(
     kg => Number.isInteger(bobinasEfetivasFase(os(), null, 1, kgFixo(kg)))),
   [1.06, 10, 20, 38.88, 40, 155.52, 64.25].map(
     kg => bobinasEfetivasFase(os(), null, 1, kgFixo(kg))).join(' '));
// A OS 0461 de 12/08/2026, que levantou a regra: 4,25 m x 1,17 m x 71 camadas
// de malha 182 g/m2 = 64,25 kg = 3,21 bobinas. Quem separa o material tira 4.
ok('12d. o caso da OS 0461: 64,25 kg -> 4 bobinas',
   bobinasEfetivasFase(os(), null, 1, kgFixo(64.25)) === 4,
   bobinasEfetivasFase(os(), null, 1, kgFixo(64.25)));

console.log('');
console.log('-- e por isso acompanham as camadas, sem proporcao nenhuma --');
ok('13. 80 camadas -> as bobinas do peso cheio',
   bobinasEfetivasFase(os(), null, 1, L80) === bobDe(L80.kg),
   bobinasEfetivasFase(os(), null, 1, L80));
// Com o arredondamento para cima a proporção deixa de ser exata: metade das
// camadas dá metade das bobinas ou uma a mais, nunca menos e nunca o dobro.
ok('14. metade das camadas -> cerca de metade das bobinas (nunca mais que isso)',
   bobinasEfetivasFase(os(), null, 1, L40) <= bobinasEfetivasFase(os(), null, 1, L80)
   && bobinasEfetivasFase(os(), null, 1, L40) * 2 >= bobinasEfetivasFase(os(), null, 1, L80),
   bobinasEfetivasFase(os(), null, 1, L40) + ' vs ' + bobinasEfetivasFase(os(), null, 1, L80));
ok('15. o peso cadastrado no cadastro da grade e IGNORADO quando ha como pesar',
   bobinasEfetivasFase(os(), 999, 1, L80) === bobDe(L80.kg),
   bobinasEfetivasFase(os(), 999, 1, L80));
// A invariante que amarra as duas contas: se uma bobina rende R camadas, então
// N camadas gastam N/R bobinas. Se isto quebrar, as duas pontas divergiram.
ok('16. camadas / (camadas por bobina) = bobinas do peso',
   perto(80 / camadasPorBobina(COMP, LARG, GRAM), L80.kg / 20),
   80 / camadasPorBobina(COMP, LARG, GRAM));

console.log('');
console.log('-- fase declarada NAO ENFESTADA na folha: so ela zera --');
const osMista = os({ 2: { 1: '0' } });
ok('17. a fase com tom 0 nao gasta bobina, mesmo tendo peso calculado',
   bobinasEfetivasFase(osMista, 14, 2, L80) === 0, bobinasEfetivasFase(osMista, 14, 2, L80));
ok('18. a fase VIZINHA, enfestada, segue com o consumo dela',
   bobinasEfetivasFase(osMista, 14, 1, L80) === bobDe(L80.kg),
   bobinasEfetivasFase(osMista, 14, 1, L80));
ok('19. tom positivo em outro campo desmente o zero (a fase aconteceu)',
   bobinasEfetivasFase(os({ 2: { 1: '0', 2: '30' } }), 14, 2, L80) === bobDe(L80.kg),
   bobinasEfetivasFase(os({ 2: { 1: '0', 2: '30' } }), 14, 2, L80));
ok('20. tom em branco nao e resposta: nao zera nada',
   bobinasEfetivasFase(os({ 2: { 1: '' } }), 14, 2, L80) === bobDe(L80.kg),
   bobinasEfetivasFase(os({ 2: { 1: '' } }), 14, 2, L80));

console.log('');
console.log('-- reserva: sem gramatura vale o cadastro, proporcional as 80 camadas --');
ok('21. enfesto cheio (80 camadas) -> o cadastro intacto',
   bobinasEfetivasFase(os(), 14, 1, semPeso(80)) === 14,
   bobinasEfetivasFase(os(), 14, 1, semPeso(80)));
ok('22. 40 camadas -> metade do cadastro',
   bobinasEfetivasFase(os(), 14, 1, semPeso(40)) === 7,
   bobinasEfetivasFase(os(), 14, 1, semPeso(40)));
ok('23. 20 camadas -> um quarto do cadastro (3,5), arredondado para 4',
   bobinasEfetivasFase(os(), 14, 1, semPeso(20)) === 4,
   bobinasEfetivasFase(os(), 14, 1, semPeso(20)));
// A referencia e ABSOLUTA: nao depende do que a OS planejou, so das camadas dela.
ok('24. 160 camadas -> o dobro (a referencia e absoluta, nao um teto)',
   bobinasEfetivasFase(os(), 14, 1, semPeso(160)) === 28,
   bobinasEfetivasFase(os(), 14, 1, semPeso(160)));
ok('25. sem camadas nao ha o que proporcionar -> devolve o cadastro',
   bobinasEfetivasFase(os(), 14, 1, semPeso(0)) === 14);
ok('26. sem peso e sem cadastro -> devolve o que veio (a folha mostra tracinho)',
   bobinasEfetivasFase(os(), null, 1, semPeso(40)) === null);

console.log('');
console.log('-- a celula da folha --');
const fmt = n => String(n);
ok('27. com peso, mostra o numero mesmo sem previsao cadastrada',
   _bobinasCelula(os(), L80, null, 1, false, fmt) === String(bobDe(L80.kg)),
   _bobinasCelula(os(), L80, null, 1, false, fmt));
ok('28. sem peso e sem previsao cadastrada, mostra tracinho',
   _bobinasCelula(os(), semPeso(40), null, 1, false, fmt) === '—',
   _bobinasCelula(os(), semPeso(40), null, 1, false, fmt));
ok('29. fase nao enfestada mostra zero, nao tracinho',
   _bobinasCelula(osMista, L80, 14, 2, true, fmt) === '0',
   _bobinasCelula(osMista, L80, 14, 2, true, fmt));

console.log('');
console.log('-- a dica explica de onde saiu o numero --');
const t = _tituloBobinas(os(), L80, 14, 1);
ok('30. com peso, mostra os kg, o divisor e as camadas por bobina',
   t.includes('20 kg') && t.includes('camadas'), t);
ok('31. fase nao enfestada diz isso, e nao um numero',
   _tituloBobinas(osMista, L80, 14, 2).toLowerCase().includes('nao enfestada'),
   _tituloBobinas(osMista, L80, 14, 2));
ok('32. sem peso, explica que veio do cadastro e cita as 80 camadas',
   _tituloBobinas(os(), semPeso(40), 14, 1).includes('80'),
   _tituloBobinas(os(), semPeso(40), 14, 1));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
