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
// escopo de fora (só `function` vaza), e PESO_BOBINA_KG é const. Devolvendo os
// símbolos explicitamente, o teste continua rodando o código do app.js e não
// depende de qual palavra-chave cada trecho usa para declarar.
const api = new Function(
  src.slice(ini, fim) + '\n' + src.slice(iniF, fimF + 2) + '\n'
  + 'return { PESO_BOBINA_KG, camadasPorBobina, bobinasEfetivasFase,'
  + ' fatorCamadasReais, _tituloBobinas, _faseNaoEnfestadaPorTom };'
)();
const { PESO_BOBINA_KG, camadasPorBobina, bobinasEfetivasFase,
        fatorCamadasReais, _tituloBobinas, _faseNaoEnfestadaPorTom } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + extra));
  if (!cond) falhas++;
};
const perto = (a, b) => Math.abs(a - b) < 1e-6;
// O resultado sai arredondado em duas casas (bobina se mede em fração, mas não
// em milésimos). Comparar contra o valor exato falharia por 4 milésimos, então
// a expectativa passa pelo mesmo arredondamento — e assim o teste também
// documenta que o arredondamento existe.
const bobDe = kg => Math.round((kg / 20) * 100) / 100;

// Uma OS reduzida ao que importa aqui. `tons` é opcional: { ordem: { tom: valor } }.
const os = (planejadas, reais, tons) => ({
  enfesto: { camadasPlanejadas: planejadas, camadas: reais },
  progresso: tons ? { enfestosTons: tons } : {}
});
// Uma linha de consumo como `consumoEnfestoOS` devolve.
const linha = (comp, larg, peso, camadas) => ({
  comp, larg, peso, camadas, kg: (comp * larg * camadas * peso) / 1000
});

// Enfesto de referência: 6 m × 1,80 m de malha 180 g/m².
// Uma camada pesa 6 × 1,8 × 180 ÷ 1000 = 1,944 kg.
const COMP = 6, LARG = 1.8, GRAM = 180, KG_CAMADA = 1.944;

console.log('-- quantas camadas uma bobina de 20 kg rende --');
ok('1. a constante da casa e 20 kg', PESO_BOBINA_KG === 20, PESO_BOBINA_KG);
ok('2. 6 m x 1,80 m x 180 g/m2 -> ~10,29 camadas por bobina',
   perto(camadasPorBobina(COMP, LARG, GRAM), 20 / KG_CAMADA),
   camadasPorBobina(COMP, LARG, GRAM));
ok('3. pano mais pesado rende MENOS camadas',
   camadasPorBobina(COMP, LARG, 360) < camadasPorBobina(COMP, LARG, GRAM));
ok('4. enfesto mais curto rende MAIS camadas',
   camadasPorBobina(3, LARG, GRAM) > camadasPorBobina(COMP, LARG, GRAM));
ok('5. sem gramatura nao ha conta -> null', camadasPorBobina(COMP, LARG, 0) === null);
ok('6. sem comprimento nao ha conta -> null', camadasPorBobina(0, LARG, GRAM) === null);
ok('7. texto invalido nao quebra', camadasPorBobina('x', LARG, GRAM) === null);

console.log('');
console.log('-- as bobinas saem do peso do enfesto --');
ok('8. 40 kg -> 2 bobinas', bobinasEfetivasFase(os(80, 80), null, 1, 40) === 2,
   bobinasEfetivasFase(os(80, 80), null, 1, 40));
ok('9. 20 kg -> 1 bobina', bobinasEfetivasFase(os(80, 80), null, 1, 20) === 1);
ok('10. 10 kg -> meia bobina', bobinasEfetivasFase(os(80, 80), null, 1, 10) === 0.5);
ok('11. sobra quebrada mantem duas casas',
   bobinasEfetivasFase(os(80, 80), null, 1, 38.88) === 1.94,
   bobinasEfetivasFase(os(80, 80), null, 1, 38.88));

console.log('');
console.log('-- e por isso acompanham as camadas, sem proporcao nenhuma --');
// A mesma fase com 80 e com 40 camadas: o peso cai pela metade, as bobinas idem.
const L80 = linha(COMP, LARG, GRAM, 80);
const L40 = linha(COMP, LARG, GRAM, 40);
ok('12. 80 camadas -> as bobinas do peso cheio',
   bobinasEfetivasFase(os(80, 80), null, 1, L80.kg) === bobDe(L80.kg),
   bobinasEfetivasFase(os(80, 80), null, 1, L80.kg));
ok('13. metade das camadas -> metade das bobinas',
   perto(bobinasEfetivasFase(os(80, 40), null, 1, L40.kg) * 2,
         bobinasEfetivasFase(os(80, 80), null, 1, L80.kg)),
   bobinasEfetivasFase(os(80, 40), null, 1, L40.kg));
// A invariante que amarra as duas contas: se uma bobina rende R camadas, então
// N camadas gastam N/R bobinas. Se isto quebrar, as duas pontas divergiram.
ok('14. camadas ÷ (camadas por bobina) = bobinas do peso',
   perto(80 / camadasPorBobina(COMP, LARG, GRAM), L80.kg / 20),
   80 / camadasPorBobina(COMP, LARG, GRAM));

console.log('');
console.log('-- fase declarada NAO ENFESTADA na folha: so ela zera --');
const osMista = os(80, 80, { 2: { 1: '0' } });
ok('15. a fase com tom 0 nao gasta bobina, mesmo tendo peso calculado',
   bobinasEfetivasFase(osMista, 14, 2, L80.kg) === 0,
   bobinasEfetivasFase(osMista, 14, 2, L80.kg));
ok('16. a fase VIZINHA, enfestada, segue com o consumo dela',
   bobinasEfetivasFase(osMista, 14, 1, L80.kg) === bobDe(L80.kg),
   bobinasEfetivasFase(osMista, 14, 1, L80.kg));
ok('17. tom positivo em outro campo desmente o zero (a fase aconteceu)',
   bobinasEfetivasFase(os(80, 80, { 2: { 1: '0', 2: '30' } }), 14, 2, L80.kg) === bobDe(L80.kg),
   bobinasEfetivasFase(os(80, 80, { 2: { 1: '0', 2: '30' } }), 14, 2, L80.kg));
ok('18. tom em branco nao e resposta: nao zera nada',
   bobinasEfetivasFase(os(80, 80, { 2: { 1: '' } }), 14, 2, L80.kg) === bobDe(L80.kg),
   bobinasEfetivasFase(os(80, 80, { 2: { 1: '' } }), 14, 2, L80.kg));

console.log('');
console.log('-- sem gramatura nao ha peso: vale o cadastro, encolhido --');
ok('19. 14 bobinas cadastradas, enfesto pela metade -> 7',
   bobinasEfetivasFase(os(80, 40), 14, 1, 0) === 7, bobinasEfetivasFase(os(80, 40), 14, 1, 0));
ok('20. enfesto inteiro -> o cadastro intacto',
   bobinasEfetivasFase(os(80, 80), 14, 1, 0) === 14);
ok('21. estendeu MAIS que o planejado -> nao infla o cadastro',
   bobinasEfetivasFase(os(80, 120), 14, 1, 0) === 14);
ok('22. sem peso e sem cadastro -> devolve o que veio (a folha mostra tracinho)',
   bobinasEfetivasFase(os(80, 40), null, 1, 0) === null);

console.log('');
console.log('-- o fator do fallback --');
ok('23. estendeu tudo -> 1', fatorCamadasReais(os(80, 80)) === 1);
ok('24. estendeu metade -> 0,5', fatorCamadasReais(os(80, 40)) === 0.5);
ok('25. principal parada NAO zera o fator geral (as outras fases sobrevivem)',
   fatorCamadasReais(os(80, 0)) === 1, fatorCamadasReais(os(80, 0)));
ok('26. sem referencia (OS antiga) -> 1', fatorCamadasReais({ enfesto: { camadas: 40 } }) === 1);
ok('27. planejadas em 0 nao divide por zero', fatorCamadasReais(os(0, 40)) === 1);
ok('28. OS vazia nao quebra', fatorCamadasReais({}) === 1 && fatorCamadasReais(undefined) === 1);

console.log('');
console.log('-- a dica explica de onde saiu o numero --');
const t = _tituloBobinas(os(80, 80), L80, 14, 1);
ok('29. com peso, mostra os kg, o divisor e as camadas por bobina',
   t.includes('20 kg') && t.includes('camadas'), t);
ok('30. fase nao enfestada diz isso, e nao um numero',
   _tituloBobinas(osMista, L80, 14, 2).toLowerCase().includes('nao enfestada'),
   _tituloBobinas(osMista, L80, 14, 2));
ok('31. sem peso e sem reducao, e a dica padrao',
   _tituloBobinas(os(80, 80), linha(COMP, LARG, 0, 80), 14, 1)
     === 'Bobinas previstas (cadastro da grade)',
   _tituloBobinas(os(80, 80), linha(COMP, LARG, 0, 80), 14, 1));

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
