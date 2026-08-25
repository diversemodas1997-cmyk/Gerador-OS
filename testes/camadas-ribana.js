/* Rode com:  node testes/camadas-ribana.js

   A conta das camadas de ribana. O que importa aqui é uma coisa só: quem
   determina o rendimento de uma camada é o TAMANHO MAIS NUMEROSO da grade, e
   não a média dela.

   Era média. Em grade uniforme média e máximo são o mesmo número, então
   funcionou por muito tempo sem ninguém notar. Em grade desigual — M-2G-GG,
   2M-4G-2GG, G-2GG-G3 — a média mentia e o programa calculava MENOS ribana do
   que a produção precisava. E não havia como contornar pelo cadastro: o erro
   era proporcional, então mexer no multiplicador mexia nos dois lados juntos.

   O teste recorta a função do app.js de verdade. Copiar a fórmula para cá
   testaria a cópia, e a cópia é exatamente o que deixou o defeito viver em três
   lugares ao mesmo tempo. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('/* ===== INICIO CAMADAS RIBANA');
const fim = src.indexOf('/* ===== FIM CAMADAS RIBANA');
if (ini < 0 || fim < 0) { console.error('nao achei o trecho de camadas de ribana no app.js'); process.exit(1); }
const MULTIPLICADOR_PECAS = { malha: 2, moletom: 1, ribana: 2, outro: 1 };
eval(src.slice(ini, fim));

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra === undefined ? '' : extra)));
  if (!cond) falhas++;
};

// Blusa de moletom: multPrincipal = 1. Alvo 200 pecas do menor tamanho.
const camadas = (qtds, unidades, camadasPrincipal, escala = true) =>
  camadasDaFaseRibana({ camadasPrincipal, multPrincipal: 1, qtdsPorTamanho: qtds, unidades, escalaComGrade: escala });

console.log('-- o tamanho que manda --');
ok('1. grade uniforme: o que manda e o proprio numero', _tamanhoQueMandaNaGrade([0,2,2,2,0,0,0]) === 2);
ok('2. grade desigual: manda o MAIOR, nao a media', _tamanhoQueMandaNaGrade([0,1,2,1,0,0,0]) === 2);
ok('3. zeros nao contam como tamanho', _tamanhoQueMandaNaGrade([0,0,4,0,0,0,0]) === 4);
ok('4. grade vazia nao quebra a conta', _tamanhoQueMandaNaGrade([]) === 1);
ok('5. texto do formulario tambem serve', _tamanhoQueMandaNaGrade(['0','1','2','1']) === 2);

console.log('');
console.log('-- grades UNIFORMES: nada pode mudar (era media, virou maximo, e sao iguais) --');
// camadasPrincipal = 100 nos casos abaixo (grade com 2 por tamanho, alvo 200).
ok('6. 2M-2G-2GG com 2x -> 100 camadas, como sempre foi',
   camadas([0,2,2,2,0,0,0], 2, 100) === 100, camadas([0,2,2,2,0,0,0], 2, 100));
ok('7. 2X P ao G3 com 2x -> 100 camadas',
   camadas([2,2,2,2,2,2,2], 2, 100) === 100, camadas([2,2,2,2,2,2,2], 2, 100));
ok('8. P ao G3 (1 por tamanho) com 2x -> metade das camadas',
   camadas([1,1,1,1,1,1,1], 2, 200) === 100, camadas([1,1,1,1,1,1,1], 2, 200));
ok('9. 4M-4G com 2x -> o dobro de camadas (a camada cobre meia grade)',
   camadas([0,4,4,0,0,0,0], 2, 50) === 100, camadas([0,4,4,0,0,0,0], 2, 50));
ok('10. 8G com 2x -> quatro vezes as camadas do corpo',
   camadas([0,0,8,0,0,0,0], 2, 25) === 100, camadas([0,0,8,0,0,0,0], 2, 25));

console.log('');
console.log('-- grades DESIGUAIS: e aqui que o numero muda, porque estava errado --');
// M-2G-GG (1,2,1), alvo 200 -> camadasPrincipal = 200 (o menor tamanho pede 1).
// A camada que rende 2 pecas por tamanho cobre 1 unidade da grade (o G leva 2).
ok('11. M-2G-GG com 2x -> 200 camadas (a media dava 134, faltando 33%)',
   camadas([0,1,2,1,0,0,0], 2, 200) === 200, camadas([0,1,2,1,0,0,0], 2, 200));
// O caso do Junior: risco de 2M-4G-2GG cortado sobre M-2G-GG. Uma camada rende
// 2M+4G+2GG = 2 unidades da grade. O tamanho que manda (G) recebe 4 -> "4x".
ok('12. M-2G-GG com o risco de 2M-4G-2GG (4x) -> 100 camadas',
   camadas([0,1,2,1,0,0,0], 4, 200) === 100, camadas([0,1,2,1,0,0,0], 4, 200));
ok('13. 2M-4G-2GG com 2x -> 200 camadas (a media dava 134)',
   camadas([0,2,4,2,0,0,0], 2, 100) === 200, camadas([0,2,4,2,0,0,0], 2, 100));
ok('14. G-2GG-G3 (1,2,2) com 2x -> manda o 2, nao a media 1,67',
   camadas([0,0,1,2,0,0,2], 2, 100) === 100, camadas([0,0,1,2,0,0,2], 2, 100));
ok('15. 2M-3G-2GG (2,3,2) com 2x -> manda o 3',
   camadas([0,2,3,2,0,0,0], 2, 100) === 150, camadas([0,2,3,2,0,0,0], 2, 100));

console.log('');
console.log('-- ribana que NAO escala com a grade (gola de malha, gola polo) --');
ok('16. sem escala, a grade nao entra na conta',
   camadas([0,1,2,1,0,0,0], 10, 200, false) === 20, camadas([0,1,2,1,0,0,0], 10, 200, false));
ok('17. sem escala, formato da grade e indiferente',
   camadas([2,2,2,2,2,2,2], 10, 200, false) === camadas([0,1,2,1,0,0,0], 10, 200, false));
ok('18. o tecido diz quem escala: "Ribana Moletom" sim',
   _ribanaEscalaComGrade({ nome: 'Ribana Moletom' }) === true);
ok('19. "Ribana Malha Algodao" nao escala',
   _ribanaEscalaComGrade({ nome: 'Ribana Malha Algodao' }) === false);
ok('20. tecido ausente nao quebra', _ribanaEscalaComGrade(null) === false);

console.log('');
console.log('-- bordas --');
ok('21. sem unidades cadastradas, usa o padrao 2',
   camadas([0,2,2,2,0,0,0], null, 100) === 100, camadas([0,2,2,2,0,0,0], null, 100));
ok('22. nunca devolve zero camadas', camadas([0,2,2,2,0,0,0], 20, 1) === 1);
ok('23. camadas do corpo em zero nao vira negativo', camadas([0,2,2,2,0,0,0], 2, 0) === 1);
ok('24. malha (multPrincipal 2) dobra a ribana',
   camadasDaFaseRibana({ camadasPrincipal: 50, multPrincipal: 2, qtdsPorTamanho: [0,2,2,2,0,0,0],
                         unidades: 2, escalaComGrade: true }) === 100);

console.log('');
console.log('-- FORRO de capuz: escala com a grade, igual a ribana moletom --');
// Confirmado por Junior em 25/08/2026: grade 2x + forro 2x = cheio do corpo (os
// dois x2 se cancelam); forro 4x = metade do corpo.
const UNIDADES_PADRAO_FORRO = 2;
const forro = (qtds, unidades, camadasPrincipal) =>
  camadasDaFaseForro({ camadasPrincipal, unidades, qtdsPorTamanho: qtds });
ok('25. grade 2x + forro 2x -> o cheio do corpo (36), nao a metade',
   forro([0,2,2,2,0,0,0], 2, 36) === 36, forro([0,2,2,2,0,0,0], 2, 36));
ok('26. 2X P ao G3 (2x) + forro 2x -> 36 tambem (a OS 0485)',
   forro([2,2,2,2,2,2,2], 2, 36) === 36, forro([2,2,2,2,2,2,2], 2, 36));
ok('27. grade 2x + forro 4x -> metade do corpo (18)',
   forro([0,2,2,2,0,0,0], 4, 36) === 18, forro([0,2,2,2,0,0,0], 4, 36));
ok('28. grade 1-por-tamanho + forro 2x -> a metade de sempre (18)',
   forro([1,1,1,1,1,1,1], 2, 36) === 18, forro([1,1,1,1,1,1,1], 2, 36));
ok('29. grade desigual manda o MAIOR pedido, nao a media',
   forro([0,1,2,1,0,0,0], 2, 36) === 36, forro([0,1,2,1,0,0,0], 2, 36));
ok('30. sem unidades cadastradas, usa o padrao 2',
   forro([0,2,2,2,0,0,0], null, 36) === 36, forro([0,2,2,2,0,0,0], null, 36));
ok('31. nunca devolve zero camadas', forro([0,2,2,2,0,0,0], 4, 1) === 1);

console.log('');
if (falhas === 0) console.log('>>> todos passaram');
else { console.log('>>> ' + falhas + ' FALHA(S)'); process.exitCode = 1; }
