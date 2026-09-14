/* Rode com:  node testes/material-reservado-fases.js

   QUAL COLUNA CADA FASE OCUPA no quadro "OSs · material reservado".

   O defeito que este teste guarda (14/09/2026): o viés é MALHA nas duas peças
   da casa, e calcularPapeisFases classifica pela categoria do TECIDO. Então o
   viés saía disfarçado de outra coisa, sem erro nenhum na tela:

     - camiseta  -> malha sem moletom na OS = CORPO. A coluna "Corpo 2" de uma
                    básica era o viés, e a tricolor mostrava quatro corpos.
     - moletom   -> malha COM moletom na OS = FORRO DE CAPUZ. Numa básica sem
                    capuz a coluna do forro mostrava viés puro; numa com capuz,
                    forro e viés somados na mesma célula.

   No banco de 11/09 eram 250 fases de viés, 38 delas dentro de OS de moletom,
   contra 21 forros de verdade ("Forro de capuz" e "Forro"). Ou seja: na maioria
   das vezes que a coluna do forro aparecia, ela não era o forro.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = i < 0 ? -1 : src.indexOf(ate, i + de.length);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------- 1. quem é viés, pelo NOME ---------- */
const ehVies = new Function([
  corta('function _normNome'), corta('function _normFaseNome'),
  corta('function _ehFaseVies'), 'return _ehFaseVies;'].join(';'))();

ok('1. "Viés" é viés', ehVies('Viés'), ehVies('Viés'));
ok('2. "viés" sem acento e em minúscula também', ehVies('vies'), ehVies('vies'));
ok('3. "Gola e Viés" é linha de viés (o cadastro não acrescenta uma segunda)',
   ehVies('Gola e Viés'), ehVies('Gola e Viés'));
ok('4. "Corpo" não é', !ehVies('Corpo'), ehVies('Corpo'));
ok('5. "Enviesado" não é — a comparação é por palavra inteira',
   !ehVies('Enviesado'), ehVies('Enviesado'));
ok('6. "Forro de capuz" não é', !ehVies('Forro de capuz'), ehVies('Forro de capuz'));

/* ---------- 2. a coluna que cada fase ocupa ---------- */
// materialPorFaseOS puxa a conta de enfesto inteira (grade, gramatura, regra de
// excedente). O que se testa aqui é a REPARTIÇÃO das fases em colunas, então ela
// entra dublada, devolvendo as fases já com as marcas que o app calcula.
const colunas = new Function('FASES', `
  function materialPorFaseOS() { return FASES; }
  ${corta('function _juntaMaterial')}
  ${cortaLinha('function corposDoMaterialOS')}
  ${cortaLinha('function ribanaDoMaterialOS')}
  ${cortaLinha('function forroDoMaterialOS')}
  return {
    corpos: corposDoMaterialOS().map(f => f.nome),
    forro: (forroDoMaterialOS() || {}).nome || null,
    ribana: (ribanaDoMaterialOS() || {}).nome || null
  };
`);

const fase = (nome, extra) => Object.assign({ nome, kg: 1, bobinas: 1, ribana: false, vies: false, forro: false }, extra || {});

// Camiseta básica: Corpo, Gola (ribana), Viés.
const camiseta = colunas([
  fase('Corpo'), fase('Gola', { ribana: true }), fase('Viés', { vies: true })
]);
ok('7. camiseta básica: UM corpo, e não dois', camiseta.corpos.length === 1, camiseta.corpos);
ok('8. ... e o viés não é esse corpo', !camiseta.corpos.includes('Viés'), camiseta.corpos);
ok('9. ... a gola continua na ribana', camiseta.ribana === 'Gola', camiseta.ribana);
ok('10. ... e a camiseta não tem forro', camiseta.forro === null, camiseta.forro);

// Camiseta tricolor: três corpos de verdade + viés.
const tricolor = colunas([
  fase('Corpo Parte 1'), fase('Corpo Parte 2'), fase('Corpo Parte 3'),
  fase('Gola', { ribana: true }), fase('Viés', { vies: true })
]);
ok('11. tricolor: TRÊS corpos, não quatro', tricolor.corpos.length === 3, tricolor.corpos);

// Moletom com capuz: o forro é o forro, e nada além dele.
const moletomCapuz = colunas([
  fase('Corpo'), fase('Forro de capuz', { forro: true }),
  fase('Barra/Punhos', { ribana: true }), fase('Viés', { vies: true })
]);
ok('12. moletom com capuz: a coluna do forro é só o forro',
   moletomCapuz.forro === 'Forro de capuz', moletomCapuz.forro);
ok('13. ... e o viés não entrou nela', !/Vi/.test(String(moletomCapuz.forro)), moletomCapuz.forro);
ok('14. ... nem virou corpo', moletomCapuz.corpos.length === 1, moletomCapuz.corpos);

// Moletom básico (sem capuz): não existe forro nenhum para mostrar.
const moletomBasico = colunas([
  fase('Corpo'), fase('Barra/Punhos', { ribana: true }), fase('Viés', { vies: true })
]);
ok('15. moletom sem capuz: nao ha coluna de forro para inventar',
   moletomBasico.forro === null, moletomBasico.forro);

// Duas ribanas (gola + barra/punhos) continuam somadas numa coluna só.
const duasRibanas = colunas([
  fase('Corpo'), fase('Gola', { ribana: true }), fase('Barra/Punhos', { ribana: true })
]);
ok('16. duas ribanas continuam numa coluna só, somadas',
   duasRibanas.ribana === 'Gola + Barra/Punhos', duasRibanas.ribana);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
