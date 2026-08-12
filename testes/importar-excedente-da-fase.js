/* Rode com:  node testes/importar-excedente-da-fase.js

   IMPORTAÇÃO DE RISCO: a medida de cadastrar sai do excedente DA FASE.

   O caso real, de 12/08/2026 (print "print importação pdf.jpg"):

     CM.REC - CORPO 2 - 2G 2GG.pdf  ·  risco 1,10 × 1,170 m
     fase do cadastro: "Corpo Parte 2", excedente 10 cm, cadastro 1,20 × 1,170

     esperado : 1,10 + 0,10 = 1,20  (igual ao cadastro — nada a corrigir)
     saía     : 1,10 + 0,15 = 1,25  (o PADRÃO da casa, e a tela dizia "(padrão)")

   A causa não era a conta: era a CÓPIA. _pastaResetFases monta a linha do
   rascunho a partir da fase resolvida do cadastro — nome, tecido, unidades,
   bobinas — e não copiava `excedente`. Ficava undefined, excedenteEnfestoM caía
   nos 15 cm, e a tela oferecia trocar um cadastro CERTO por um errado.

   Duas coisas travadas aqui, e a segunda importa tanto quanto a primeira:

   1. O rascunho carrega o excedente da fase, e a medida sai por ele.
   2. É SÓ LEITURA. A importação não grava `excedente` em lugar nenhum — nem na
      correção nem na grade nova. Quem cadastra o excedente é a grade; se algum
      dia um caminho de gravação escrever ali, este teste cai.

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

// STATE entra porque a excecao da ribana reconhece a fase tambem pelo TECIDO.
const api = new Function('STATE', `
  ${/const EXCEDENTE_ENFESTO_PADRAO_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/.exec(src)[0]}
  ${/const EXCEDENTE_RIBANA_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_VIES_CM = \d+;/.exec(src)[0]}
  ${recorte('function _normNome', 'a normalizacao de nome')}
  ${recorte('function _normFaseNome', 'a normalizacao de nome de fase')}
  ${recorte('function _ehFaseVies', 'o reconhecedor do vies')}
  ${recorte('function categoriaEfetivaTecido', 'a categoria do tecido')}
  ${recorte('function isTecidoRibana', 'o reconhecedor de tecido ribana')}
  ${recorte('function _ehFaseRibana', 'o reconhecedor da fase ribana')}
  ${recorte('function excedentePorComprimento', 'a regra das faixas')}
  ${recorte('function excedenteRegraDaFase', 'a regra inteira da fase')}
  ${recorte('function excedenteEnfestoM', 'o excedente da fase em metros')}
  ${/const excedenteEnfestoCm = [^;]+;/.exec(src)[0]}
  ${/const _riscoCompCadastro = \(compPdf, fase\) =>[\s\S]*?;/.exec(src)[0]}
  return { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoM, excedenteEnfestoCm, _riscoCompCadastro };
`)({ tecidos: [] });
const { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoCm, _riscoCompCadastro } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const r2 = n => (n == null ? null : +n.toFixed(2));

// A COPIA que _pastaResetFases faz da fase do cadastro para a linha do
// rascunho, recortada do proprio app.js: e a linha que estava faltando.
const copiaExcedente = /excedente: \(f && f\.excedente != null\) \? f\.excedente : ''/.exec(src);

console.log('-- a copia existe no _pastaResetFases --');
ok('1. o rascunho copia o excedente da fase resolvida', !!copiaExcedente,
   'a linha "excedente: (f && f.excedente != null) ? f.excedente : \'\'" sumiu de _pastaResetFases');
{
  // Ela tem que estar DENTRO de _pastaResetFases, nao em outro lugar qualquer.
  const bloco = recorte('function _pastaResetFases', '_pastaResetFases');
  ok('1b. e esta dentro de _pastaResetFases', /excedente:/.test(bloco), bloco.slice(0, 200));
}

console.log('\n-- o caso do print: fase de 10 cm --');
{
  const fase = { nome: 'Corpo Parte 2', excedente: 10, comp: '1.20', larg: '1.170' };
  ok('2. 1,10 do risco + 10 cm da fase = 1,20', r2(_riscoCompCadastro(1.10, fase)) === 1.20,
     r2(_riscoCompCadastro(1.10, fase)));
  ok('2b. bate com o que ja esta no cadastro (1,20) — nada a corrigir',
     r2(_riscoCompCadastro(1.10, fase)) === parseFloat(fase.comp));
  ok('2c. o rotulo da tela diz 10 cm, nao 15', excedenteEnfestoCm(fase) === 10, excedenteEnfestoCm(fase));
  ok('2d. e NAO da o valor errado que aparecia', r2(_riscoCompCadastro(1.10, fase)) !== 1.25);
}

// Sem excedente proprio E sem comprimento em que se apoiar, sobra o padrao da
// casa. Havendo comprimento, quem manda e a faixa — ver
// testes/excedente-padrao-por-faixa.js.
console.log('\n-- fase sem excedente proprio e sem comprimento: padrao da casa --');
{
  const vazio = { nome: 'Corpo', excedente: '' };
  const nulo = { nome: 'Corpo', excedente: null };
  const semCampo = { nome: 'Corpo' };
  ok('3. vazio -> padrao', excedenteEnfestoCm(vazio) === EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoCm(vazio));
  ok('3b. null -> padrao', excedenteEnfestoCm(nulo) === EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoCm(nulo));
  ok('3c. campo ausente -> padrao', excedenteEnfestoCm(semCampo) === EXCEDENTE_ENFESTO_PADRAO_CM,
     excedenteEnfestoCm(semCampo));
  ok('3d. 4,5493 + 15 = 4,70 (o caso do BM.TRICOLOR que calibrou o padrao)',
     r2(_riscoCompCadastro(4.5493, vazio)) === 4.70, r2(_riscoCompCadastro(4.5493, vazio)));
}

console.log('\n-- zero cadastrado e zero de verdade, nao "sem valor" --');
{
  const zero = { nome: 'Gola', excedente: 0 };
  ok('4. excedente 0 nao vira 15', excedenteEnfestoCm(zero) === 0, excedenteEnfestoCm(zero));
  ok('4b. e a medida nao ganha nada', r2(_riscoCompCadastro(2.00, zero)) === 2.00,
     r2(_riscoCompCadastro(2.00, zero)));
}

console.log('\n-- as faixas novas chegam inteiras na importacao --');
{
  // Depois da alteracao em massa as fases guardam 10/15/20. A importacao tem
  // que respeitar os tres, e nao so o que por acaso e igual ao padrao.
  [[10, 1.10, 1.20], [15, 6.50, 6.65], [20, 9.00, 9.20]].forEach(([cm, risco, esperado]) => {
    const f = { excedente: cm };
    ok(`5. fase de ${cm} cm: ${risco} -> ${esperado}`, r2(_riscoCompCadastro(risco, f)) === esperado,
       r2(_riscoCompCadastro(risco, f)));
  });
}

console.log('\n-- SO LEITURA: a importacao nao grava excedente --');
{
  // Na correcao de grade existente, o app escreve comp, larg, tecidoId,
  // unidades e bobinas na fase do cadastro. `excedente` NAO pode estar ali.
  const corpo = src.slice(src.indexOf('/* ---- CORRIGIR grade existente ---- */'));
  const ateOFim = corpo.slice(0, corpo.indexOf('\n}'));
  ok('6. o caminho de CORRECAO nao escreve f.excedente', !/\bf\.excedente\s*=/.test(ateOFim),
     (/\bf\.excedente\s*=.*/.exec(ateOFim) || [''])[0]);
  ok('6b. e continua escrevendo o que deve (comp e larg)',
     /\bf\.comp\s*=/.test(ateOFim) && /\bf\.larg\s*=/.test(ateOFim));
}
{
  // Fase nova nasce no padrao da casa: `excedente: ''`, nunca um numero.
  const nova = /excedente: '',\s*\/\/ fase nova nasce no padrão da casa/.test(src);
  ok('7. grade NOVA nasce com excedente vazio (padrao da casa)', nova);
}
{
  // Nenhum lugar do wizard de pasta atribui excedente a uma fase do cadastro.
  const wiz = src.slice(src.indexOf('function _pastaResetFases'));
  ok('8. o wizard inteiro nao tem atribuicao a .excedente',
     !/\.excedente\s*=\s/.test(wiz), (/.{0,60}\.excedente\s*=\s.{0,40}/.exec(wiz) || [''])[0]);
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
