/* Rode com:  node testes/excedente-padrao-por-faixa.js

   FASE SEM EXCEDENTE PRÓPRIO CAI NA FAIXA, e não num número fixo.

   O caso real de 12/08/2026 (print "importação 2", risco 11 de 20):

     CM.REC - CORPO 2 - M-2G-GG.pdf
     risco 0,84 × 1,175 m  +  15 cm "(padrão)"  =  0,99 × 1,175

   0,84 m está na primeira faixa (até 1,50 m → 10 cm) e mesmo assim somava 15.
   A regra das faixas existia, mas só tinha sido usada na alteração em massa
   sobre fases JÁ cadastradas; quem não tinha excedente próprio continuava
   caindo no EXCEDENTE_ENFESTO_PADRAO_CM, um 15 fixo. O certo é 0,94.

   O que este teste protege:

   1. Sem excedente próprio, vale a faixa do comprimento.
   2. COM excedente próprio, manda o cadastro — inclusive ZERO, que é um valor
      de verdade ("esta fase não leva sobra") e não "não sei".
   3. A base da faixa é o comprimento do RISCO, não o `comp` da fase. O `comp`
      já tem um excedente dentro; usá-lo seria medir a sobra em cima da sobra.
   4. Os 15 cm sobrevivem só onde não há comprimento nenhum em que se apoiar.

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

const api = new Function(`
  ${/const EXCEDENTE_ENFESTO_PADRAO_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/.exec(src)[0]}
  ${recorte('function excedentePorComprimento', 'a regra das faixas')}
  ${recorte('function excedenteEnfestoM', 'o excedente da fase')}
  ${/const excedenteEnfestoCm = [^;]+;/.exec(src)[0]}
  ${recorte('function excedenteEnfestoOrigem', 'a origem do excedente')}
  ${/const _riscoCompCadastro = \(compPdf, fase\) =>[\s\S]*?;/.exec(src)[0]}
  return { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoCm, excedenteEnfestoOrigem, _riscoCompCadastro };
`)();
const { EXCEDENTE_ENFESTO_PADRAO_CM, excedenteEnfestoCm: excCm, excedenteEnfestoOrigem: origem,
        _riscoCompCadastro: comp } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const r2 = n => (n == null ? null : +n.toFixed(2));

console.log('-- o caso do print (risco 11 de 20) --');
{
  const fase = { nome: 'Corpo Parte 2', excedente: '', comp: '0.94' };
  ok('1. 0,84 do risco cai na faixa de 10 cm, nao nos 15', excCm(fase, 0.84) === 10, excCm(fase, 0.84));
  ok('1b. e a medida de cadastrar vira 0,94', r2(comp(0.84, fase)) === 0.94, r2(comp(0.84, fase)));
  ok('1c. que e exatamente o que ja esta no cadastro — nada a corrigir',
     r2(comp(0.84, fase)) === parseFloat(fase.comp));
  ok('1d. e NAO o 0,99 que aparecia', r2(comp(0.84, fase)) !== 0.99);
  ok('1e. a tela diz que veio da faixa, nao "padrao"', origem(fase, 0.84) === 'faixa', origem(fase, 0.84));
}

console.log('\n-- sem excedente proprio: cada faixa --');
{
  const vazia = { excedente: '' };
  [[0.36, 10], [1.50, 10], [1.51, 15], [6.50, 15], [8, 15], [8.01, 20], [12, 20]].forEach(([c, cm]) =>
    ok(`2. risco de ${c} m -> ${cm} cm`, excCm(vazia, c) === cm, excCm(vazia, c)));
}

console.log('\n-- COM excedente proprio, manda o cadastro --');
{
  ok('3. fase de 20 cm num risco curto continua 20', excCm({ excedente: 20 }, 0.5) === 20,
     excCm({ excedente: 20 }, 0.5));
  ok('3b. e a origem e a fase', origem({ excedente: 20 }, 0.5) === 'fase');
  ok('3c. ZERO cadastrado e zero, nao "sem valor"', excCm({ excedente: 0 }, 0.5) === 0,
     excCm({ excedente: 0 }, 0.5));
  ok('3d. zero tambem conta como vindo da fase', origem({ excedente: 0 }, 9) === 'fase');
  ok('3e. e a medida nao ganha nada', r2(comp(2.00, { excedente: 0 })) === 2.00);
  ok('3f. fase de 10 cm (a da alteracao em massa) sobrevive num risco de 9 m',
     excCm({ excedente: 10 }, 9) === 10, excCm({ excedente: 10 }, 9));
}

console.log('\n-- a base e o comprimento do RISCO, nao o comp da fase --');
{
  // A fase tem comp 1,60 (que ja inclui excedente). O risco e 1,45. A faixa tem
  // de sair do 1,45 — senao a sobra e medida em cima da sobra e cai na faixa
  // errada, bem na borda.
  const fase = { excedente: '', comp: '1.60' };
  ok('4. risco 1,45 com fase de comp 1,60 -> 10 cm (a faixa do risco)',
     excCm(fase, 1.45) === 10, excCm(fase, 1.45));
  ok('4b. e _riscoCompCadastro usa a mesma base', r2(comp(1.45, fase)) === 1.55, r2(comp(1.45, fase)));
  // Sem base explicita, ai sim vale o comp da fase — e o unico numero que existe.
  ok('4c. sem base, cai no comp da fase (1,60 -> 15 cm)', excCm(fase) === 15, excCm(fase));
}

console.log('\n-- os 15 cm so onde nao ha comprimento nenhum --');
{
  const nada = { excedente: '' };
  ok('5. fase sem comp e sem base -> o padrao da casa', excCm(nada) === EXCEDENTE_ENFESTO_PADRAO_CM,
     excCm(nada));
  ok('5b. e a origem diz "padrao"', origem(nada) === 'padrao', origem(nada));
  ok('5c. risco acima de 12 m (fora da tabela) -> padrao', excCm(nada, 30) === EXCEDENTE_ENFESTO_PADRAO_CM,
     excCm(nada, 30));
  ok('5d. e ali a origem tambem e "padrao"', origem(nada, 30) === 'padrao', origem(nada, 30));
}

console.log('\n-- fase nova da importacao (a que nasce sem excedente) --');
{
  // pastaSalvarPasso cria a fase com excedente: '' de proposito. Antes isso
  // significava "15 cm"; agora significa "a faixa manda".
  const nova = { excedente: '' };
  ok('6. fase nova de risco 0,36 -> 10 cm', excCm(nova, 0.36) === 10, excCm(nova, 0.36));
  ok('6b. fase nova de risco 4,55 -> 15 cm', excCm(nova, 4.5493) === 15, excCm(nova, 4.5493));
  ok('6c. 4,5493 + 15 = 4,70 (o BM.TRICOLOR que calibrou o padrao segue igual)',
     r2(comp(4.5493, nova)) === 4.70, r2(comp(4.5493, nova)));
  ok('6d. fase nova de risco 9,10 -> 20 cm', excCm(nova, 9.10) === 20, excCm(nova, 9.10));
}

console.log('\n-- a fase vies, que nasce em 1,17 --');
{
  ok('7. o vies cai sozinho na faixa de 10 cm', excCm({ excedente: '' }, 1.17) === 10,
     excCm({ excedente: '' }, 1.17));
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
