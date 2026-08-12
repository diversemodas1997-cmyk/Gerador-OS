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

const api = new Function('STATE', `
  ${/const EXCEDENTE_ENFESTO_PADRAO_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/.exec(src)[0]}
  ${/const EXCEDENTE_GOLA_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_VIES_CM = \d+;/.exec(src)[0]}
  ${/const EXCEDENTE_BARRA_CM = \d+;/.exec(src)[0]}
  ${/const _EXC_LIGACAO = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
  ${/const _PAL_VIES = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
  ${/const _PAL_GOLA = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
  ${/const _PAL_BARRA = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
  ${recorte('function _normNome', 'a normalizacao de nome')}
  ${recorte('function _normFaseNome', 'a normalizacao de nome de fase')}
  ${recorte('function _faseSoDe', 'o reconhecedor por nome inteiro')}
  ${recorte('function excedenteCfg', 'a regra cadastrada')}
  ${recorte('function excedentePorComprimento', 'a regra das faixas')}
  ${recorte('function excedenteRegraDaFase', 'a regra inteira da fase')}
  ${recorte('function excedenteEnfestoM', 'o excedente da fase')}
  ${/const excedenteEnfestoCm = [^;]+;/.exec(src)[0]}
  ${recorte('function excedenteEnfestoOrigem', 'a origem do excedente')}
  ${/const _riscoCompCadastro = \(compPdf, fase\) =>[\s\S]*?;/.exec(src)[0]}
  return { EXCEDENTE_ENFESTO_PADRAO_CM, EXCEDENTE_GOLA_CM, EXCEDENTE_VIES_CM, EXCEDENTE_BARRA_CM,
           excedenteEnfestoCm, excedenteEnfestoOrigem, _riscoCompCadastro, excedenteRegraDaFase };
`)({ meta: {} });
const { EXCEDENTE_ENFESTO_PADRAO_CM, EXCEDENTE_GOLA_CM, EXCEDENTE_VIES_CM, EXCEDENTE_BARRA_CM,
        excedenteEnfestoCm: excCm, excedenteEnfestoOrigem: origem,
        _riscoCompCadastro: comp, excedenteRegraDaFase: regra } = api;

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
  [[0.36, 10], [1.50, 10], [1.51, 15], [6.50, 15], [8.20, 15], [9, 15], [9.01, 20], [12, 20]].forEach(([c, cm]) =>
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

/* ===========================================================================
   AS EXCEÇÕES: GOLA 5 cm, VIÉS 0 cm — antes da faixa.

   Não foram inventadas: as 53 fases "Gola" do cadastro real já estavam com
   5 cm e os viéses com 0, digitados à mão, um a um.

   O RECONHECIMENTO É PELO NOME INTEIRO, e é aqui que mora o perigo. O cadastro
   tem três nomes que a regra do "contém a palavra" estragaria:

     "Corpo + Gola" ......  6 fases — é pano de CORPO
     "Separação de gola" .  1 fase  — é outra coisa
     "Barra/Punhos" ...... 46 fases, 41 delas com 15 cm cadastrado

   A primeira versão reconhecia a gola pelo TECIDO, e como Barra/Punhos é de
   ribana, 42 fases de 15 cm virariam 5 numa tacada. É a mesma armadilha que
   _opFaseForaDoPlanoPorPadrao já documenta, e a mesma defesa.
   =========================================================================== */

console.log('\n-- EXCECAO: vies leva 0, valha o comprimento que valer --');
{
  const vies = { nome: 'Viés', excedente: '' };
  ok('7. o vies de 1,17 leva 0, e nao os 10 cm da faixa', excCm(vies, 1.17) === EXCEDENTE_VIES_CM,
     excCm(vies, 1.17));
  ok('7b. e a medida de cadastrar nao ganha nada', r2(comp(1.17, vies)) === 1.17, r2(comp(1.17, vies)));
  ok('7c. um vies longo tambem leva 0 (nao depende da faixa)', excCm(vies, 9) === 0, excCm(vies, 9));
  ok('7d. a tela diz que veio da regra do vies', origem(vies, 1.17) === 'vies', origem(vies, 1.17));
  ok('7e. sem acento tambem', excCm({ nome: 'Vies', excedente: '' }, 9) === 0);
  ok('7f. vies SEM comprimento tambem sabe que leva 0', regra({ nome: 'Viés', excedente: '' }) === 0,
     regra({ nome: 'Viés', excedente: '' }));
}

console.log('\n-- EXCECAO: GOLA leva 5 --');
{
  const gola = { nome: 'Gola', excedente: '' };
  ok('8. gola leva 5 cm', excCm(gola, 2.5) === EXCEDENTE_GOLA_CM, excCm(gola, 2.5));
  ok('8b. e nao os 15 da faixa de 2,5 m', excCm(gola, 2.5) !== 15);
  ok('8c. gola longa tambem leva 5', excCm(gola, 9) === 5, excCm(gola, 9));
  ok('8d. a origem e a regra da gola', origem(gola, 2.5) === 'gola', origem(gola, 2.5));
  ok('8e. 2,50 + 0,05 = 2,55', r2(comp(2.50, gola)) === 2.55, r2(comp(2.50, gola)));
  ok('8f. gola SEM comprimento tambem sabe que leva 5', regra(gola) === 5, regra(gola));
  ok('8g. "Gola e Viés" e um pano de gola', excCm({ nome: 'Gola e Viés', excedente: '' }, 3) === 5,
     excCm({ nome: 'Gola e Viés', excedente: '' }, 3));
}

console.log('\n-- O NOME INTEIRO: os tres nomes que a regra frouxa estragaria --');
{
  // "Corpo + Gola" e pano de corpo. _normFaseNome troca o "+" por espaco.
  const cg = { nome: 'Corpo + Gola', excedente: '' };
  ok('9. "Corpo + Gola" NAO e gola — segue a faixa', excCm(cg, 2.71) === 15, excCm(cg, 2.71));
  ok('9b. e a origem e a faixa, nao a gola', origem(cg, 2.71) === 'faixa', origem(cg, 2.71));
  const sep = { nome: 'Separação de gola', excedente: '' };
  ok('9c. "Separação de gola" NAO e gola', excCm(sep, 3) === 15, excCm(sep, 3));
}

console.log('\n-- EXCECAO: BARRA/PUNHOS --');
{
  // Entrou como excecao em 12/08/2026. O padrao sao os 15 cm que as 41 fases
  // desse nome ja tinham cadastrado — assim a regra nasceu sem mudar nada.
  //
  // E de tecido de RIBANA, e foi por causa dela que o reconhecimento da gola
  // por TECIDO teve de sair: 42 fases de 15 cm virariam 5 numa tacada.
  const bp = { nome: 'Barra/Punhos', excedente: '' };
  ok('10. "Barra/Punhos" leva o valor da excecao', excCm(bp, 1.55) === EXCEDENTE_BARRA_CM,
     excCm(bp, 1.55));
  ok('10b. e NAO os 5 cm da gola', excCm(bp, 1.55) !== 5);
  ok('10c. a origem e a regra da barra', origem(bp, 1.55) === 'barra', origem(bp, 1.55));
  ok('10d. num risco CURTO tambem, sem cair na faixa dos 10',
     excCm(bp, 0.9) === EXCEDENTE_BARRA_CM, excCm(bp, 0.9));
  ok('10e. e num LONGO, sem cair na faixa dos 20',
     excCm(bp, 11) === EXCEDENTE_BARRA_CM, excCm(bp, 11));
  ok('10f. "Punhos/Barra" idem', excCm({ nome: 'Punhos/Barra', excedente: '' }, 1.55) === EXCEDENTE_BARRA_CM);
  ok('10g. "Barra" sozinha idem', excCm({ nome: 'Barra', excedente: '' }, 1.55) === EXCEDENTE_BARRA_CM);
  ok('10h. "Punhos" sozinho idem', excCm({ nome: 'Punhos', excedente: '' }, 1.55) === EXCEDENTE_BARRA_CM);
  ok('10i. SEM comprimento tambem sabe o valor', regra(bp) === EXCEDENTE_BARRA_CM, regra(bp));
  ok('10j. "Corpo + Barra" NAO e barra — segue a faixa',
     excCm({ nome: 'Corpo + Barra', excedente: '' }, 3) === 15,
     excCm({ nome: 'Corpo + Barra', excedente: '' }, 3));
  ok('10k. valor cadastrado na fase continua mandando',
     excCm({ nome: 'Barra/Punhos', excedente: 22 }, 1.55) === 22);
}

console.log('\n-- as excecoes NAO passam por cima do cadastro da fase --');
{
  ok('10. gola com 12 cm cadastrado continua 12', excCm({ nome: 'Gola', excedente: 12 }, 2.5) === 12,
     excCm({ nome: 'Gola', excedente: 12 }, 2.5));
  ok('10b. vies com 3 cm cadastrado continua 3', excCm({ nome: 'Viés', excedente: 3 }, 1.17) === 3,
     excCm({ nome: 'Viés', excedente: 3 }, 1.17));
  ok('10c. e as duas contam como vindas da fase',
     origem({ nome: 'Gola', excedente: 12 }, 2.5) === 'fase'
     && origem({ nome: 'Viés', excedente: 3 }, 1.17) === 'fase');
}

console.log('\n-- corpo nao e confundido com as excecoes --');
{
  ok('11. "Corpo Parte 2" segue a faixa', excCm({ nome: 'Corpo Parte 2', excedente: '' }, 0.84) === 10);
  ok('11b. "Enviesado" nao e vies', excCm({ nome: 'Enviesado', excedente: '' }, 3) === 15,
     excCm({ nome: 'Enviesado', excedente: '' }, 3));
  ok('11c. fase sem nome segue a faixa', excCm({ excedente: '' }, 3) === 15, excCm({ excedente: '' }, 3));
  ok('11d. "Golas" no plural conta', excCm({ nome: 'Golas', excedente: '' }, 3) === 5);
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
