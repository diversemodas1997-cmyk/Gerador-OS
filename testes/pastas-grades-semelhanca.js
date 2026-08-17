/* Rode com:  node testes/pastas-grades-semelhanca.js

   A ORDEM DAS PASTAS DO CADASTRO DE GRADES.

   As pastas de mesma família têm nome de mesma família, e é assim que se procura
   uma grade: as camisetas na sequência, as bermudas juntas, as blusas juntas.

   A lista saía em dois blocos — primeiro as pastas FIXAS do programa (camiseta,
   blusa_moletom, outro), na ordem em que estão escritas no código, e depois as
   criadas pelo usuário, em ordem alfabética. Com os dados reais da fábrica isso
   espalhava as famílias: "Camiseta Básica" abria a lista, "Blusa Moletom" vinha
   em seguida, e "Camiseta Oversized" e "Camiseta Polo" iam para o fim, longe da
   camiseta com que se parecem. A distinção fixa/criada é do CÓDIGO; quem lê a
   tela vê seis pastas iguais, com nomes.

   Duas coisas continuam mandando mais que o nome: a ordem posta à mão com as
   setas ↑↓, e os baldes ("Outro", "Sem categoria"), que vão para o fim.

   OBS: houve um terceiro nível de pasta (faixa de tamanhos) em 17/08/2026, que
   foi removido no mesmo dia — medido com as 130 grades reais, dava 86 grupos e
   72% deles com uma grade só. Ver o comentário em app.js, acima de renderGrades. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

// `labelTp`/`labelVr` leem os apelidos que o usuário deu às pastas (STATE), então
// entram de verdade no recorte — é o rótulo VISÍVEL que manda na ordem, e um
// apelido muda a posição da pasta.
const motor = [
  cortaLinha('const LABELS_TP_PADRAO'),
  cortaLinha('const LABELS_VR_PADRAO'),
  corta('function _gfl'),
  corta('function labelTp'),
  corta('function labelVr'),
  cortaLinha('const _PASTAS_BALDE'),
  corta('function _ordenarPastas')
].join('\n');

function ordenar(chaves, ordemManual, tipo, apelidos) {
  const fn = new Function('CHAVES', 'MANUAL', 'TIPO', 'APELIDOS', `
    const STATE = { gradeFolderLabels: APELIDOS };
    ${motor}
    return _ordenarPastas(CHAVES, MANUAL, TIPO === 'vr' ? labelVr : labelTp);
  `);
  return fn(chaves, ordemManual || [], tipo || 'tp',
    apelidos || { tp: {}, vr: {}, tpOrder: [], vrOrder: [] });
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => {
  const a = JSON.stringify(got), b = JSON.stringify(esperado);
  ok(nome + ' → ' + b, a === b, got);
};

/* ---------- 1. os dados reais da fábrica (17/08/2026) ----------
   Seis pastas: duas fixas do programa (camiseta, blusa_moletom) e quatro criadas
   pelo usuário. E `camiseta` tem apelido: "Camiseta Básica". */

const REAIS = ['Bermuda Moletom', 'Bermuda Tactel', 'Camiseta Oversized', 'Camiseta Polo', 'blusa_moletom', 'camiseta'];
const APELIDOS = { tp: { camiseta: 'Camiseta Básica' }, vr: { bicolor: 'Recortada' }, tpOrder: [], vrOrder: [] };

eq('as famílias saem na sequência, pelo nome que aparece',
  ordenar(REAIS, [], 'tp', APELIDOS),
  ['Bermuda Moletom', 'Bermuda Tactel', 'blusa_moletom', 'camiseta', 'Camiseta Oversized', 'Camiseta Polo']);

// O que a tela mostrava antes: Camiseta Básica, Blusa Moletom, e as outras
// camisetas no fim. Este teste existe para essa ordem não voltar.
const antes = ['camiseta', 'blusa_moletom', 'Bermuda Moletom', 'Bermuda Tactel', 'Camiseta Oversized', 'Camiseta Polo'];
ok('a ordem antiga (fixas primeiro) não volta',
  JSON.stringify(ordenar(REAIS, [], 'tp', APELIDOS)) !== JSON.stringify(antes), antes);

// O APELIDO é que manda, não a chave técnica: renomear "camiseta" para "Zebra"
// tira a pasta do meio das camisetas.
eq('o apelido muda a posição da pasta',
  ordenar(REAIS, [], 'tp', { tp: { camiseta: 'Zebra' }, vr: {}, tpOrder: [], vrOrder: [] }),
  ['Bermuda Moletom', 'Bermuda Tactel', 'blusa_moletom', 'Camiseta Oversized', 'Camiseta Polo', 'camiseta']);

/* ---------- 2. os baldes vão para o fim ---------- */

eq('"Outro" e "Sem categoria" não entram no meio dos nomes',
  ordenar(['outro', '', 'camiseta', 'Bermuda Tactel'], [], 'tp'),
  ['Bermuda Tactel', 'camiseta', 'outro', '']);

eq('subpasta sem variação também vai para o fim',
  ordenar(['', 'tricolor', 'basica'], [], 'vr'),
  ['basica', 'tricolor', '']);

/* ---------- 3. a mão do usuário manda mais que o nome ---------- */

eq('a ordem posta com as setas ↑↓ é respeitada',
  ordenar(REAIS, ['Camiseta Polo', 'camiseta'], 'tp', APELIDOS),
  ['Camiseta Polo', 'camiseta', 'Bermuda Moletom', 'Bermuda Tactel', 'blusa_moletom', 'Camiseta Oversized']);

eq('quem não está na ordem manual segue pelo nome, depois dela',
  ordenar(['camiseta', 'Bermuda Tactel', 'Bermuda Moletom'], ['camiseta'], 'tp'),
  ['camiseta', 'Bermuda Moletom', 'Bermuda Tactel']);

// Ordem manual guardada de uma pasta que não existe mais (renomeada, ou a última
// grade dela saiu) não pode abrir buraco nem sumir com as outras.
eq('pasta que saiu do cadastro é ignorada na ordem manual',
  ordenar(['camiseta', 'Bermuda Tactel'], ['Pasta Que Nao Existe', 'Bermuda Tactel'], 'tp'),
  ['Bermuda Tactel', 'camiseta']);

/* ---------- 4. detalhes de nome ---------- */

// "Básica Azul" e "basica azul 2" ficam VIZINHAS (é o que importa): acento e
// caixa não contam, então elas comparam como "basica azul" e "basica azul 2", e a
// mais curta vem primeiro.
eq('acento e caixa não separam famílias',
  ordenar(['basica azul 2', 'Básica Azul', 'Bermuda'], [], 'tp'),
  ['Básica Azul', 'basica azul 2', 'Bermuda']);

eq('número no nome sai em ordem de número, não de texto',
  ordenar(['Linha 10', 'Linha 2', 'Linha 1'], [], 'tp'),
  ['Linha 1', 'Linha 2', 'Linha 10']);

eq('lista vazia não explode', ordenar([], [], 'tp'), []);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
