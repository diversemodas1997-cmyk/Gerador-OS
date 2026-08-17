/* Rode com:  node testes/forma-do-pano.js

   TUBULAR OU ABERTO: QUANTAS UNIDADES UMA CAMADA RENDE.

   Tubular é o pano que vem em tubo, sem abertura lateral: enfestada, a camada já
   são duas espessuras e o corte sai em DUAS unidades iguais. Aberto rende UMA.

   Isso sempre esteve no programa, escondido dentro da CATEGORIA: a tabela
   MULTIPLICADOR_PECAS diz malha 2, moletom 1 — a mesma coisa dita de outro jeito
   ("malha vem em tubo, moletom vem aberto"). Em 17/08/2026 a forma do pano passou
   a ser um campo do tecido, e é ela que manda quando está preenchida.

   ESTE TESTE EXISTE POR UM MOTIVO SÓ, E É O MAIS IMPORTANTE DE TODOS AQUI:
   tecido SEM a forma declarada tem de continuar contando exatamente como contava.
   São 130 grades e 207 OS emitidas com a regra da categoria; se um cadastro em
   branco mudasse a conta, toda peça já cortada passaria a ter outro número na
   folha, no estoque de corte e nos pacotes da expedição — e ninguém pediu isso.
   Vazio é PERGUNTA, não resposta. */
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

const motor = [
  cortaLinha('const MULTIPLICADOR_PECAS'),
  corta('function _normNome'),
  corta('function _sufixoTecidoNorm'),
  corta('function categoriaEfetivaTecido'),
  corta('function unidadesPorCamadaTecido'),
  corta('function unidadesPorCamadaPrincipal'),
  corta('function tecidosDaOS'),
  corta('function multiplicadorPecaOS')
].join('\n');

function api(tecidos) {
  const fn = new Function('TECIDOS', `
    const STATE = { tecidos: TECIDOS };
    ${motor}
    return { unidadesPorCamadaTecido, unidadesPorCamadaPrincipal, multiplicadorPecaOS, categoriaEfetivaTecido };
  `);
  return fn(tecidos || []);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

// O cadastro como ele é hoje: nenhum tecido tem forma declarada.
const MALHA   = { id: 't1', nome: 'Malha Algodão', categoria: 'malha' };
const MOLETOM = { id: 't2', nome: 'Moletom Bulk', categoria: 'moletom' };
const RIBANA  = { id: 't3', nome: 'Ribana Bulk', categoria: '' };        // ribana é pelo NOME
const OUTRO   = { id: 't4', nome: 'Tactel', categoria: 'outro' };
const SEMCAT  = { id: 't5', nome: 'Cotton', categoria: '' };
const TODOS = [MALHA, MOLETOM, RIBANA, OUTRO, SEMCAT];
const A = api(TODOS);

/* ---------- 1. NADA MUDA para quem não declarou a forma ---------- */

eq('malha sem forma declarada: 2 por camada, como sempre', A.unidadesPorCamadaTecido(MALHA), 2);
eq('moletom sem forma declarada: 1 por camada, como sempre', A.unidadesPorCamadaTecido(MOLETOM), 1);
eq('ribana (reconhecida pelo nome): 2 por camada, como sempre', A.unidadesPorCamadaTecido(RIBANA), 2);
eq('"outro": 1 por camada, como sempre', A.unidadesPorCamadaTecido(OUTRO), 1);
eq('sem categoria e sem forma: 1 por camada, como sempre', A.unidadesPorCamadaTecido(SEMCAT), 1);
eq('tecido que não existe não vira 2 por acidente', A.unidadesPorCamadaTecido(null), 1);

/* ---------- 2. a forma declarada manda ---------- */

eq('tubular = 2 unidades iguais por camada',
  A.unidadesPorCamadaTecido({ nome: 'Cotton', categoria: 'outro', tubular: 'tubular' }), 2);
eq('não-tubular = 1 unidade por camada',
  A.unidadesPorCamadaTecido({ nome: 'Cotton', categoria: 'outro', tubular: 'aberto' }), 1);
// O caso que motivou o campo: a forma DISCORDA da categoria e vence.
eq('malha declarada ABERTA passa a contar 1 (a forma vence a categoria)',
  A.unidadesPorCamadaTecido({ nome: 'Malha Especial', categoria: 'malha', tubular: 'aberto' }), 1);
eq('moletom declarado TUBULAR passa a contar 2',
  A.unidadesPorCamadaTecido({ nome: 'Moletom Tubo', categoria: 'moletom', tubular: 'tubular' }), 2);
eq('ribana declarada ABERTA conta 1, apesar do nome',
  A.unidadesPorCamadaTecido({ nome: 'Ribana Aberta', categoria: '', tubular: 'aberto' }), 1);
// Valor estranho gravado (import, dado velho) não pode virar decisão.
eq('valor desconhecido no campo cai na categoria',
  A.unidadesPorCamadaTecido({ nome: 'Malha', categoria: 'malha', tubular: 'talvez' }), 2);

/* ---------- 3. a PEÇA PRINCIPAL entre os tecidos da OS ----------
   A escolha é a mesma de antes: o moletom, se houver; senão a malha. */

eq('OS de camiseta (malha + ribana): a malha manda → 2',
  A.unidadesPorCamadaPrincipal([MALHA, RIBANA]), 2);
eq('OS de moletom (moletom + malha de forro + ribana): o moletom manda → 1',
  A.unidadesPorCamadaPrincipal([MOLETOM, MALHA, RIBANA]), 1);
eq('a ordem em que os tecidos aparecem não muda a escolha',
  A.unidadesPorCamadaPrincipal([RIBANA, MALHA, MOLETOM]), 1);
eq('só ribana: contava 1 antes, continua 1',
  A.unidadesPorCamadaPrincipal([RIBANA]), 1);
eq('só "outro": contava 1 antes, continua 1',
  A.unidadesPorCamadaPrincipal([OUTRO]), 1);
eq('lista vazia: 1', A.unidadesPorCamadaPrincipal([]), 1);

// Com a forma declarada, ela vale para a peça principal escolhida.
eq('moletom declarado tubular manda na OS de moletom → 2',
  A.unidadesPorCamadaPrincipal([{ ...MOLETOM, tubular: 'tubular' }, MALHA]), 2);
eq('malha declarada aberta manda na OS de camiseta → 1',
  A.unidadesPorCamadaPrincipal([{ ...MALHA, tubular: 'aberto' }, RIBANA]), 1);
// Nem moletom nem malha, mas alguém DECLAROU a forma: é para isso que o campo existe.
eq('pano "outro" declarado tubular passa a contar 2',
  A.unidadesPorCamadaPrincipal([{ ...OUTRO, tubular: 'tubular' }]), 2);
eq('a forma de uma RIBANA não decide pela peça principal quando há malha',
  A.unidadesPorCamadaPrincipal([{ ...RIBANA, tubular: 'aberto' }, MALHA]), 2);

/* ---------- 4. a OS inteira (é o número que vai para a folha) ---------- */

const os = (fases, tecs) => ({ fases: fases || [], tecidos: tecs || [] });

eq('OS antiga de camiseta (fases da grade): 2, como antes',
  A.multiplicadorPecaOS(os([{ tecidoId: 't1' }, { tecidoId: 't3' }])), 2);
eq('OS antiga de moletom: 1, como antes',
  A.multiplicadorPecaOS(os([{ tecidoId: 't2' }, { tecidoId: 't1' }])), 1);
eq('OS sem fases, só com a lista de tecidos: também conta',
  A.multiplicadorPecaOS(os([], [{ tecidoId: 't1' }])), 2);
eq('OS sem tecido nenhum: 1', A.multiplicadorPecaOS(os()), 1);
eq('OS nula não quebra a folha', A.multiplicadorPecaOS(null), 1);

// A MESMA OS, depois de o cadastro do tecido declarar a forma: o número muda —
// é o efeito pedido, e é por isso que a dica do campo avisa.
const B = api([{ ...MALHA, tubular: 'aberto' }, RIBANA]);
eq('declarar a malha como aberta muda a OS de 2 para 1',
  B.multiplicadorPecaOS(os([{ tecidoId: 't1' }, { tecidoId: 't3' }])), 1);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
