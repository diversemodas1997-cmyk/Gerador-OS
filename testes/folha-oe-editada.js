/* Rode com:  node testes/folha-oe-editada.js

   FOLHA DE OE EDITADA À MÃO (carga.folha).

   Cada quadro da folha de OE nasce calculado: nome da peça do desenho, cores das
   variantes, peças dos componentes, volumes dos pacotes da carga. Desde
   17/08/2026 o usuário pode reescrever esses campos POR OS ALOCADA — é o papel
   da doca, e nem sempre o cadastro tem a palavra final.

   O que este teste guarda é a fronteira: o texto reescrito manda no que a folha
   mostra e nos totais da perna, e NÃO manda em mais nada. `carga.pacotes` e
   `carga.volumes` continuam intactos — são eles que movem peça entre Estoque de
   corte e Expedição (ver lote-parcial-expedicao.js). Se um dia o override
   escorregar para dentro de carga.volumes, o saldo da fábrica muda sem que
   ninguém tenha pedido, e nenhuma tela vai reclamar.

   Campo em branco tem que continuar sendo PERGUNTA (usa o calculado) e o zero
   digitado, RESPOSTA — a mesma regra das camadas do enfesto.

   O teste recorta _expFolhaOS e resumoPernaExpedicao do app.js de verdade; o
   resto entra dublado (o que se mede aqui é o cruzamento override × calculado,
   não como as peças de um pacote são contadas). */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';

const motor = [
  corta('function _expFasesDaOS'),
  corta('function _expFasesDaCarga'),
  corta('function _expFasesTexto'),
  corta('function _expFolhaOS'),
  corta('function resumoPernaExpedicao')
].join('\n');

// Resumo da perna de ida de uma ocorrência, com o STATE dado.
function resumo(estado) {
  const fn = new Function('STATE', `
    const expCfg = () => ({ unidadeA: 'Fabrica', unidadeB: 'Loja', volMin: 0, volMax: 0 });
    const _expNum = (v, d) => { const n = Number(v); return Number.isFinite(n) ? n : d; };
    const _expCargasDa = (janelaId, data, perna) =>
      (STATE.expedicaoCargas || []).filter(c => c.janelaId === janelaId && c.data === data && c.perna === perna);
    // Dublês: 200 peças na OS; 4 vagas de tamanho, 50 pç cada.
    const _expPecasOS = () => 200;
    const _expPecasDaComposicao = (o, lista) => ({ pecas: lista.length * 50, total: 200, fracao: lista.length / 4 });
    const nomePecaOS = (o) => o.nomeDoDesenho || '';
    // Dublê: tirar o tecido do nome da cor tem teste próprio; aqui a cor entra
    // já curta ("Preto") e o que se mede é o rótulo da fase.
    const corNomeCurto = (n) => String(n == null ? '' : n).trim();
    ${motor}
    const oc = { janela: { id: 'j1' }, dataOrig: '2026-08-20' };
    return resumoPernaExpedicao(oc, 'ida');
  `);
  return fn(estado);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

const os = { id: 'os_1', os: '0501', nomeDoDesenho: 'Camiseta Recortada' };
// Meio lote (2 de 4 vagas) = 100 pç, 3 volumes com o de reposição.
const carga = (folha) => Object.assign({
  id: 'c1', osId: 'os_1', janelaId: 'j1', data: '2026-08-20', perna: 'ida',
  pacotes: [{ tam: 'P', tom: null }, { tam: 'M', tom: null }], reposicao: true, volumes: 3
}, folha ? { folha } : {});
const estado = (c) => ({ ordens: [os], expedicaoCargas: [c], expedicaoExcecoes: [] });

/* ---------- 1. sem nada reescrito: tudo calculado ---------- */

let r = resumo(estado(carga()));
eq('sem override, o nome vem do desenho', r.itens[0].modelo, 'Camiseta Recortada');
eq('sem override, as peças vêm dos pacotes', r.itens[0].pecas, 100);
eq('sem override, os volumes vêm da carga', r.itens[0].volumes, 3);
eq('sem override, a cor fica vazia (a folha calcula)', r.itens[0].cor, '');
ok('sem override, nada é marcado como editado', r.itens[0].folha.tem === false, r.itens[0].folha);

/* ---------- 2. cada campo reescrito manda na folha ---------- */

r = resumo(estado(carga({ modelo: 'Conjunto moletom', cor: 'Preto', pecas: 90, volumes: 12 })));
eq('o nome reescrito manda', r.itens[0].modelo, 'Conjunto moletom');
eq('a cor reescrita manda', r.itens[0].cor, 'Preto');
eq('as peças reescritas mandam', r.itens[0].pecas, 90);
eq('os volumes reescritos mandam', r.itens[0].volumes, 12);
eq('o total de volumes da perna segue o reescrito', r.volumes, 12);
eq('o total de peças da perna segue o reescrito', r.pecas, 90);
ok('é marcado como editado', r.itens[0].folha.tem === true, r.itens[0].folha);

// O calculado continua ao lado — é ele que a tela e o modal mostram como "de
// onde o número saiu". Sem isto, quem não editou não teria contra o que comparar.
eq('o nome calculado continua disponível', r.itens[0].modeloCalc, 'Camiseta Recortada');
eq('as peças calculadas continuam disponíveis', r.itens[0].pecasCalc, 100);
eq('os volumes calculados continuam disponíveis', r.itens[0].volumesCalc, 3);

/* ---------- 3. a fronteira: a carga não é tocada ---------- */

const c = carga({ pecas: 90, volumes: 12 });
resumo(estado(c));
eq('carga.volumes fica intacto (é ele que a regra da grade recalcula)', c.volumes, 3);
eq('carga.pacotes fica intacto (é ele que move peça no estoque)', c.pacotes.length, 2);

/* ---------- 4. vazio é pergunta, zero é resposta ---------- */

r = resumo(estado(carga({ modelo: '', cor: '', pecas: '', volumes: '' })));
eq('campo em branco cai no calculado (nome)', r.itens[0].modelo, 'Camiseta Recortada');
eq('campo em branco cai no calculado (peças)', r.itens[0].pecas, 100);
eq('campo em branco cai no calculado (volumes)', r.itens[0].volumes, 3);
ok('override todo em branco não conta como editado', r.itens[0].folha.tem === false, r.itens[0].folha);

r = resumo(estado(carga({ volumes: 0 })));
eq('zero volumes digitado é resposta, não campo vazio', r.itens[0].volumes, 0);
ok('zero digitado conta como editado', r.itens[0].folha.tem === true, r.itens[0].folha);

r = resumo(estado(carga({ pecas: 0 })));
eq('zero peças digitado é resposta', r.itens[0].pecas, 0);

/* ---------- 5. sujeira gravada não vira número torto ---------- */

r = resumo(estado(carga({ pecas: 'abc', volumes: -5 })));
eq('texto no lugar do número cai no calculado', r.itens[0].pecas, 100);
eq('número negativo cai no calculado', r.itens[0].volumes, 3);

r = resumo(estado(carga({ volumes: '9' })));
eq('número gravado como texto é lido como número', r.itens[0].volumes, 9);

r = resumo(estado(carga({ modelo: '   Espaço   ' })));
eq('o nome reescrito é aparado', r.itens[0].modelo, 'Espaço');

/* ---------- 6. carga sem composição (lote cheio antigo) ---------- */

const cheia = { id: 'c9', osId: 'os_1', janelaId: 'j1', data: '2026-08-20', perna: 'ida', volumes: 8 };
r = resumo(estado(cheia));
eq('carga antiga: peças = lote inteiro', r.itens[0].pecas, 200);
r = resumo(estado(Object.assign({}, cheia, { folha: { volumes: 10 } })));
eq('carga antiga também aceita o volume reescrito', r.itens[0].volumes, 10);

/* ---------- 7. fases do enfesto que a carga leva (carga.fases) ----------
   A expedição pode levar só parte da peça: corpo e manga vão costurar, a ribana
   espera. Ausente/vazio tem que continuar significando "a peça inteira" — é como
   toda carga anterior a este campo está gravada. */

const osComFases = Object.assign({}, os, { fases: [
  { ordem: 1, nome: 'Corpo', corNome: 'Preto', tecidoNome: 'Malha Algodão' },
  { ordem: 2, nome: 'Manga', corNome: 'Preto', tecidoNome: 'Malha Algodão' },
  { ordem: 3, nome: 'Ribana', corNome: 'Preto', tecidoNome: 'Ribana' }
] });
const estadoF = (c) => ({ ordens: [osComFases], expedicaoCargas: [c], expedicaoExcecoes: [] });
const fasesDe = (c) => resumo(estadoF(c)).itens[0].fases;

let f = fasesDe(carga());
ok('sem fases gravadas, a carga leva a peça inteira', f.todas === true && f.levam.length === 3, f);
eq('nenhuma fase fica para depois', f.ficam.length, 0);

f = fasesDe(Object.assign(carga(), { fases: [1, 2] }));
ok('duas fases marcadas: só elas embarcam', f.todas === false && f.levam.length === 2, f);
eq('a fase não marcada fica para depois', f.ficam.map(x => x.nome).join(','), 'Ribana');
eq('o rótulo da fase traz a cor', f.levam[0].rotulo, 'Corpo (Preto)');

f = fasesDe(Object.assign(carga(), { fases: [] }));
ok('lista vazia = peça inteira (não é "carga sem nada")', f.todas === true && f.levam.length === 3, f);

f = fasesDe(Object.assign(carga(), { fases: [1, 2, 3] }));
ok('todas marcadas = peça inteira', f.todas === true && f.ficam.length === 0, f);

// Grade alterada depois de alocar: a ordem gravada não existe mais na OS.
f = fasesDe(Object.assign(carga(), { fases: [1, 9] }));
eq('ordem que não existe mais na OS é ignorada', f.levam.map(x => x.nome).join(','), 'Corpo');
eq('e as outras duas fases ficam', f.ficam.length, 2);

// Ordem gravada como texto (JSON de fora, import): tem que ser lida como número,
// senão a carga inteira vira "peça inteira" calada.
f = fasesDe(Object.assign(carga(), { fases: ['2'] }));
eq('ordem gravada como texto é lida como número', f.levam.map(x => x.nome).join(','), 'Manga');

// OS sem fases de enfesto (grade antiga): não há o que escolher, e a folha não
// pode passar a comentar fase nenhuma.
f = resumo(estado(carga({ /* folha */ }))).itens[0].fases;
ok('OS sem fases: temFases false e nada a comentar', f.temFases === false && f.todas === true, f);

eq('texto das fases junta com ponto do meio',
  (() => { const fn = new Function('l', `
    const corNomeCurto = (n) => String(n == null ? '' : n).trim();
    ${motor}
    return _expFasesTexto(l);`);
    return fn([{ rotulo: 'Corpo (Preto)' }, { rotulo: 'Manga (Preto)' }]); })(),
  'Corpo (Preto) · Manga (Preto)');

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
