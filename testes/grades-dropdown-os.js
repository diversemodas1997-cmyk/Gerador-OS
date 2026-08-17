/* Rode com:  node testes/grades-dropdown-os.js

   QUAIS GRADES APARECEM NO DROPDOWN DA OS.

   O campo "Carregar grade pré-cadastrada" não mostra todas as grades: ele filtra
   pelo desenho e pelo modelo escolhidos (uma OS de camiseta não quer ver grade de
   moletom). O filtro é útil e é MUDO — grade que não casa simplesmente não
   existe na lista. Foi assim que as grades de uma família nova de produto
   (SKU COT, 17/08/2026) sumiram do cadastro para quem estava criando a OS: a
   pasta delas era nova, e a pasta nunca casa com o que o modelo sabe pedir.

   A regra que este teste guarda: cada filtro só corta quando OS DOIS LADOS são
   conhecidos. O que não se pode avaliar PASSA — senão o cadastro cresce e a lista
   encolhe, e ninguém descobre por quê.

   Recorta do app.js o filtro de verdade; os três leitores de tela
   (categoriaDesenhoOS, tipoPecaModeloOS, variacaoDesenhoOS) entram dublados,
   porque aqui o que se mede é o cruzamento, não a leitura do <select>. */
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
  cortaLinha('const _GRADE_TIPOS_CONHECIDOS'),
  cortaLinha('const _GRADE_VARIACOES_CONHECIDAS'),
  corta('function categoriaEfetivaTecido'),
  corta('function categoriaPrincipalGrade'),
  corta('function _ctxDropdownGradesOS'),
  corta('function _motivoGradeForaDoDropdownOS')
].join('\n');

// Motivo pelo qual cada grade fica fora do dropdown, dado o que o desenho e o
// modelo da tela pedem. '' = aparece.
function motivos(estado, tela, extraIds) {
  const fn = new Function('STATE', 'TELA', 'EXTRA', `
    const categoriaDesenhoOS = () => TELA.catDesenho || '';
    const tipoPecaModeloOS = () => TELA.tipoModelo || '';
    const variacaoDesenhoOS = () => TELA.variacao || '';
    ${motor}
    const ctx = _ctxDropdownGradesOS(EXTRA || []);
    const out = {};
    (STATE.grades || []).forEach(g => { out[g.nome] = _motivoGradeForaDoDropdownOS(g, ctx); });
    return out;
  `);
  return fn(estado, tela || {}, extraIds || []);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

const tecidos = [
  { id: 't_malha', nome: 'Malha Algodão', categoria: 'malha' },
  { id: 't_mole', nome: 'Moletom Flanelado', categoria: 'moletom' },
  { id: 't_rib', nome: 'Ribana Malha', categoria: 'ribana' },
  { id: 't_cot', nome: 'Cotton', categoria: '' }            // família nova, ainda sem categoria
];
const grade = (nome, tecidoId, tipoPeca, variacao) => ({
  id: 'g_' + nome, nome, tipoPeca, variacao,
  fases: [{ ordem: 1, nome: 'Corpo', tecidoId }, { ordem: 2, nome: 'Gola', tecidoId: 't_rib' }]
});
const estado = (grades) => ({ tecidos, grades });

/* ---------- 1. o filtro que faz sentido continua cortando ---------- */

let m = motivos(estado([
  grade('camiseta malha', 't_malha', 'camiseta', 'basica'),
  grade('moletom', 't_mole', 'blusa_moletom', 'basica')
]), { catDesenho: 'malha', tipoModelo: 'camiseta', variacao: 'basica' });
eq('grade que casa aparece', m['camiseta malha'], '');
eq('grade de moletom não aparece numa OS de malha', m['moletom'], 'categoria');

m = motivos(estado([grade('camiseta bicolor', 't_malha', 'camiseta', 'bicolor')]),
  { catDesenho: 'malha', tipoModelo: 'camiseta', variacao: 'basica' });
eq('variação diferente das cores do desenho não aparece', m['camiseta bicolor'], 'variacao');

m = motivos(estado([grade('moletom em pasta de camiseta', 't_mole', 'camiseta', 'basica')]),
  { catDesenho: 'moletom', tipoModelo: 'blusa_moletom', variacao: 'basica' });
eq('pasta conhecida e diferente da do modelo não aparece', m['moletom em pasta de camiseta'], 'tipo');

/* ---------- 2. o que NÃO se pode avaliar tem que passar ----------
   Era aqui que as grades sumiam. */

// Pasta NOVA (rótulo livre criado em "+ Nova pasta…"): o lado da OS só sabe
// produzir camiseta/blusa_moletom/outro, então comparar não diz nada.
m = motivos(estado([grade('P-M-G-GG-G1 | COT.PR | 152cm', 't_cot', 'COT', 'basica')]),
  { catDesenho: '', tipoModelo: 'camiseta', variacao: 'basica' });
eq('grade em pasta nova aparece', m['P-M-G-GG-G1 | COT.PR | 152cm'], '');

// Subpasta nova, mesma história.
m = motivos(estado([grade('COT bordada', 't_cot', 'COT', 'bordada')]),
  { catDesenho: '', tipoModelo: 'camiseta', variacao: 'basica' });
eq('grade em subpasta nova aparece', m['COT bordada'], '');

// Modelo categoria 'outro' vira tipoModelo 'outro' — que é o balde de tudo que
// não é camiseta nem moletom e não descreve a peça. Não pode exigir nada.
m = motivos(estado([grade('camiseta', 't_malha', 'camiseta', 'basica')]),
  { catDesenho: '', tipoModelo: 'outro', variacao: '' });
eq('modelo "outro" não exige pasta nenhuma', m['camiseta'], '');

// Tecido da 1ª fase sem categoria (família nova) e nenhuma fase categorizada:
// a categoria da grade é desconhecida, então o filtro de categoria não corta.
const soCot = { id: 'g_cot2', nome: 'COT sem categoria', tipoPeca: '', variacao: '',
  fases: [{ ordem: 1, nome: 'Corpo', tecidoId: 't_cot' }] };
m = motivos({ tecidos, grades: [soCot] }, { catDesenho: 'malha', tipoModelo: '', variacao: '' });
eq('grade de categoria desconhecida aparece', m['COT sem categoria'], '');

// Tecido excluído do cadastro: mesma situação — desconhecido, não errado.
const semTecido = { id: 'g_x', nome: 'grade com tecido excluido', fases: [{ ordem: 1, tecidoId: 't_sumiu' }] };
m = motivos({ tecidos, grades: [semTecido] }, { catDesenho: 'malha', tipoModelo: '', variacao: '' });
eq('grade cujo tecido saiu do cadastro aparece', m['grade com tecido excluido'], '');

// O caso que mais escondia grade: tecido novo na 1ª fase (sem categoria) e gola
// de RIBANA na 2ª. A categoria da grade não pode virar 'ribana' por causa da
// gola — nenhum desenho é de ribana, então a grade era cortada em toda OS.
m = motivos(estado([grade('COT com gola de ribana', 't_cot', 'COT', '')]),
  { catDesenho: 'malha', tipoModelo: '', variacao: '' });
eq('a gola de ribana não define a categoria da grade', m['COT com gola de ribana'], '');

// Mas grade que é MESMO de ribana (1ª fase ribana) continua sendo de ribana, e
// segue fora de uma OS de malha.
const soRibana = { id: 'g_rib', nome: 'ribana pura', tipoPeca: '', variacao: '',
  fases: [{ ordem: 1, nome: 'Ribana', tecidoId: 't_rib' }] };
m = motivos({ tecidos, grades: [soRibana] }, { catDesenho: 'malha', tipoModelo: '', variacao: '' });
eq('grade só de ribana continua sendo de ribana', m['ribana pura'], 'categoria');

// E o moletom com punho de ribana continua sendo moletom (1ª fase conhecida).
m = motivos(estado([grade('moletom com punho', 't_mole', 'blusa_moletom', 'basica')]),
  { catDesenho: 'moletom', tipoModelo: 'blusa_moletom', variacao: 'basica' });
eq('moletom com fase de ribana continua moletom', m['moletom com punho'], '');

/* ---------- 3. conjugadas e a grade já salva ---------- */

const conjA = grade('mãe', 't_malha', 'camiseta', 'basica');
conjA.conjugadaGradeId = 'g_filha';
const conjB = grade('filha', 't_malha', 'camiseta', 'basica');
conjB.id = 'g_filha';
m = motivos({ tecidos, grades: [conjA, conjB] }, { catDesenho: 'malha', tipoModelo: 'camiseta', variacao: 'basica' });
eq('a grade apontada como conjugada não aparece', m['filha'], 'conjugada');
eq('quem aponta continua aparecendo', m['mãe'], '');

// A grade salva da OS em edição aparece mesmo fora de todo filtro — senão abrir
// uma OS antiga trocaria a grade dela sem ninguém pedir.
m = motivos(estado([grade('moletom', 't_mole', 'blusa_moletom', 'basica')]),
  { catDesenho: 'malha', tipoModelo: 'camiseta', variacao: 'basica' }, ['g_moletom']);
eq('a grade já salva na OS aparece apesar do filtro', m['moletom'], '');

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
