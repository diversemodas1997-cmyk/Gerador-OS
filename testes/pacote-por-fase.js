/* O PACOTE SÓ ACABA QUANDO TODAS AS FASES DELE FORAM (20/10/2026, Junior:
   "analise porque não consigo alocar na OE a OS661, sendo que pela manhã não
   alocou partes frente parte 1 parte 2 parte 3 e à tarde não consigo alocar as
   partes que não foram").

   _expAlocadoOS recortada do app.js; _expPacotesCanonicos dublada (3 pacotes).
   Rode com:  node testes/pacote-por-fase.js */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const bloco = (ini, fim) => src.slice(src.indexOf(ini), src.indexOf(fim, src.indexOf(ini)) + fim.length);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const STATE = { desenhos: [], modelos: [], ordens: [], expedicaoCargas: [] };
const api = new Function('STATE', `
  const corNomeCurto = s => String(s || '');
  const _expCancelSet = () => new Set();
  const _expPacotesCanonicos = () => ({ itens: [{ tam: 'P', tom: 1 }, { tam: 'M', tom: 1 }, { tam: 'G', tom: 1 }], temReposicao: true, total: 4 });
  ${bloco('const _OE_ITEM =', "reposicao: 'Reposição' };")}
  ${['_normNome', 'desenhoDaOS', '_skuBaseDaOS', '_oeLinhaDaOS', '_oeGruposDaFase', '_expOSTemItem', '_expFasesDaOS',
     '_expFasesDaCarga', '_expChavePacote', '_expContarPacotes', '_expAlocadoOS'].map(pegaFuncao).join('\n')}
  return { _expAlocadoOS };
`)(STATE);

const o = { id: 'os1', skuOverride: 'BM.TRI-BEGE', componentes: [],
  fases: ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Forro de capuz', 'Barra/Punhos', 'Viés'].map((nome, i) => ({ ordem: i + 1, nome })) };
STATE.ordens.push(o);
const PK = [{ tam: 'P', tom: 1 }, { tam: 'M', tom: 1 }, { tam: 'G', tom: 1 }];
const carga = (id, x) => Object.assign({ id, osId: 'os1', janelaId: 'j', data: '2026-10-20', perna: 'ida', pacotes: PK.slice(), reposicao: false }, x);
const total = a => { let q = 0; a.contagem.forEach(x => { q += x.qtd; }); return q; };
const aloc = () => api._expAlocadoOS('os1', '', 'ida');

STATE.expedicaoCargas = [carga('m', { fases: [3.6, 4, 5, 6], fasesComItens: true, reposicao: true })];
ok('1. manhã levou Capuz, Forro, Barra e Viés com todos os pacotes: nenhum pacote está alocado ainda', total(aloc()) === 0, total(aloc()));
ok('1b. a reposição foi', aloc().reposicao === true);
ok('1c. as fases que já foram inteiras ficam marcadas como feitas', [...aloc().fasesFeitas].join(',') === '3.6,4,5,6', [...aloc().fasesFeitas]);

STATE.expedicaoCargas.push(carga('t', { janelaId: 'j2', fases: [1, 2, 3, 3.3], fasesComItens: true }));
ok('2. à tarde as partes do corpo e as mangas: os 3 pacotes estão alocados', total(aloc()) === 3, total(aloc()));

STATE.expedicaoCargas = [carga('m', { fases: [1, 2, 3, 3.3], fasesComItens: true, pacotes: PK.slice(0, 2) }),
                         carga('t', { janelaId: 'j2', fases: [3.6, 4, 5, 6], fasesComItens: true })];
ok('3. P e M foram com o corpo; G só com o resto: só P e M estão alocados', total(aloc()) === 2, total(aloc()));

STATE.expedicaoCargas = [carga('m', { fases: null }), carga('t', { janelaId: 'j2', fases: [5], fasesComItens: true, pacotes: [] })];
ok('4. carga "só fase" (barra depois, sem pacote) não muda o que já foi inteiro', total(aloc()) === 3, total(aloc()));

STATE.expedicaoCargas = [carga('m', { fases: [1, 2, 3], fasesComItens: true }), carga('t', { janelaId: 'j2', fases: [3.3, 3.6, 4, 5, 6], fasesComItens: true, pacotes: [] })];
ok('5. carga "só fase" depois das partes do corpo completa a OS', total(aloc()) === 3, total(aloc()));

STATE.expedicaoCargas = [carga('m', { fases: [1, 4] })];
ok('6. carga parcial antiga (sem fasesComItens) continua valendo o pacote inteiro', total(aloc()) === 3, total(aloc()));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
