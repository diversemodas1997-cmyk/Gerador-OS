/* ITENS DE CADA FASE NA FOLHA DE OE, COMO CHECKLIST (25/09/2026, Junior:
   "quando a OS for CM.LISA deve mostrar os itens frente, costa, mangas,
   ribana, viés, tecido de reposição. CM.REC igual a CM.LISA mais corpo parte 2
   ... corrija corpo parte 2, corpo parte 3 por frente parte 2, frente parte 3").

   A lista e recortada do app.js e montada com OS de mentira. */
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

const api = new Function('STATE', `
  const corNomeCurto = s => String(s || '');
  ${bloco('const _OE_ITEM =', "reposicao: 'Reposição' };")}
  ${['_normNome', '_skuBaseDaOS', '_oeLinhaDaOS', '_oeGruposDaFase', '_expFasesDaOS', '_expFasesDaCarga', '_expItensPorFase'].map(pegaFuncao).join('\n')}
  const ordenarComponentesPorFase = () => [];
  return { _expItensPorFase, _expFasesDaCarga, _oeGruposDaFase };
`)({ desenhos: [], modelos: [] });

const os = (sku, fases, componentes) => ({
  id: sku, skuOverride: sku, fases: fases.map((nome, i) => ({ ordem: i + 1, nome })), componentes: componentes || []
});
const lista = (o, carga) => {
  const c = carga || { osId: o.id };
  const g = api._expItensPorFase(o, c, api._expFasesDaCarga(c, o));
  return g.map(x => '[' + x.titulo + '] ' + x.itens.map(i => i.nome + (i.porPeca > 1 ? 'x' + i.porPeca : '')).join(', ')).join(' ');
};

console.log('-- a lista de cada linha --');
const cmLisa = lista(os('CM.LISA-PRE', ['Corpo', 'Gola', 'Viés']));
ok('1. CM.LISA: frente, costa, mangas, ribana, viés e tecido de reposição',
   cmLisa === '[Corpo] Frente, Costa, Mangasx2 [Gola] Ribana [Viés] Viés [Reposição] Tecido de reposição', cmLisa);
const cmRec = lista(os('CM.REC-VERM', ['Corpo Parte 1', 'Corpo Parte 2', 'Gola', 'Viés']));
ok('2. CM.REC: a CM.LISA + frente parte 2, na fase Corpo Parte 2',
   /\[Corpo Parte 2\] Frente parte 2 /.test(cmRec) && /\[Corpo Parte 1\] Frente, Costa, Mangasx2/.test(cmRec) && !/parte 3/.test(cmRec), cmRec);
const cmTri = lista(os('CM.TRI-CAQUI', ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Gola', 'Viés']));
ok('3. CM.TRI: peça por peça — frente 1, 2, 3; costa; mangas — num bloco só, depois ribana e viés',
   cmTri === '[Corpo] Frente parte 1, Frente parte 2, Frente parte 3, Costa, Mangasx2 [Gola] Ribana [Viés] Viés [Reposição] Tecido de reposição', cmTri);
const bmLisa = lista(os('BM.LISA-PRE', ['Corpo', 'Barra/Punhos', 'Viés']));
ok('4. BM.LISA: frente, costa, mangas, barra, punhos, viés e tecido de reposição',
   bmLisa === '[Corpo] Frente, Costa, Mangasx2 [Barra/Punhos] Barra, Punhosx2 [Viés] Viés [Reposição] Tecido de reposição', bmLisa);
const bmTri = lista(os('BM.TRI-BEGE', ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Forro de capuz', 'Barra/Punhos', 'Viés']));
ok('5. BM.TRI: peça por peça — frente 1, 2, 3; costa 1, 2, 3; mangas 1, 2, 3; capuz — num bloco só',
   /^\[Corpo\] Frente parte 1, Frente parte 2, Frente parte 3, Costa parte 1, Costa parte 2, Costa parte 3, Mangas parte 1x2, Mangas parte 2x2, Mangas parte 3x2, Capuzx2 \[Forro de capuz\]/.test(bmTri), bmTri);
ok('5d. o capuz nao engole o forro do capuz (componente "Forro do capuz" fica no Forro)',
   (bmTri.match(/Capuz/g) || []).length === 1 && /Forro de capuzx2/.test(bmTri), bmTri);
{
  const o = os('BM.TRI-BEGE', ['Corpo Parte 1', 'Corpo Parte 2', 'Corpo Parte 3', 'Forro de capuz', 'Barra/Punhos', 'Viés']);
  o.fases[1].corNome = 'Mostarda';
  const soParte2 = lista(o, { osId: 'x', fases: [2], pacotes: [{ tam: 'M', tom: null }] });
  ok('5a. BM.TRI em carga só da Corpo Parte 2: só as três peças da parte 2, com a cor dela',
     soParte2 === '[Corpo] Frente parte 2 (Mostarda), Costa parte 2 (Mostarda), Mangas parte 2 (Mostarda)x2', soParte2);
}
ok('5b. BM.TRI: o forro de capuz, na fase Forro de capuz, 2 por peça',
   /\[Forro de capuz\] Forro de capuzx2 \[Barra\/Punhos\]/.test(bmTri), bmTri);
ok('5c. a BM.LISA nao ganha forro de capuz', !/Forro/.test(bmLisa), bmLisa);
{
  const trocada = lista(os('BM.LISA-PRE', ['Viés', 'Barra/Punhos', 'Corpo']));
  ok('5e. a ordem e a da lista (frente, costa, mangas, barra, punhos, viés, reposição), e nao a das fases da OS',
     trocada === '[Corpo] Frente, Costa, Mangasx2 [Barra/Punhos] Barra, Punhosx2 [Viés] Viés [Reposição] Tecido de reposição', trocada);
}
ok('6. o nome e FRENTE parte 2/3, nao corpo parte 2/3', !/Corpo parte [23]/.test(cmTri + bmTri + cmRec));

console.log('-- o que muda a lista --');
const semVies = lista(os('CM.LISA-PRE', ['Corpo', 'Gola']));
ok('7. a OS sem fase de viés ainda lista o viés, num grupo com o nome dele', /\[Viés\] Viés/.test(semVies), semVies);
const umaManga = lista(os('CM.LISA-PRE', ['Corpo', 'Gola', 'Viés'], [{ nome: 'Mangas Camiseta', qtdPorPeca: 1 }]));
ok('8. quantas por peça vêm do componente da OS (relatório do rodapé) quando há',
   /Mangas,|Mangas \[/.test(umaManga) || /Frente, Costa, Mangas \[/.test(umaManga), umaManga);
const soCorpo = lista(os('CM.LISA-PRE', ['Corpo', 'Gola', 'Viés']), { osId: 'x', fases: [1], pacotes: [{ tam: 'M', tom: null }] });
ok('9. carga parcial só com o Corpo e sem reposição: só frente, costa e mangas',
   soCorpo === '[Corpo] Frente, Costa, Mangasx2', soCorpo);
const comRepos = lista(os('CM.LISA-PRE', ['Corpo', 'Gola', 'Viés']), { osId: 'x', fases: [2], pacotes: [], reposicao: true });
ok('10. carga com o pacote de reposição leva o tecido de reposição', comRepos === '[Gola] Ribana [Reposição] Tecido de reposição', comRepos);
ok('11. "Corpo + Gola" responde pelo corpo e pela ribana; "Corpo 2" é a parte 2',
   api._oeGruposDaFase('Corpo + Gola').join() === 'corpo1,ribana' && api._oeGruposDaFase('Corpo 2').join() === 'corpo2'
   && api._oeGruposDaFase('Punhos/Barra').join() === 'barra');
ok('12. a folha de OE usa o checklist no lugar da linha corrida das fases',
   /const itensHtml = _expItensFaseHtml\(o, i\.carga, fi\);\s*const fasesHtml = itensHtml \|\|/.test(src));

console.log('-- sem espaco excedente --');
{
  const duas = new Function(pegaFuncao('_expDuasColunas') + 'return _expDuasColunas;')();
  const q = n => '<div>' + '<tr>'.repeat(n) + '</div>';
  const h = duas([q(20), q(8), q(8)]);
  const cols = h.split('<div class="col">').slice(1);
  ok('13. perna larga: o quadro grande numa coluna, os dois menores na outra',
     cols.length === 2 && (cols[0].match(/<tr>/g) || []).length === 20 && (cols[1].match(/<tr>/g) || []).length === 16, h);
  ok('14. um quadro so nao se divide', !/duas-col/.test(duas([q(5)])));
  ok('15. perna vazia ao lado de cheia vira uma linha, e a cheia ocupa a largura toda',
     /if \(vIda === vVolta\) return/.test(src) && /exp-print-pernas uma-so/.test(src) && /exp-print-perna sem-carga/.test(src));
  ok('16. a observacao do aviamento vai na mesma linha do item',
     /_aviItemTexto\(a\)\)\}\$\{a\.obs \? ` <span/.test(src));
}

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
