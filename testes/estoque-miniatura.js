/* Rode com:  node testes/estoque-miniatura.js

   A MINIATURA DO DESENHO TÉCNICO NO ESTOQUE DE PRODUTOS (28/10/2026, Jones:
   "Insira no estoque de produtos coluna com miniatura do desenho técnico dos
   produtos"). Protege:
     · o SKU do desenho mostra a imagem do desenho;
     · sem desenho com o SKU, vale o desenho da OS que compõe o SKU;
     · a linha sozinha (modelo novo) usa um desenho da mesma linha;
     · produto sem desenho nenhum mostra "—";
     · Estoque atual e Produtos (SKUs) têm a coluna Desenho.
   Carrega o estoque-confeccao.js de verdade. */
const fs = require('fs');
const path = require('path');
const mod = fs.readFileSync(path.join(__dirname, '..', 'estoque-confeccao.js'), 'utf8');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const STATE = {
  meta: { acabConfig: { inicio: '2026-10-22T00:00:00.000Z' } }, acabProdutos: [], acabMov: [],
  cores: [{ id: 'c1', nome: 'Preto', siglaSku: 'PRE' }],
  modelos: [{ id: 'm1', nome: 'Calça', skuLinha: 'CJM' }, { id: 'm2', nome: 'Polo', skuLinha: 'PM.TRI' }],
  desenhos: [
    { id: 'd1', desc: 'Camiseta | Preto', skuLinha: 'CM.LISA-PRE', img: 'cm-pre.png' },
    { id: 'd2', desc: 'Calça | Preto', skuLinha: 'CJM', modeloId: 'm1', corPrincipalId: 'c1', img: 'cjm.png' },
    { id: 'd3', desc: 'Recortada', skuLinha: '', img: 'rec.png' }
  ],
  ordens: [{ id: 'o1', desenhoId: 'd3', modeloNome: 'Recortada' }]
};
const els = {};
Object.assign(globalThis, {
  document: { getElementById: id => (els[id] = els[id] || { id, value: '', innerHTML: '' }), querySelector: () => null, querySelectorAll: () => [] },
  exigirEdicao: () => true, temAcesso: () => true, toast: () => {}, saveState: async () => {}, desfazerNomearAcao: () => {},
  confirm: () => true, goto: () => {}, urlDesenho: v => 'https://srv/desenhos/' + v
});
const window = {};
new Function('window', 'STATE', 'formatDate', 'esc', 'uid', 'totaisPorTamanhoTomOS', 'skusDaOS', '_statusOS', 'catalogoSkus', 'currentRole', mod)(
  window, STATE, iso => iso, s => String(s == null ? '' : s), () => 'id', () => ({ colTotal: () => 0 }),
  o => (o.id === 'o1' ? ['CM.REC-PRE'] : []), () => '', [], 'admin');
const I = window._acabInterno;
const imgs = I.imagensDosSkus();
const cel = sku => I.celMiniatura(imgs, { sku });

ok('1. o SKU do desenho mostra a imagem dele', /src="https:\/\/srv\/desenhos\/cm-pre\.png"/.test(cel('CM.LISA-PRE')), cel('CM.LISA-PRE'));
ok('2. o desenho que compoe o SKU (linha + sigla) tambem', /cjm\.png/.test(cel('CJM-PRE')), cel('CJM-PRE'));
ok('3. sem desenho com o SKU, vale o desenho da OS', /rec\.png/.test(cel('CM.REC-PRE')), cel('CM.REC-PRE'));
ok('4. a linha sozinha usa um desenho da mesma linha', /cm-pre\.png/.test(cel('CM.LISA')), cel('CM.LISA'));
ok('5. produto sem desenho mostra "—"', !/<img/.test(cel('PM.TRI')) && /—/.test(cel('PM.TRI')));

window.renderEstoqueConfeccao('acab-estoque');
ok('6. o Estoque atual tem a coluna Desenho e a miniatura', /<th[^>]*>Desenho<\/th><th>SKU<\/th>/.test(els['acab-estoque-painel'].innerHTML)
  && /cm-pre\.png/.test(els['acab-estoque-painel'].innerHTML));
window.renderEstoqueConfeccao('acab-produtos');
ok('7. Produtos (SKUs) tambem', /<th[^>]*>Desenho<\/th><th>Código<\/th>/.test(els['acab-produtos-painel'].innerHTML)
  && /cjm\.png/.test(els['acab-produtos-painel'].innerHTML));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
