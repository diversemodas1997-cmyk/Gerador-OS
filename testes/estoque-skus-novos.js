/* Rode com:  node testes/estoque-skus-novos.js

   O SKU NOVO ENTRA SOZINHO NO ESTOQUE DE PRODUTOS (28/10/2026, Jones: "insira
   função de atualização automática do Estoque de produtos, para sincronizar
   novos SKU quando forem cadastrados"). O que este teste protege:
     · o desenho com SKU completo (CM.LISA-PRE) continua entrando como está;
     · o desenho sem SKU, ou só com a linha (CJM), compõe linha + Sigla SKU da
       cor principal — a linha do desenho, senão a do modelo — e entra;
     · sem sigla na cor, não inventa SKU;
     · cadastrou o desenho (ou deu a sigla à cor) depois, o SKU aparece na
       próxima leitura, sem botão de sincronizar;
     · gravar desenho/modelo/cor/OS com a tela de consulta do estoque aberta a
       redesenha (gancho do saveState no app.js).
   Carrega o estoque-confeccao.js de verdade. */
const fs = require('fs');
const path = require('path');
const mod = fs.readFileSync(path.join(__dirname, '..', 'estoque-confeccao.js'), 'utf8');
const app = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function carrega(STATE) {
  const window = {};
  Object.assign(globalThis, {
    document: { getElementById: () => null, querySelector: () => null },
    exigirEdicao: () => true, toast: () => {}, saveState: async () => {}, desfazerNomearAcao: () => {},
    confirm: () => true, goto: () => {}
  });
  new Function('window', 'STATE', 'formatDate', 'esc', 'uid', 'totaisPorTamanhoTomOS', 'skusDaOS', '_statusOS',
    'catalogoSkus', 'currentRole', mod)(
    window, STATE, iso => iso, s => String(s == null ? '' : s), () => 'id',
    () => ({ colTotal: () => 0 }), () => [], () => '', [], 'admin');
  return window._acabInterno;
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const STATE = {
  meta: {}, acabProdutos: [], acabMov: [], ordens: [],
  cores: [{ id: 'c1', nome: 'Preto', siglaSku: 'PRE' }, { id: 'c2', nome: 'Rosa chá', siglaSku: '' }],
  modelos: [{ id: 'm1', nome: 'Calça Moletom', skuLinha: 'CJM' }, { id: 'm2', nome: 'Short tactel', skuLinha: 'SM.LISA' }],
  desenhos: [
    { id: 'd1', desc: 'Camiseta Básica | Preto', skuLinha: 'CM.LISA-PRE', modeloId: 'x', corPrincipalId: 'c1' },
    { id: 'd2', desc: 'Calça Moletom | Preto', skuLinha: 'CJM', modeloId: 'm1', corPrincipalId: 'c1' },
    { id: 'd3', desc: 'Short tactel | Preto', skuLinha: '', modeloId: 'm2', corPrincipalId: 'c1' },
    { id: 'd4', desc: 'Short tactel | Rosa chá', skuLinha: '', modeloId: 'm2', corPrincipalId: 'c2' }
  ]
};
const a = carrega(STATE);
const skus = () => a.produtos().map(p => p.sku);

ok('1. desenho com SKU completo entra como esta', skus().includes('CM.LISA-PRE'), skus());
ok('2. desenho so com a linha compoe com a sigla da cor', skus().includes('CJM-PRE'), skus());
ok('3. desenho sem SKU usa a linha do modelo', skus().includes('SM.LISA-PRE'), skus());
ok('4. cor sem sigla nao inventa SKU', !skus().some(s => /^SM\.LISA-?$/.test(s)) && skus().length === 3, skus());
const p = a.produtos().find(x => x.sku === 'CJM-PRE');
ok('5. o SKU composto vem do programa, com a descricao do desenho', p && p.doPrograma && p.descricao === 'Calça Moletom | Preto', p);

(async () => {
await Promise.resolve(); // a lista é refeita a cada rodada de código
STATE.desenhos.push({ id: 'd5', desc: 'Bermuda | Preto', skuLinha: 'BERM', modeloId: 'm1', corPrincipalId: 'c1' });
STATE.cores[1].siglaSku = 'ROSA';
ok('6. desenho cadastrado depois aparece na proxima leitura', skus().includes('BERM-PRE'), skus());
ok('7. dar a sigla a cor faz o SKU aparecer', skus().includes('SM.LISA-ROSA'), skus());

ok('8. saveState redesenha a tela do estoque ao gravar cadastro de SKU',
  /_CHAVES_SKU_ESTOQUE\.includes\(key\)[\s\S]{0,200}renderEstoqueConfeccao\(pg\)/.test(app)
  && /const _CHAVES_SKU_ESTOQUE = \['desenhos', 'modelos', 'cores', 'ordens'\]/.test(app));
ok('9. so as telas de consulta (nao as de lancamento)',
  /const _TELAS_SKU_ESTOQUE = \['acab-estoque', 'acab-produtos', 'acab-cmp'\]/.test(app));

})().then(() => {
console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
});
