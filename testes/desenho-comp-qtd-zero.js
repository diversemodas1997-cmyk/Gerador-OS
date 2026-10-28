/* Rode com:  node testes/desenho-comp-qtd-zero.js

   QTD/PEÇA DOS COMPONENTES PADRÃO DO DESENHO: VAZIO E ZERO (28/10/2026,
   Jones: "Corrija o dropdown de número e insira a opção 0 ou vazio" — "faça
   isso para componentes padrão do desenho"). Antes a lista ia de 1 a 20, e
   vazio ou 0 viravam 1 ao gravar. Protege:
     · a lista tem "—" (vazio) e 0 antes do 1;
     · o componente sem quantidade aparece em "—", não em 1;
     · gravar vazio guarda null (não informado) e gravar 0 guarda 0;
     · na OS, vazio vale 1 (o padrão de sempre) e 0 não conta peça;
     · no risco de forro e ribana, 0 não conta e vazio segue o padrão.
   Lê o app.js de verdade. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// A lista, montada como no formulário do desenho.
const i = src.indexOf('const q = atual.qtdPorPeca;');
const trecho = src.slice(i, src.indexOf('return `<select class="m-comp-qtd"', i));
const lista = q => new Function('atual', trecho + 'return opts;')({ qtdPorPeca: q });
ok('1. a lista tem vazio e 0 antes do 1', /^<option value="" [^>]*>—<\/option><option value="0" >0<\/option><option value="1" >1</.test(lista(undefined)), lista(undefined).slice(0, 120));
ok('2. componente sem quantidade aparece em vazio', /<option value="" selected>—/.test(lista(undefined)) && /<option value="" selected>—/.test(lista('')));
ok('3. o 0 gravado aparece em 0', /<option value="0" selected>0/.test(lista(0)));
ok('4. o 2 gravado aparece em 2', /<option value="2" selected>2/.test(lista(2)) && !/value="" selected/.test(lista(2)));
ok('5. vai ate 20', /<option value="20" >20<\/option>$/.test(lista(undefined)));

// O que grava.
const grava = v => {
  const qtdEl = { value: v };
  return (qtdEl && qtdEl.value !== '' && Number.isFinite(parseFloat(qtdEl.value))) ? parseFloat(qtdEl.value) : null;
};
ok('6. o salvar usa a regra nova (vazio = null, 0 = 0)', src.includes("qtdPorPeca: (qtdEl && qtdEl.value !== '' && Number.isFinite(parseFloat(qtdEl.value))) ? parseFloat(qtdEl.value) : null")
  && grava('') === null && grava('0') === 0 && grava('3') === 3);

// Na OS: null vale 1, 0 vale 0 (a conta que monta os componentes da OS).
ok('7. a OS nova usa 1 para vazio e respeita o 0', /const qtdPorPeca = c\.qtdPorPeca != null \? c\.qtdPorPeca : 1;/.test(src));

// O risco de forro e ribana.
const j = src.indexOf('const qtdDoComp = c => {');
const qtdDoComp = new Function('return ' + src.slice(j + 'const qtdDoComp = '.length, src.indexOf('};', j) + 1))();
ok('8. no risco, 0 nao conta', qtdDoComp({ qtdPorPeca: 0, nome: 'Punho' }) === 0);
ok('9. no risco, vazio segue o padrao (manga/punho = 2, o resto = 1)', qtdDoComp({ qtdPorPeca: null, nome: 'Punho' }) === 2 && qtdDoComp({ qtdPorPeca: null, nome: 'Barra' }) === 1);
ok('10. o forro tambem respeita o 0', src.includes('const qty = v > 0 ? v : (v === 0 ? 0 : 1);'));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
