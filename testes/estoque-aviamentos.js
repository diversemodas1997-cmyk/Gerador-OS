/* ESTOQUE DE AVIAMENTOS (24/09/2026, Junior: "Insira um Estoque de aviamentos
   na barra lateral abaixo de Estoque de tecidos ... volume de entrada, volume
   de saída, volume residual (estocado), volume corrente e volume total ...
   quadro para cada tipo: Fio, linha, etiqueta, botão, viés ... cor e peso").

   A conta e recortada do app.js e rodada com lancamentos de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const tiposLinha = (src.match(/const AVIAMENTO_TIPOS = [^\n]*/) || [''])[0];

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const api = new Function(`
  ${tiposLinha}
  ${pegaFuncao('_normNome')}
  ${pegaFuncao('calcularEstoqueAviamentos')}
  return { calcularEstoqueAviamentos, AVIAMENTO_TIPOS };
`)();

console.log('-- os cinco tipos --');
ok('1. um quadro para cada tipo pedido',
   api.AVIAMENTO_TIPOS.join(',') === 'Fio,Linha,Etiqueta,Botão,Viés', api.AVIAMENTO_TIPOS);

/* Linha preta: 10 kg em agosto (antes do periodo), no periodo de setembro
   entram 5 e saem 3, e em outubro (depois do periodo) saem mais 4. Hoje e
   15/10, entao o corrente ja viu a saida de outubro. */
const mov = [
  { tipo: 'entrada', item: 'Linha', cor: 'Preto', kg: 10, data: '2026-08-20' },
  { tipo: 'entrada', item: 'Linha', cor: 'preto', kg: 5, data: '2026-09-05' },
  { tipo: 'saida', item: 'linha', cor: 'Preto', kg: 3, data: '2026-09-10' },
  { tipo: 'saida', item: 'Linha', cor: 'Preto', kg: 4, data: '2026-10-02' },
  { tipo: 'entrada', item: 'Botão', cor: 'Bege', kg: 1.25, data: '2026-09-30' },
  // Lancamento com data depois de hoje: ainda nao aconteceu.
  { tipo: 'entrada', item: 'Botão', cor: 'Bege', kg: 9, data: '2026-12-01' }
];
const r = api.calcularEstoqueAviamentos(mov, '2026-09-01', '2026-09-30', '2026-10-15');
const linha = r.find(x => x.item === 'Linha');
const botao = r.find(x => x.item === 'Botão');

console.log('-- os cinco volumes --');
ok('2. mesma cor com outra caixa cai na MESMA linha', r.filter(x => x.item === 'Linha').length === 1, r);
ok('3. ENTRADA e SAIDA contam so o periodo', linha.entrada === 5 && linha.saida === 3, linha);
ok('4. RESIDUAL = o de antes + entrada - saida (10 + 5 - 3 = 12)', linha.residual === 12, linha);
ok('5. TOTAL = o que ja havia + o que entrou (10 + 5 = 15)', linha.total === 15, linha);
ok('6. CORRENTE e o de hoje, e ve a saida de outubro (12 - 4 = 8)', linha.corrente === 8, linha);
ok('7. lancamento com data futura nao entra no corrente', botao.corrente === 1.25, botao);
ok('8. com o periodo terminando hoje, residual e corrente se encontram',
   (() => { const x = api.calcularEstoqueAviamentos(mov, '2026-09-01', '2026-10-15', '2026-10-15').find(y => y.item === 'Linha');
     return x.residual === x.corrente && x.corrente === 8; })());

console.log('-- a tela --');
ok('9. o item de menu fica logo abaixo do Estoque de tecidos',
   /data-page="estoque"[^>]*>Estoque de tecidos<\/a>\s*<a class="nav-btn" href="#estoque-aviamentos" data-page="estoque-aviamentos"[^>]*>Estoque de aviamentos<\/a>/.test(html));
ok('10. a pagina e o modal existem',
   /<section class="page hidden" data-page="estoque-aviamentos">/.test(html) && /id="modal-aviamento"/.test(html));
ok('11. a rota desenha a tela', /if \(page === 'estoque-aviamentos'\) renderEstoqueAviamentos\(\);/.test(src));
ok('12. a chave aviamentosMov e sincronizada e carregada (as duas listas de chaves)',
   (src.match(/'compraOCs','aviamentosMov'/g) || []).length === 2, (src.match(/'compraOCs','aviamentosMov'/g) || []).length);
ok('13. o lancamento guarda cor e peso', /cor: v\('ma-cor'\)\.trim\(\),\s*kg: Math\.round/.test(src));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
