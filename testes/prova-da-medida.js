/* Rode com:  node testes/prova-da-medida.js

   DE ONDE VEIO ESTE NÚMERO? — o aviso ao gerar a OS.

   A medida de enfesto digitada à mão é a origem do pedido curto ao fornecedor:
   a bobina a menos só aparece no dia do corte, e aí o lote já parou. O PDF do
   encaixe é a prova de onde o número saiu, e mora em `fase.risco`.

   Desde 20/08/2026, gerar OS cuja grade tem fase com medida SEM esse PDF mostra
   um aviso dizendo exatamente qual fase está sem prova — com a opção de seguir
   assim mesmo, porque 53 das 111 grades em uso ainda não têm o relatório
   exportado e uma trava sem saída pararia a produção por pendência de cadastro.

   O que este teste protege:

     · VIÉS e GOLA nunca são cobrados. Nenhuma grade vai ter PDF de viés (é tira
       cortada da sobra do mesmo pano, não encaixe) e Junior decidiu o mesmo para
       a gola. Cobrar prova impossível vira hábito de clicar "continuar", e aí o
       aviso não vale para nada;
     · fase SEM medida não é cobrada: onde não há número não há o que provar;
     · fase com medida E com PDF passa calada.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}
// Uma linha inteira do app.js, para os `const ... = new Set([...])`.
const linha = comeco => {
  const i = src.indexOf(comeco);
  if (i < 0) { console.error('nao achei ' + comeco + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n', i) + 1);
};

const api = new Function('STATE', `
  ${linha('const _EXC_LIGACAO = new Set')}
  ${linha('const _PAL_VIES = new Set')}
  ${linha('const _PAL_GOLA = new Set')}
  ${recorte('function _normNome(', '_normNome')}
  ${recorte('function _normFaseNome(', '_normFaseNome')}
  ${recorte('function _faseSoDe(', '_faseSoDe')}
  ${recorte('function _faseSemEncaixe(', '_faseSemEncaixe')}
  ${recorte('function fasesSemProvaDeMedida', 'a busca das fases sem prova')}
  return fasesSemProvaDeMedida;
`);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const fase = (nome, comp, larg, risco) => ({ nome, comp, larg, risco: risco || '' });
const com = (...fases) => api({ grades: [{ id: 'g1', nome: 'M-GG-G3 | CM.LISA | 117cm', fases }] })('g1');

const PDF = 'CM.LISA/M-GG-G3/117 cm/CM.LISA CORPO M-GG-G3.pdf';

console.log('-- o caso que motivou o aviso --');
ok('1. corpo com medida e SEM pdf e cobrado',
  com(fase('Corpo', '3.88', '1.170')).map(f => f.nome).join() === 'Corpo',
  com(fase('Corpo', '3.88', '1.170')));
ok('2. o mesmo corpo COM pdf passa calado',
  com(fase('Corpo', '3.88', '1.170', PDF)).length === 0);

console.log('');
console.log('-- vies e gola nunca sao cobrados --');
ok('3. vies sem pdf nao e cobrado', com(fase('Viés', '4', '1.17')).length === 0);
ok('4. gola sem pdf nao e cobrada', com(fase('Gola', '0.50', '0.67')).length === 0);
ok('5. nem com a grafia sem acento', com(fase('Vies', '4', '1.17')).length === 0);
ok('6. nem em maiuscula ou composto',
  com(fase('VIÉS DA GOLA', '4', '1.17')).length === 0);
// "Corpo + Gola" e o pano do corpo que leva a gola junto: e encaixe, tem
// relatorio (3 no cadastro, todos com prova) e nao pode ser isentado so porque
// a palavra "gola" aparece no nome.
ok('7. "Corpo + Gola" NAO e isento — e encaixe',
  com(fase('Corpo + Gola', '3.88', '1.170')).map(f => f.nome).join() === 'Corpo + Gola',
  com(fase('Corpo + Gola', '3.88', '1.170')));
ok('7b. "Gola e Vies" continua isento — os dois sao convencao',
  com(fase('Gola e Viés', '0.50', '0.67')).length === 0,
  com(fase('Gola e Viés', '0.50', '0.67')));

console.log('');
console.log('-- onde nao ha numero, nao ha o que provar --');
ok('8. fase sem medida nenhuma nao e cobrada', com(fase('Corpo', '', '')).length === 0);
ok('9. so o comprimento nao basta para cobrar', com(fase('Corpo', '3.88', '')).length === 0);
ok('10. so a largura tambem nao', com(fase('Corpo', '', '1.17')).length === 0);
ok('11. zero e o mesmo que vazio', com(fase('Corpo', '0', '0')).length === 0);

console.log('');
console.log('-- a grade inteira --');
const varias = com(
  fase('Corpo Parte 1', '5.15', '1.165', PDF),
  fase('Corpo Parte 2', '0.44', '1.165'),
  fase('Corpo Parte 3', '1.45', '1.165'),
  fase('Gola', '0.60', '0.65'),
  fase('Viés', '5.20', '1.17'));
ok('12. cobra so as partes de corpo sem pdf, e nomeia cada uma',
  varias.map(f => f.nome).join(' / ') === 'Corpo Parte 2 / Corpo Parte 3',
  varias.map(f => f.nome));
ok('13. grade toda provada nao gera aviso',
  com(fase('Corpo', '3.88', '1.170', PDF), fase('Gola', '0.50', '0.67')).length === 0);
ok('14. grade que nao existe nao trava a gravacao',
  api({ grades: [] })('some') .length === 0);
ok('15. risco so com espaco nao vale como prova',
  com(fase('Corpo', '3.88', '1.170', '   ')).length === 1);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
