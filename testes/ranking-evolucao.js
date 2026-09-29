/* A EVOLUCAO NO TEMPO DO RANKING (28/09/2026, Junior: "insira abaixo do quadro
   tamanho x cor, um grafico que mostre a evolucao do volume de cada resultado
   da tabela de acordo com o tempo, mostrando em dias, semanas, meses ou ano de
   acordo com o filtro").

   As contas do grafico sao recortadas do app.js e rodadas com fatos de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const linha = re => (src.match(re) || [''])[0];

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

// O cadastro de cores de mentira: o nome longo (com tecido) e o curto que o
// grafico usa. corNomeCurto entra dublado: tira o " Malha"/" Moletom" do fim.
const STATE = { cores: [
  { nome: 'Preto Malha Algodão', hex: '#111111' },
  { nome: 'Preto Moletom', hex: '#222222' },
  { nome: 'Branco Malha Algodão', hex: '#ffffff' },
  { nome: 'Marinho Malha Algodão' }             // sem hex
] };
const api = new Function('STATE', `
  function corNomeCurto(n) { return String(n || '').replace(/ (Malha Algodão|Moletom)$/, ''); }
  function esc(s) { return String(s); }
  ${linha(/const RANK_CORES_SERIE = [^\n]*/)}
  ${linha(/const _rankEscalaDoFiltro = [^\n]*/)}
  ${linha(/const RANK_COR_OUTROS = [^\n]*/)}
  ${pegaFuncao('_rankSomaDias')}
  ${pegaFuncao('_rankBalde')}
  ${pegaFuncao('_rankBaldes')}
  ${pegaFuncao('_rankHexDaCor')}
  ${pegaFuncao('_rankingSeries')}
  ${pegaFuncao('_rankTopoEixo')}
  ${pegaFuncao('_rankDivisoesEixo')}
  return { _rankEscalaDoFiltro, _rankBalde, _rankBaldes, _rankingSeries, _rankTopoEixo, _rankDivisoesEixo };
`)(STATE);

console.log('-- a escala segue o filtro --');
ok('1. um mes se le por dia', api._rankEscalaDoFiltro('2026', '2026-09') === 'dia');
ok('2. um ano se le por mes', api._rankEscalaDoFiltro('2026', '') === 'mes');
ok('3. a fabrica inteira se le por mes', api._rankEscalaDoFiltro('', '') === 'mes');

console.log('-- os intervalos --');
ok('4. a semana comeca na segunda (qui 17/09 -> seg 14/09)', api._rankBalde('2026-09-17', 'semana') === '2026-09-14');
ok('5. domingo e o fim da semana (dom 20/09 -> seg 14/09)', api._rankBalde('2026-09-20', 'semana') === '2026-09-14');
const dias = api._rankBaldes('2026-09-01', '2026-09-30', 'dia');
ok('6. setembro tem 30 dias, vazios inclusive', dias.length === 30 && dias[29] === '2026-09-30', dias.length);
const meses = api._rankBaldes('2025-11-10', '2026-02-03', 'mes');
ok('7. os meses viram o ano', meses.join(',') === '2025-11,2025-12,2026-01,2026-02', meses);
ok('8. os anos', api._rankBaldes('2024-05-01', '2026-01-01', 'ano').join(',') === '2024,2025,2026');
ok('9. a semana que atravessa o mes', api._rankBaldes('2026-08-30', '2026-09-08', 'semana').join(',') === '2026-08-24,2026-08-31,2026-09-07');

console.log('-- as series --');
const fatos = [
  { tamanho: 'P', data: '2026-09-01', produtos: 10.4 },
  { tamanho: 'M', data: '2026-09-01', produtos: 20 },
  { tamanho: 'P', data: '2026-09-03', produtos: 5 },
  { tamanho: 'G', data: '2026-09-03', produtos: 7 },
  { tamanho: 'P', data: '', produtos: 99 }   // OS sem data: fica fora do tempo
];
const b3 = api._rankBaldes('2026-09-01', '2026-09-03', 'dia');
const s = api._rankingSeries(fatos, 'tamanho', ['M', 'P', 'G'], b3, 'dia');
ok('10. uma serie por resultado da tabela, na ordem dela', s.map(x => x.rotulo).join(',') === 'M,P,G', s);
ok('11. o dia sem OS entra com zero', s[1].valores.join(',') === '10,0,5', s[1].valores);
ok('12. a cor segue a ordem da tabela', s[0].cor === '#2a78d6' && s[1].cor === '#eb6834');
const t = api._rankingSeries(fatos, '', [], b3, 'dia');
ok('13. sem variavel, uma linha so: o total', t.length === 1 && t[0].valores.join(',') === '30,0,12', t);
const muitos = 'ABCDEFGHIJ'.split('');
const fm = muitos.map((r, i) => ({ cor: r, data: '2026-09-01', produtos: 100 - i }));
const sm = api._rankingSeries(fm, 'cor', muitos, ['2026-09-01'], 'dia');
ok('14. mais de oito: sete ficam e o resto vira Outros', sm.length === 8 && sm[7].outros && /Outros \(3\)/.test(sm[7].rotulo), sm.map(x => x.rotulo));
ok('15. Outros soma o que sobrou (93+92+91)', sm[7].valores[0] === 276, sm[7].valores);
ok('16. a soma das linhas bate com o total', sm.reduce((a, x) => a + x.valores[0], 0) === fm.reduce((a, f) => a + f.produtos, 0));

console.log('-- o ponto na cor do produto (29/09/2026) --');
const fc = [
  { cor: 'Preto', data: '2026-09-01', produtos: 5 },
  { cor: 'Branco', data: '2026-09-01', produtos: 4 },
  { cor: 'Marinho', data: '2026-09-01', produtos: 3 }
];
const sc = api._rankingSeries(fc, 'cor', ['Preto', 'Branco', 'Marinho'], ['2026-09-01'], 'dia');
ok('19. o ponto do Preto sai na cor do cadastro (o primeiro Preto com hex)', sc[0].ponto === '#111111', sc[0]);
ok('20. o Branco sai branco', sc[1].ponto === '#ffffff', sc[1]);
ok('21. cor sem hex no cadastro: sem ponto proprio (fica a cor da linha)', sc[2].ponto === '', sc[2]);
ok('22. a LINHA continua na paleta, separando as series', sc[0].cor === '#2a78d6' && sc[1].cor === '#eb6834', sc.map(x => x.cor));
const st = api._rankingSeries(fatos, 'tamanho', ['M', 'P', 'G'], b3, 'dia');
ok('23. serie que nao e cor (tamanho) nao ganha ponto de produto', st.every(x => !x.ponto), st.map(x => x.ponto));

console.log('-- o eixo --');
ok('17. topo redondo acima do maior', api._rankTopoEixo(2370) === 2500 && api._rankTopoEixo(11400) === 20000, [api._rankTopoEixo(2370), api._rankTopoEixo(11400)]);
ok('18. 2.500 em 5 faixas (500), 20.000 em 4 (5.000)', api._rankDivisoesEixo(2500) === 5 && api._rankDivisoesEixo(20000) === 4);

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
