/* Rode com:  node testes/oe-fracoes-do-lote.js

   O RESTO DA MESMA OS, NOS OUTROS DIAS — a linha nova de cada quadro da folha
   de OE (18/09/2026, Junior).

   Um lote repartido vira duas ou três cargas em dias diferentes, e cada uma
   ganha o seu quadro na folha. Até aqui os quadros não sabiam um do outro: quem
   recebia em São Carlos lia "OS 0537 · G, GG" e não tinha como saber se o P e o
   M já tinham chegado ou se ainda vinham.

   O que este teste guarda é o que a linha NÃO pode dizer, que é onde ela
   estraga em vez de ajudar:

     · a própria carga do quadro — repetiria a tabela logo acima;
     · uma expedição CANCELADA — mandaria a doca esperar um pacote que não sai;
     · a data ORIGINAL de uma ocorrência remarcada — a fração viaja no dia novo,
       e é o dia novo que quem confere vai procurar;
     · nada, quando a OS não foi repartida: o silêncio quer dizer "esta carga é
       a OS inteira", e é a maioria dos quadros.

   Recorta as funções do app.js de verdade. */
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
  // 'function esc' sozinho casaria com esconderAlertaSalvamento, que vem antes
  // no arquivo: o indexOf pega o primeiro que comeca igual.
  corta('function esc(s)'),
  corta('function formatDate'),
  corta('function _expChavePacote'),
  corta('function _expContarPacotes'),
  corta('function _expRotuloPacote'),
  corta('function _expCancelSet'),
  corta('function _expDataEfetivaCarga'),
  corta('function _expOutrasFracoesOS'),
  corta('function _expOutrasFracoesTexto')
].join('\n');

function comMotor(estado) {
  const fn = new Function('STATE', `
    ${motor}
    return { _expOutrasFracoesOS, _expOutrasFracoesTexto };
  `);
  return fn(estado);
}

let falhas = 0;
function ok(nome, cond, extra) {
  if (cond) { console.log('  ok  ' + nome); return; }
  falhas++;
  console.log('  FALHOU  ' + nome + (extra ? '  -> ' + extra : ''));
}

const OS = { id: 'os1', os: '0537' };
// Três cargas da mesma OS: a de hoje (a do quadro) e duas outras.
const estadoBase = () => ({
  expedicaoExcecoes: [],
  expedicaoCargas: [
    { id: 'c1', osId: 'os1', janelaId: 'j1', data: '2026-09-10', perna: 'ida',
      pacotes: [{ tam: 'P', tom: 1 }, { tam: 'M', tom: 1 }], reposicao: true, feita: true },
    { id: 'c2', osId: 'os1', janelaId: 'j1', data: '2026-09-17', perna: 'ida',
      pacotes: [{ tam: 'G', tom: 1 }, { tam: 'GG', tom: 1 }] },
    { id: 'c3', osId: 'os1', janelaId: 'j1', data: '2026-09-24', perna: 'ida',
      pacotes: [{ tam: 'G1', tom: 1 }] },
    // Outra OS na mesma janela: nao tem nada a ver com o lote deste quadro.
    { id: 'x1', osId: 'os2', janelaId: 'j1', data: '2026-09-17', perna: 'ida',
      pacotes: [{ tam: 'P', tom: 1 }] }
  ]
});
const cargaDe = (st, id) => st.expedicaoCargas.find(c => c.id === id);

console.log('-- quais frações entram --');
let st = estadoBase();
let api = comMotor(st);
let f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('1. as OUTRAS cargas da mesma OS entram, a do quadro nao',
   f.length === 2 && f.every(x => x.data !== '2026-09-17'), JSON.stringify(f));
ok('2. e carga de outra OS nunca entra',
   !f.some(x => x.pacotes === 'P · tom 1'), JSON.stringify(f));
ok('3. saem em ordem de calendario',
   f[0].data === '2026-09-10' && f[1].data === '2026-09-24', JSON.stringify(f.map(x => x.data)));
ok('4. os pacotes daquele dia sao listados, com a reposicao',
   f[0].pacotes === 'P · tom 1, M · tom 1, reposição', f[0].pacotes);
ok('5. a carga ja marcada como feita vem marcada', f[0].feita === true && f[1].feita === false,
   JSON.stringify(f.map(x => x.feita)));

console.log('');
console.log('-- o que a linha NAO pode dizer --');
st = estadoBase();
st.expedicaoExcecoes = [{ janelaId: 'j1', data: '2026-09-24', tipo: 'cancelada' }];
api = comMotor(st);
f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('6. expedicao cancelada sai da lista — aquele lote nao vai sair',
   f.length === 1 && f[0].data === '2026-09-10', JSON.stringify(f));

st = estadoBase();
st.expedicaoExcecoes = [{ janelaId: 'j1', data: '2026-09-24', tipo: 'remarcada', novaData: '2026-09-25' }];
api = comMotor(st);
f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('7. remarcada: sai a data EFETIVA, nao a original',
   f.some(x => x.data === '2026-09-25') && !f.some(x => x.data === '2026-09-24'),
   JSON.stringify(f.map(x => x.data)));

st = { expedicaoExcecoes: [], expedicaoCargas: [estadoBase().expedicaoCargas[1]] };
api = comMotor(st);
ok('8. OS numa carga so: a linha nao existe',
   api._expOutrasFracoesTexto(OS, cargaDe(st, 'c2')) === '',
   api._expOutrasFracoesTexto(OS, cargaDe(st, 'c2')));

console.log('');
console.log('-- casos de borda do pacote --');
st = estadoBase();
// Carga ANTIGA: so o numero de volumes, sem composicao. Ela e o lote inteiro.
st.expedicaoCargas.push({ id: 'c4', osId: 'os1', janelaId: 'j2', data: '2026-09-30', perna: 'ida', volumes: 9 });
api = comMotor(st);
f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('9. carga antiga (sem pacotes) se anuncia como lote inteiro',
   f.some(x => x.data === '2026-09-30' && x.pacotes === 'lote inteiro'), JSON.stringify(f));

st = estadoBase();
st.expedicaoCargas[0].pacotes = [{ tam: 'P', tom: 1 }, { tam: 'P', tom: 1 }, { tam: 'M', tom: 2 }];
st.expedicaoCargas[0].reposicao = false;
api = comMotor(st);
f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('10. vagas repetidas do mesmo tamanho contam juntas',
   f[0].pacotes === '2× P · tom 1, M · tom 2', f[0].pacotes);

st = estadoBase();
st.expedicaoCargas[2].perna = 'volta';
api = comMotor(st);
f = api._expOutrasFracoesOS(OS, cargaDe(st, 'c2'));
ok('11. fracao na outra perna vem identificada',
   f.some(x => x.perna === 'volta' && x.outraPerna === true), JSON.stringify(f));

console.log('');
console.log('-- o texto que vai ao papel --');
st = estadoBase();
api = comMotor(st);
let txt = api._expOutrasFracoesTexto(OS, cargaDe(st, 'c2'));
ok('12. traz o dia e os tamanhos de cada outra fracao',
   /10\/09\/2026/.test(txt) && /24\/09\/2026/.test(txt) && /G1/.test(txt) && /P · tom 1/.test(txt), txt);
ok('13. diz quem ja foi feita', /10\/09\/2026<\/b> \(feita\)/.test(txt), txt);
ok('14. e nao repete a carga deste quadro', !/17\/09\/2026/.test(txt), txt);
ok('15. sai na classe propria da folha (.fracoes), fora da tarja do recado',
   /^<div class="fracoes">/.test(txt), txt);

// O quadro da folha tem TRES formatos (sem grade, carga parcial, lote cheio) e
// o recado tem de sair nos tres — pela mesma razao da observacao da alocacao.
// E sai por ULTIMO, depois da tarja amarela: o recado e de alguem para esta
// carga e le-se primeiro; o resto do lote e referencia, e fecha o quadro.
const quadro = recorte('const osPrint = (i) =>', '\n  };', 'o quadro da folha de OE');
ok('16. a linha entra nos tres formatos do quadro, sempre ABAIXO da observacao',
   (quadro.match(/\$\{fasesHtml\}\$\{obsHtml\}\$\{fracHtml\}/g) || []).length === 3,
   String((quadro.match(/\$\{fracHtml\}/g) || []).length));
ok('17. e a folha calcula a linha a partir da carga do quadro',
   /const fracHtml = _expOutrasFracoesTexto\(o, i\.carga\)/.test(quadro));

const css = fs.readFileSync(path.join(__dirname, '..', 'styles.css'), 'utf8');
ok('18. a classe existe no styles.css', /\.exp-print-os > \.fracoes\s*\{/.test(css));

console.log('');
if (falhas) { console.log(falhas + ' teste(s) falharam'); process.exit(1); }
console.log('tudo certo');
