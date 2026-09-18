/* Rode com:  node testes/horarios-expedicao.js

   OS HORÁRIOS DA EXPEDIÇÃO — os quatro do padrão e os de cada janela
   (18/09/2026, Junior).

   A fábrica tem dois turnos de caminhão, matinal e da tarde, e cada um vai e
   volta: quatro horários. Até aqui eles não existiam em lugar nenhum — o
   formulário da janela nascia com 08:00 e 17:00 escritos no meio do HTML,
   números que não vieram desta fábrica e que ninguém mudava sem mexer no
   código. Agora moram na configuração (Unidades e carga), ao lado das unidades
   e do limite de volume.

   O QUE ESTE TESTE GUARDA é a fronteira entre PADRÃO e VALOR GRAVADO:

     · o padrão preenche o formulário de uma janela nova e não manda em mais
       nada. Cada janela guarda o horário DELA, escrito no registro;
     · por isso mudar o padrão NÃO pode reescrever janela nenhuma: se o horário
       viesse do turno por referência, mudar a configuração mexeria em silêncio
       no dia de trabalho de expedições já planejadas, com OS alocada e folha
       impressa;
     · horário em branco ou pela metade cai no padrão de fábrica, em vez de
       gravar '' e fazer a próxima janela nascer com o campo vazio.

   Recorta as funções e o texto do app.js de verdade. */
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
const cortaObj = (nome) => recorte(nome, '\n};', nome) + '\n};';

const motor = [
  cortaObj('const EXP_CFG_PADRAO'),
  corta('function expCfg'),
  corta('function _expHora'),
  corta('function _expUsarTurnoJanela')
].join('\n');

// Os dois campos de hora do formulário da janela, dublados.
function comMotor(estado) {
  const campos = { 'ej-hora-ida': { value: '' }, 'ej-hora-volta': { value: '' } };
  const fn = new Function('STATE', 'campos', `
    const document = { getElementById: (id) => campos[id] || null };
    ${motor}
    return { EXP_CFG_PADRAO, expCfg, _expHora, _expUsarTurnoJanela };
  `);
  return { api: fn(estado, campos), campos };
}

let falhas = 0;
function ok(nome, cond, extra) {
  if (cond) { console.log('  ok  ' + nome); return; }
  falhas++;
  console.log('  FALHOU  ' + nome + (extra ? '  -> ' + extra : ''));
}

console.log('-- o padrao de fabrica --');
let t = comMotor({ meta: {} });
let cfg = t.api.expCfg();
ok('1. os quatro horarios existem na configuracao',
   cfg.horaIdaManha === '09:30' && cfg.horaVoltaManha === '09:30'
   && cfg.horaIdaTarde === '16:00' && cfg.horaVoltaTarde === '16:00', JSON.stringify(cfg));
ok('2. e sao os dois turnos que a fabrica roda hoje, nao o 08:00/17:00 do codigo antigo',
   cfg.horaIdaManha !== '08:00' && cfg.horaVoltaTarde !== '17:00', JSON.stringify(cfg));

console.log('');
console.log('-- o que a fabrica cadastrou manda --');
t = comMotor({ meta: { expedicao: { horaIdaManha: '07:15', horaVoltaTarde: '18:40' } } });
cfg = t.api.expCfg();
ok('3. horario cadastrado ganha do padrao', cfg.horaIdaManha === '07:15' && cfg.horaVoltaTarde === '18:40',
   JSON.stringify(cfg));
ok('4. e o que nao foi cadastrado continua no padrao',
   cfg.horaVoltaManha === '09:30' && cfg.horaIdaTarde === '16:00', JSON.stringify(cfg));

console.log('');
console.log('-- campo vazio e pergunta, nao horario --');
const h = t.api._expHora;
ok('5. hora valida passa', h('06:05', '09:30') === '06:05');
ok('6. vazio cai no padrao', h('', '09:30') === '09:30');
ok('7. nulo cai no padrao', h(null, '16:00') === '16:00');
ok('8. lixo cai no padrao', h('9h', '09:30') === '09:30' && h('930', '09:30') === '09:30',
   String(h('9h', '09:30')) + ' / ' + String(h('930', '09:30')));
// A conferencia e de FORMATO, nao de relogio: '99:99' passa. Quem digita e o
// <input type="time"> do navegador, que nao deixa passar hora que nao existe —
// e inventar aqui uma segunda validacao seria ter duas regras para a mesma
// coisa, com o risco de uma recusar o que a outra aceita.
ok('8b. e a conferencia e so de formato, como esta escrito no codigo',
   h('99:99', '09:30') === '99:99', String(h('99:99', '09:30')));

console.log('');
console.log('-- os botoes de turno do formulario --');
t = comMotor({ meta: { expedicao: { horaIdaManha: '09:30', horaVoltaManha: '10:10',
                                    horaIdaTarde: '16:00', horaVoltaTarde: '17:30' } } });
t.api._expUsarTurnoJanela('manha');
ok('9. Manha escreve o par da manha nos dois campos',
   t.campos['ej-hora-ida'].value === '09:30' && t.campos['ej-hora-volta'].value === '10:10',
   JSON.stringify(t.campos));
t.api._expUsarTurnoJanela('tarde');
ok('10. Tarde escreve o par da tarde',
   t.campos['ej-hora-ida'].value === '16:00' && t.campos['ej-hora-volta'].value === '17:30',
   JSON.stringify(t.campos));
t = comMotor({ meta: {} });
t.api._expUsarTurnoJanela('tarde');
ok('11. sem nada cadastrado, o botao usa o padrao de fabrica',
   t.campos['ej-hora-ida'].value === '16:00', JSON.stringify(t.campos));

console.log('');
console.log('-- o que a tela mostra e o que a tela grava --');
const modalCfg = recorte('function abrirModalExpConfig', '\n}', 'o modal de Unidades e carga');
['ex-hora-ida-manha', 'ex-hora-volta-manha', 'ex-hora-ida-tarde', 'ex-hora-volta-tarde'].forEach((id, i) => {
  ok((12 + i) + '. Unidades e carga tem o campo ' + id,
     modalCfg.includes(`id="${id}"`) && modalCfg.includes('type="time"'));
});
ok('16. e o bloco se anuncia como horarios PADRAO', /Horários padrão da expedição/.test(modalCfg));

const salvar = recorte("} else if (ctx.tipo === 'config')", "} else if (ctx.tipo === 'ocorrencia')", 'o salvar da configuracao');
ok('17. o salvar grava os quatro em meta.expedicao',
   /horaIdaManha, horaVoltaManha, horaIdaTarde, horaVoltaTarde/.test(salvar), salvar.slice(-300));
ok('18. e passa cada um pelo _expHora (vazio volta ao padrao)',
   (salvar.match(/_expHora\(v\('ex-hora-/g) || []).length === 4,
   String((salvar.match(/_expHora\(v\('ex-hora-/g) || []).length));

const modalJanela = recorte('function abrirModalExpJanela', '\n}', 'o modal da janela');
ok('19. a janela nova nasce com o horario da manha da configuracao',
   /id="ej-hora-ida" value="\$\{esc\(j \? \(j\.horaIda \|\| ''\) : _expHora\(cfg\.horaIdaManha/.test(modalJanela));
ok('20. e o 08:00/17:00 escrito no HTML sumiu',
   !modalJanela.includes("'08:00'") && !modalJanela.includes("'17:00'"), modalJanela.slice(0, 200));
ok('21. os dois botoes de turno estao no formulario',
   /_expUsarTurnoJanela\('manha'\)/.test(modalJanela) && /_expUsarTurnoJanela\('tarde'\)/.test(modalJanela));

// A janela guarda HORA, nunca "turno": o padrão preenche e vai embora.
const salvarJanela = recorte("const horaIda = v('ej-hora-ida')", 'saveState(\'expedicaoJanelas\')', 'o salvar da janela');
ok('22. a janela grava horaIda/horaVolta e nao guarda turno nenhum',
   /horaIda, horaVolta/.test(salvarJanela) && !/turno/i.test(salvarJanela), salvarJanela.slice(0, 200));

console.log('');
if (falhas) { console.log(falhas + ' teste(s) falharam'); process.exit(1); }
console.log('tudo certo');
