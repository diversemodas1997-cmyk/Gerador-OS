/* AVIAMENTO DAS OS: RESERVA E BAIXA AUTOMATICA (25/09/2026, Junior: "Programa
   deve dar baixa automaticamente na quantidade de estoque de aviamentos de
   acordo com as quantidades utilizadas em cada OS", "a regra de aviamentos
   reservado deve funcionar igual a de tecidos reservados para cada OS nao
   iniciada" e, perguntado quando sai: "na costura").

   A conta e recortada do app.js e rodada com OS de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const linha = re => { const m = src.match(re); if (!m) throw new Error('nao achei ' + re); return m[0]; };

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const STATE = {
  materiais: [{ id: 'et', desc: 'Etiqueta de tamanho' }, { id: 'co', desc: 'Cordão' }, { id: 'bo', desc: 'Botão' }],
  ordens: []
};
// "Total por tamanho" de mentira: o.tt = { m: 100, g: 50 }.
const TT = o => ({ tamanhos: Object.keys(o.tt || {}), colTotal: k => (o.tt || {})[k] || 0 });
const api = new Function('STATE', 'totaisPorTamanhoTomOS', '_statusOS', `
  ${linha(/const AVIAMENTO_TIPOS = [^\n]*/)}
  ${linha(/const AVIAMENTO_COR_ETIQUETA = [^\n]*/)}
  ${linha(/const AVI_BAIXA_DESDE = [^\n]*/)}
  ${linha(/const AVI_STATUS_RESERVA = \[[\s\S]*?\];/)}
  ${['_normNome', '_aviTipoDoMaterial', '_aviNecessidadeOS', '_aviDiaDoCarimbo', '_aviCosturaDaOS', '_aviDasOS'].map(pegaFuncao).join('\n')}
  return { _aviDasOS, _aviNecessidadeOS, _aviTipoDoMaterial, _aviCosturaDaOS };
`)(STATE, TT, o => o.st || 'nao-iniciado');

const ts = dia => new Date(dia + 'T10:00:00').getTime();
const etapas = ['Corte', 'Costura CM.LISA | Descalvado', 'Costura CM.LISA | São Carlos'];
const os = (num, extra) => Object.assign({
  id: 'o' + num, os: num, etapas, tt: { m: 100, g: 50 },
  aviamentos: [{ material: 'et', qtdPorPeca: 1 }, { material: 'co', qtdPorPeca: 1 }]
}, extra);
const costurada = (unid, dia) => ({ progresso: {
  etapasCheck: { ['Costura CM.LISA | ' + unid]: true },
  etapasSeq: { ['Costura CM.LISA | ' + unid]: ts(dia) } } });

console.log('-- o que a OS usa --');
ok('1. "Etiqueta de tamanho" cai no quadro Etiqueta; cordão não tem quadro',
   api._aviTipoDoMaterial('003 · Etiqueta de tamanho') === 'Etiqueta' && api._aviTipoDoMaterial('Cordão') === '' && api._aviTipoDoMaterial('Botão') === 'Botão');
const nec = api._aviNecessidadeOS(os('0001'));
ok('2. etiqueta por tamanho, na cor Preto: M 100, G 50 (1 por peça × produtos)',
   nec.length === 2 && nec.find(n => n.tam === 'M').qtd === 100 && nec.find(n => n.tam === 'G').qtd === 50 && nec.every(n => n.cor === 'Preto'), nec);
const nec2 = api._aviNecessidadeOS(os('0002', { aviamentos: [{ material: 'bo', qtdPorPeca: 3 }] }));
ok('3. botão, 3 por peça: 450 no total, sem tamanho', nec2.length === 1 && nec2[0].qtd === 450 && nec2[0].tam === '', nec2);

console.log('-- reservado, como o tecido --');
STATE.ordens = [os('0010'), os('0011', { st: 'cortando' }), os('0012', { st: 'transito-ida' }), os('0013', { st: 'cancelado' })];
let d = api._aviDasOS();
ok('4. não iniciada, cortando e em trânsito de ida reservam; cancelada não',
   [...new Set(d.reservas.map(r => r.osNumero))].join() === '0010,0011,0012' && !d.baixas.length, d);

console.log('-- baixa na costura --');
STATE.ordens = [os('0020', costurada('São Carlos', '2026-09-26')), os('0021', costurada('Descalvado', '2026-09-25'))];
d = api._aviDasOS();
const b20 = d.baixas.filter(b => b.osNumero === '0020');
ok('5. costura de São Carlos marcada: sai do estoque de São Carlos, na data da marcação',
   b20.length === 2 && b20.every(b => b.unidade === 'sc' && b.data === '2026-09-26' && b.tipo === 'saida' && b.auto), b20);
ok('6. costura de Descalvado: sai de Descalvado', d.baixas.filter(b => b.osNumero === '0021').every(b => b.unidade === 'desc'));
ok('7. costurada não reserva mais', !d.reservas.length, d.reservas);

console.log('-- o passado fica de fora --');
STATE.ordens = [os('0030', costurada('Descalvado', '2026-09-20')), os('0031', { st: 'finalizado' }),
  os('0032', { etapas, progresso: { etapasCheck: { 'Costura CM.LISA | Descalvado': true }, etapasSeq: { 'Costura CM.LISA | Descalvado': 57 } } })];
d = api._aviDasOS();
ok('8. costurada antes de 25/09, finalizada sem costura marcada, e marcação antiga sem hora: nem reserva nem baixa',
   !d.baixas.length && !d.reservas.length, d);
STATE.ordens = [os('0040', { st: 'costurando-sc', statusOSEm: '2026-09-27T12:00:00Z' })];
d = api._aviDasOS();
ok('9. status Costurando | São Carlos carimbado à mão também baixa, de São Carlos',
   d.baixas.length === 2 && d.baixas.every(b => b.unidade === 'sc' && b.data === '2026-09-27'), d);
STATE.ordens = [os('0050', { etapas, progresso: { etapasCheck: {
  'Costura CM.LISA | Descalvado': true, 'Costura CM.LISA | São Carlos': true }, etapasSeq: {
  'Costura CM.LISA | Descalvado': ts('2026-09-25'), 'Costura CM.LISA | São Carlos': ts('2026-09-26') } } })];
d = api._aviDasOS();
ok('10. as duas costuras marcadas: vale a marcada por último (a correção)', d.baixas.every(b => b.unidade === 'sc'), d);

console.log('-- a tela --');
ok('11. o saldo da tela e o limite da OE contam as baixas das OS',
   /const mov = _aviMovTodos\(\);/.test(src) && /const mov = _aviMovTodos\(\)\.filter\(m => m\.id !== semId\);/.test(src));
ok('12. o quadro do reservado aparece na tela, contra as duas unidades',
   /Reservado para OS ainda não costuradas/.test(src) && /\$\{reservaHtml\}/.test(src));
ok('13. a baixa da OS não tem "apagar" (sai desmarcando a costura)', /m\.auto\s*\?\s*'<span class="muted"[^']*>da OS<\/span>'/.test(src));
ok('14. a descrição da página não diz mais que tudo é manual', !/Os lançamentos são manuais/.test(html));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
