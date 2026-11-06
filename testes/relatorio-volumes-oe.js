/* Rode com:  node testes/relatorio-volumes-oe.js

   RELATÓRIO DE VOLUMES POR OE (06/11/2026): no planejamento de expedição, por
   dia, semana, mês e ano, o número da OE, as OS dela e o total de volumes de
   cada OS naquela expedição.
     · o período Anual vai de 1º/jan a 31/dez e navega de ano em ano;
     · a mesma OS na ida e na volta da OE soma numa linha só;
     · expedição cancelada e expedição sem OS ficam de fora;
     · OS em ordem numérica dentro da OE; OE em ordem de data.

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(de, ate) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + de + ' no app.js'); process.exit(1); }
  return src.slice(i, j + ate.length);
}

let falhas = 0;
const ok = (cond, msg) => { if (cond) console.log('ok   ' + msg); else { falhas++; console.log('FALHA ' + msg); } };

// Itens por perna, como resumoPernaExpedicao devolveria. `volumes` são os
// PACOTES e não podem aparecer no relatório; `pecas` são as unidades completas.
const PERNAS = {
  'M|2026-11-04|ida': [{ os: { id: 'o10', sku: ['CM.LISA-PRE'] }, osNumero: '0610', volumes: 5, pecas: 500 }, { os: { id: 'o2', sku: [], base: 'BM.TRI' }, osNumero: '0602', volumes: 3, pecas: 300 }],
  'M|2026-11-04|volta': [{ os: { id: 'o10', sku: ['CM.LISA-PRE'] }, osNumero: '0610', volumes: 2, pecas: 200 }],
  'T|2026-11-04|ida': [],
  'T|2026-11-04|volta': [],
  'X|2026-11-03|ida': [{ os: { id: 'o9', sku: ['CM.LISA-AZU', 'CM.LISA-BRA'] }, osNumero: '0609', volumes: 4, pecas: 400 }]
};
const STATE = {
  expedicaoCargas: [
    { id: 'a', janelaId: 'M', data: '2026-11-04', oe: 85 },
    { id: 'b', janelaId: 'X', data: '2026-11-03', oe: 84 }
  ]
};
const f = new Function('STATE', 'PERNAS', `
  const _EXP_DIAS = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
  const _EXP_DIAS_CURTO = ['DOM','SEG','TER','QUA','QUI','SEX','SÁB'];
  const _EXP_MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const formatDate = iso => iso.split('-').reverse().join('/');
  const esc = s => String(s == null ? '' : s);
  const resumoPernaExpedicao = (oc, perna) => ({ itens: PERNAS[oc.janela.id + '|' + oc.dataOrig + '|' + perna] || [] });
  const skusDaOS = o => o.sku;
  const _skuBaseDaOS = o => o.base || '';
  ${corta('function _expIso', '\n}')}
  ${corta('function _expData', '\n}')}
  ${corta('function _expAddDias', '\n}')}
  ${corta('function _expRange', '\n}')}
  ${corta('function _expNavegar', '\n}')}
  ${corta('function _expLabelPeriodo', '\n}')}
  ${corta('const _oeFmt', ';\n')}
  ${corta('function _oeNumeroDe', '\n}')}
  ${corta('function _expRelatorioVolumes', '\n}')}
  ${corta('function _expRelatorioLinhas', '\n}')}
  return { _expRange, _expNavegar, _expLabelPeriodo, _expRelatorioVolumes, _expRelatorioLinhas };
`)(STATE, PERNAS);

// Período anual
const r = f._expRange('ano', '2026-11-06');
ok(r.ini === '2026-01-01' && r.fim === '2026-12-31', 'Anual cobre 01/01 a 31/12');
ok(f._expNavegar('ano', '2026-11-06', -1) === '2025-01-01', 'Anual volta um ano');
ok(f._expNavegar('ano', '2026-11-06', 1) === '2027-01-01', 'Anual avança um ano');
ok(f._expLabelPeriodo('ano', '2026-11-06') === 'Ano de 2026', 'rótulo do período anual');
ok(f._expRange('mes', '2026-11-06').fim === '2026-11-30', 'Mensal continua igual');

// Relatório
const oc = (id, nome, dataOrig, extra) => Object.assign({ janela: { id, nome }, dataOrig, data: dataOrig }, extra || {});
const rel = f._expRelatorioVolumes([
  oc('M', 'Manhã', '2026-11-04'),
  oc('T', 'Tarde', '2026-11-04'),                          // sem OS: sai
  oc('X', 'Extra', '2026-11-03'),
  oc('M', 'Manhã', '2026-11-04', { cancelada: true })       // cancelada: sai
]);
ok(rel.oes.length === 2, 'só as OE com OS e não canceladas (' + rel.oes.length + ')');
ok(rel.oes[0].oe === 84 && rel.oes[1].oe === 85, 'OE em ordem de data');
const oe85 = rel.oes[1];
ok(oe85.os.map(o => o.os).join(',') === '0602,0610', 'OS em ordem numérica dentro da OE');
ok(oe85.os.find(o => o.os === '0610').volumes === 700, 'volume em unidades completas, não pacotes; ida + volta somam (500 + 200)');
ok(oe85.volumes === 1000, 'total da OE 85 = 1000 un.');
ok(rel.totalVolumes === 1400 && rel.totalOS === 3, 'total do período: 1400 un., 3 OS');
ok(oe85.os.find(o => o.os === '0610').sku === 'CM.LISA-PRE', 'SKU completo da OS');
ok(oe85.os.find(o => o.os === '0602').sku === 'BM.TRI', 'sem sigla de cor, cai na linha do SKU');
ok(rel.oes[0].os[0].sku === 'CM.LISA-AZU / CM.LISA-BRA', 'OS com duas cores mostra os dois SKUs');
const html = f._expRelatorioLinhas(rel);
ok(html.includes('OE 0085') && html.includes('OS 0610') && html.includes('CM.LISA-PRE') && html.includes('Total da OE'), 'tabela mostra nº da OE, OS, SKU e total');
ok(/OS 0610<\/td>\s*<td class="rel-sku">CM\.LISA-PRE/.test(html), 'SKU logo na frente da OS');

console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
