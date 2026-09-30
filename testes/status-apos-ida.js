/* Rode com:  node testes/status-apos-ida.js

   DEPOIS DA IDA, O CHECKLIST DE DESCALVADO NÃO PUXA A OS DE VOLTA (30/09/2026).

   A OS 0565 foi alocada numa OE às 08:16 (a caixa "Expedição Desc X São Carlos"
   marcada pela alocação) e o Ensaque foi marcado às 08:22 — e o status voltou
   para "Ensacado | Descalvado" com a carga já na OE. O que o teste guarda:
     · ida marcada + Ensaque marcado DEPOIS = Em trânsito | IDA;
     · o mesmo com o corte ou o enfesto marcados depois;
     · o que vem depois da ida (recebido em SC, costura de lá) continua vencendo;
     · sem a ida marcada, o Ensaque vale como sempre;
     · OS antiga, sem etapasSeq, continua lendo pela `ordem`.

   Recorta do app.js as funções reais. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaArr = (nome) => recorte(nome, '\n];', nome) + '\n];';

const _statusOS = new Function('STATE', `
  ${cortaArr('const STATUS_OS')}
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE', 'constantes das unidades')}
  ${corta('function osEtapaMarcada')}
  ${corta('function _marcasDoStatus')}
  ${corta('function _statusDoChecklistOS')}
  ${corta('function _ultimaMarcacaoChecklist')}
  ${corta('function _statusOS')}
  return _statusOS;
`)({});

let falhas = 0;
const ok = (nome, obtido, esperado) => {
  const cond = obtido === esperado;
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + obtido + ' (esperado ' + esperado + ')'));
  if (!cond) falhas++;
};

const IDA = 'Expedição Desc X São Carlos';
const os = (seq, extra) => {
  const etapasCheck = {};
  Object.keys(seq).forEach(n => { etapasCheck[n] = true; });
  return Object.assign({ id: 'x', os: '0565',
    etapas: ['Corte', 'Ensaque', 'Costura CM.LISA | Descalvado', IDA, 'Recebido em São Carlos',
             'Costura CM.LISA | São Carlos', 'Expedição São Carlos X Desc.', 'Recebido em Descalvado', 'Estoque'],
    progresso: { etapasCheck, etapasSeq: seq } }, extra || {});
};

ok('0565: alocada 08:16, Ensaque 08:22 -> Em trânsito | IDA',
   _statusOS(os({ 'Corte': 100, [IDA]: 200, 'Ensaque': 300 })), 'transito-ida');
ok('corte marcado depois da ida também não puxa de volta',
   _statusOS(os({ [IDA]: 200, 'Corte': 300 })), 'transito-ida');
ok('enfesto marcado depois da ida também não',
   _statusOS(Object.assign(os({ [IDA]: 200 }), { progresso: { etapasCheck: { [IDA]: true }, etapasSeq: { [IDA]: 200 },
     enfestosCheck: { 1: true }, enfestosSeq: { 1: 300 } } })), 'transito-ida');
ok('recebido em São Carlos depois da ida vence (Ensacado | SC)',
   _statusOS(os({ [IDA]: 200, 'Ensaque': 300, 'Recebido em São Carlos': 400 })), 'ensacado-sc');
ok('costura de São Carlos depois da ida vence',
   _statusOS(os({ [IDA]: 200, 'Recebido em São Carlos': 300, 'Costura CM.LISA | São Carlos': 400 })), 'costurando-sc');
ok('volta marcada depois da ida vence',
   _statusOS(os({ [IDA]: 200, 'Expedição São Carlos X Desc.': 400 })), 'transito-volta');
ok('sem a ida marcada, o Ensaque vale como sempre',
   _statusOS(os({ 'Corte': 100, 'Ensaque': 300 })), 'ensacado');
ok('carimbo à mão mais novo que a última marca continua valendo',
   _statusOS(os({ [IDA]: 200, 'Ensaque': 300 }, { statusOS: 'parado', statusOSEm: new Date(400).toISOString() })), 'parado');

// OS antiga: marcas sem etapasSeq. Continua lendo pela `ordem` -> Ensacado.
const antiga = os({});
antiga.progresso.etapasCheck = { 'Corte': true, 'Ensaque': true, [IDA]: true };
ok('OS antiga sem carimbo de hora continua lendo pela ordem', _statusOS(antiga), 'ensacado');

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
