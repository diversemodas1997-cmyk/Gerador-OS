/* Rode com:  node testes/compra-oc-automatica.js

   A OC AUTOMÁTICA DA FALTA DE MATÉRIA-PRIMA (03/11/2026, Jones: "o programa
   deve preencher automaticamente as ordens de compra de matéria-prima, com base
   no aviso de falta calculado por meio das OS criadas... e aguardar até que o
   volume atinja o ideal para pedido de compra").

   O que este teste guarda:

     - a falta é somada por tecido e cor, e o tecido sem ideal NÃO gera OC;
     - abaixo do ideal o tecido ESPERA, e diz quanto falta;
     - atingido o ideal, nasce UMA OC por fornecedor, com cada cor e o quilo;
     - a segunda passada não pede de novo (a OC guarda em `cobre` o que pediu);
     - OC cancelada devolve a falta para "a pedir";
     - OC feita à mão, aberta, abate o mesmo tecido e cor.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pega(nome) {
  let ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) ini = src.indexOf('\nasync function ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome); process.exit(1); }
  return src.slice(ini, fim(ini));
}
function fim(ini) {
  const m = /\r?\n\}/g;
  m.lastIndex = ini;
  const r = m.exec(src);
  return r.index + r[0].length;
}
function pegaConst(nome) {
  const i = src.search(new RegExp('^const ' + nome + ' = ', 'm'));
  if (i < 0) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return src.slice(i, src.indexOf(';', i) + 1);
}

function monta(M) {
  return new Function('M', `
    var STATE = M.STATE;
    var n = 0;
    var uid = () => 'oc' + (n++);
    var toast = () => {};
    var saveState = async () => true;
    var podeGravar = () => true;
    var _contaPodeRegistrar = () => true;
    var agendarAutoSaveOC = () => {};
    var compraAplicarFaltas = () => false;
    function compraFaltasAbertas() { return M.faltas; }
    ${pega('_normNome')}
    ${pega('_ocLista')}
    ${pega('_ocNumeroNovo')}
    ${pegaConst('MP_OC_VIVA')}
    ${pega('_mpChave')}
    ${pega('_mpTecidoPorNome')}
    ${pega('_mpIdealDe')}
    ${pega('_mpCobertoAuto')}
    ${pega('_mpEmOCManual')}
    ${pega('compraDemandaMP')}
    let _mpGerando = false;
    ${pega('compraGerarOCsAutomaticas')}
    return { compraDemandaMP, compraGerarOCsAutomaticas };
  `)(M);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

(async () => {
  const M = {
    STATE: {
      meta: {}, compraOCs: [],
      tecidos: [
        { id: 't1', nome: 'Moletom', fornecedorId: 'f1', pesoBobina: 25 },
        { id: 't2', nome: 'Malha Algodão', fornecedorId: 'f1' },
        { id: 't3', nome: 'Piquet', fornecedorId: 'f2' }
      ]
    },
    faltas: [
      { osId: 'a', osNumero: '0700', tecidos: [{ tecidoNome: 'Moletom', corNome: 'Preto', falta: 120 }, { tecidoNome: 'Malha Algodão', corNome: 'Branco', falta: 40 }] },
      { osId: 'b', osNumero: '0701', tecidos: [{ tecidoNome: 'Moletom', corNome: 'Preto', falta: 100 }, { tecidoNome: 'Moletom', corNome: 'Cinza', falta: 30 }] },
      { osId: 'c', osNumero: '0702', tecidos: [{ tecidoNome: 'Piquet', corNome: 'Azul', falta: 90 }] }
    ]
  };
  const f = monta(M);
  const grupo = nome => f.compraDemandaMP().find(g => g.tecidoNome === nome);

  let g = grupo('Moletom');
  ok('1. a falta soma por tecido e cor', g.falta === 250 && g.cores.find(c => c.corNome === 'Preto').falta === 220, g);
  ok('2. sem ideal, o tecido fica em "defina o ideal"', g.situacao === 'sem-ideal', g.situacao);
  ok('3. sem ideal, nenhuma OC nasce', (await f.compraGerarOCsAutomaticas()) === 0 && !M.STATE.compraOCs.length);

  M.STATE.tecidos[0].pedidoIdealKg = 300;
  g = grupo('Moletom');
  ok('4. abaixo do ideal, espera', g.situacao === 'aguardando' && Math.round(g.pct) === 83, g);
  ok('5. esperando, nenhuma OC nasce', (await f.compraGerarOCsAutomaticas()) === 0);

  M.STATE.meta.compraIdealPadraoKg = 80;   // Piquet (90) passa; Malha (40) não
  M.STATE.tecidos[0].pedidoIdealKg = 250;  // Moletom (250) passa
  ok('6. o ideal padrão vale para o tecido sem ideal próprio', grupo('Piquet').ideal.padrao && grupo('Piquet').situacao === 'pronto');
  const nOC = await f.compraGerarOCsAutomaticas();
  const ocs = M.STATE.compraOCs;
  ok('7. uma OC por fornecedor', nOC === 2 && ocs.length === 2, ocs.map(o => o.fornecedorId));
  const ocF1 = ocs.find(o => o.fornecedorId === 'f1');
  ok('8. a OC leva cada cor do tecido que atingiu, e não o que espera',
    ocF1 && ocF1.itens.length === 2 && ocF1.itens.every(i => i.tecidoNome === 'Moletom'), ocF1 && ocF1.itens);
  ok('9. bobinas pelo peso da bobina, arredondadas para cima',
    ocF1.itens.find(i => i.corNome === 'Preto').bobinas === 9, ocF1.itens);
  ok('10. números sem repetir', new Set(ocs.map(o => o.numero)).size === 2 && M.STATE.meta.ocCounter === 2);
  ok('11. já pedido sai do "a pedir"', grupo('Moletom').situacao === 'pedido' && grupo('Moletom').pendente === 0);
  ok('12. a segunda passada não pede de novo', (await f.compraGerarOCsAutomaticas()) === 0 && ocs.length === 2);

  ocF1.status = 'recebida';
  ok('13. recebida (sem a entrada lançada) continua pedida', grupo('Moletom').situacao === 'pedido');
  ocF1.status = 'cancelada';
  ok('14. cancelada devolve a falta', grupo('Moletom').pendente === 250);

  M.STATE.compraOCs.push({ id: 'm', numero: 'OC-0009', status: 'enviada', itens: [{ tecidoNome: 'Moletom', corNome: 'Preto', kg: 200 }] });
  g = grupo('Moletom');
  ok('15. OC à mão, enviada, abate o mesmo tecido e cor', g.emOC === 200 && g.pendente === 50 && g.situacao === 'aguardando', g);

  console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
