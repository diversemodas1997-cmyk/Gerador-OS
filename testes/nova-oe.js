/* Rode com:  node testes/nova-oe.js

   BOTÃO NOVA OE (11/11/2026): a OE nova sai com o próximo número e o dia de
   hoje. Recorta a função real do app.js e confere que:
     · havendo expedição de hoje ainda sem OE, é nela que a OS vai;
     · sem nenhuma, cria uma expedição extra para hoje, em horário livre;
     · o título da janela mostra o próximo número (maior OE + 1);
     · clicar de novo reaproveita a expedição criada (não cria outra). */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = src.indexOf('async function novaOE()');
const fim = src.indexOf('\n}\n', ini);
if (ini < 0 || fim < 0) { console.error('nao achei novaOE no app.js'); process.exit(1); }
const corpo = src.slice(ini, fim + 3);

let falhas = 0;
const ok = (cond, msg) => { console.log((cond ? 'ok   ' : 'FALHA ') + msg); if (!cond) falhas++; };

function monta(janelas, cargas, agora) {
  const STATE = { expedicaoJanelas: janelas, expedicaoCargas: cargas };
  const t = { textContent: '' };
  const abertos = [];
  let n = 0;
  const HOJE = '2026-11-11';
  const f = new Function('STATE', 't', 'abertos', 'HOJE', 'AGORA', `
    let n = 0;
    const exigirEdicao = () => true;
    const _expHoje = () => HOJE;
    const uid = () => 'j' + (++n);
    const saveState = async () => {};
    const formatDate = d => d.split('-').reverse().join('/');
    const _oeFmt = x => String(x).padStart(4, '0');
    const document = { getElementById: () => t };
    const abrirModalExpCarga = (j, d, p) => abertos.push([j, d, p]);
    const _oeNumeroDe = (j, d) => { const c = STATE.expedicaoCargas.find(c => c.janelaId === j && c.data === d && c.oe > 0); return c ? c.oe : 0; };
    const _expHorasIdaDoDia = d => STATE.expedicaoJanelas.filter(j => j.data === d || !j.data).map(j => j.horaIda);
    const ocorrenciasExpedicao = (a, b) => STATE.expedicaoJanelas.filter(j => !j.data || j.data === a)
      .map(j => ({ janela: j, dataOrig: a, data: a, horaIda: j.horaIda, cancelada: false }));
    const Date = function () { return AGORA; };   // o relógio do teste
    ${corpo}
    return novaOE;`);
  return { STATE, t, abertos, novaOE: f(STATE, t, abertos, HOJE, agora) };
}
const agora = h => ({ getHours: () => h[0], getMinutes: () => h[1] });

(async () => {
  // Matinal de hoje já é a OE 94; a Tarde de hoje ainda não tem OS.
  let m = monta([{ id: 'M', nome: 'Matinal', horaIda: '09:30' }, { id: 'T', nome: 'Tarde', horaIda: '16:00' }],
    [{ janelaId: 'M', data: '2026-11-11', oe: 94 }], agora([10, 7]));
  await m.novaOE();
  ok(m.abertos[0] && m.abertos[0][0] === 'T' && m.abertos[0][1] === '2026-11-11', 'usa a expedição de hoje ainda sem OE (Tarde)');
  ok(m.STATE.expedicaoJanelas.length === 2, 'não cria expedição quando já há uma livre hoje');
  ok(m.t.textContent.startsWith('Nova OE 0095 · hoje, 11/11/2026'), 'título com o próximo número e o dia: ' + m.t.textContent);

  // Todas as de hoje já têm OE: cria uma extra para hoje.
  m = monta([{ id: 'M', nome: 'Matinal', horaIda: '09:30' }],
    [{ janelaId: 'M', data: '2026-11-11', oe: 94 }], agora([10, 7]));
  await m.novaOE();
  const nova = m.STATE.expedicaoJanelas[1];
  ok(nova && nova.data === '2026-11-11' && nova.tipo === 'data', 'cria expedição extra só para hoje');
  ok(nova && nova.horaIda === '10:15' && nova.horaVolta === '11:15', 'no próximo quarto de hora livre, volta 1 h depois: ' + (nova && nova.horaIda));
  ok(m.abertos[0][0] === nova.id, 'abre a alocação nela');
  // Cancelou (não salvou OS) e clicou de novo: reaproveita a mesma.
  await m.novaOE();
  ok(m.STATE.expedicaoJanelas.length === 2 && m.abertos[1][0] === nova.id, 'clicar de novo reaproveita a expedição criada');

  // Horário ocupado: anda até achar um livre.
  m = monta([{ id: 'M', nome: 'Matinal', horaIda: '09:30' }, { id: 'X', nome: 'Extra', horaIda: '10:15', data: '2026-11-11' }],
    [{ janelaId: 'M', data: '2026-11-11', oe: 94 }, { janelaId: 'X', data: '2026-11-11', oe: 95 }], agora([10, 7]));
  await m.novaOE();
  ok(m.STATE.expedicaoJanelas[2].horaIda === '10:16', 'hora ocupada: pega a próxima livre');
  ok(m.t.textContent.startsWith('Nova OE 0096'), 'próximo número depois da 95');

  console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
