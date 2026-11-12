/* Rode com:  node testes/ordem-preparo-op.js

   A ORDEM DE PREPARO DE MATÉRIA-PRIMA — OP (12/11/2026, Junior: "adicione em
   estoque de tecidos quadro para gerenciar as tarefas da fase Preparação de
   matéria-prima [...] Cada lote avaliado deve conter número, assim como uma OS.
   A folha de Ordem de preparo OP deve conter quantidade de bobinas totais, tipo
   de tecido e cor. Os lotes devem ser formados por bobinas de tecido e ribana.
   [...] botão na coluna de ação com nome Avaliar amostras [...] As cores
   avaliadas devem gerar relatório na OP"). Protege:
     · a lista tem a coluna Ações com "Avaliar amostras";
     · o número nasce como o da OS: maior + 1, sem repetir;
     · a folha traz bobinas totais, tecido, cor e ribana;
     · o relatório do Colorímetro vai para a folha da OP;
     · lote sem bobina, sem tecido ou sem cor não é salvo.
   Carrega o bloco da OP do app.js de verdade. */
const fs = require('fs');
const path = require('path');
const app = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const ini = app.indexOf('/*     ORDEM DE PREPARO DE MATÉRIA-PRIMA — OP (12/11/2026)   */');
const fim = app.indexOf('function abrirMovEstoque(tipo) {');
if (ini < 0 || fim < ini) { console.log('FALHA não achei o bloco da OP no app.js'); process.exit(1); }
const bloco = app.slice(ini, fim);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra).slice(0, 400)));
  if (!cond) falhas++;
};
const plano = s => String(s).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');

const els = {};
const classes = () => { const c = new Set(); return { toggle: (k, v) => (v ? c.add(k) : c.delete(k)), remove: k => c.delete(k), contains: k => c.has(k) }; };
const doc = { getElementById: id => (els[id] = els[id] || { id, value: '', innerHTML: '', textContent: '', classList: classes(), parentNode: null }) };
const STATE = { tecidos: [{ nome: 'Malha PV' }, { nome: 'Ribana PV' }], cores: [{ nome: 'Marinho' }], ordensPreparo: [] };
const salvos = [], toasts = [], avaliados = [];
let pagina = '', respostaPrompt = '';
const env = {
  document: doc, STATE, window: { Colorimetro: {} },
  Colorimetro: { avaliarOS: id => avaliados.push(id) },
  esc: s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'),
  formatDate: iso => { const [a, m, d] = String(iso).split('-'); return `${d}/${m}/${a}`; },
  uid: (() => { let i = 0; return () => 'p' + (++i); })(),
  exigirEstoqueTecidos: () => true, saveState: async k => salvos.push(k), toast: (m, t) => toasts.push([m, t]),
  openModal: () => {}, closeModal: () => {}, confirm: () => true, prompt: () => respostaPrompt,
  _carimboCanceladoHtml: () => '<div class="sheet-carimbo-area"><div class="sheet-carimbo-cancelado"><span>CANCELADO</span></div></div>', _carimbosDiagonaisOE: () => {}, goto: p => { pagina = p; }, _obsQuemSou: () => 'nathaly'
};
const nomes = Object.keys(env);
const api = new Function(...nomes, bloco + '\nreturn { salvarTextoRelatorioOP, mudarStatusOP,renderOrdensPreparo, abrirModalOP, salvarOP, excluirOP, avaliarAmostrasOP, salvarRelatorioCorOP, verOP, renderPrintPreparo };')(...nomes.map(k => env[k]));

(async () => {
  const preencher = (v) => { api.abrirModalOP(); Object.entries(v).forEach(([k, x]) => { doc.getElementById(k).value = x; }); };
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': 'Marinho', 'opp-bt': '4', 'opp-ribana': 'Ribana PV', 'opp-br': '2', 'opp-data': '2026-11-12', 'opp-obs': '' });
  await api.salvarOP();
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': 'Marinho', 'opp-bt': '3', 'opp-ribana': '', 'opp-br': '0', 'opp-data': '2026-11-13', 'opp-obs': 'NF 9' });
  await api.salvarOP();
  const [a, b] = STATE.ordensPreparo;
  ok('1. duas OP salvas, número 1 e 2', a && b && a.numero === 1 && b.numero === 2 && salvos.filter(k => k === 'ordensPreparo').length === 2, STATE.ordensPreparo);
  ok('2. o lote tem tecido, cor, bobinas de tecido e de ribana', a.tecidoNome === 'Malha PV' && a.corNome === 'Marinho' && a.bobinasTecido === 4 && a.ribanaNome === 'Ribana PV' && a.bobinasRibana === 2, a);
  // número não se repete
  await api.excluirOP(a.id);
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': 'Marinho', 'opp-bt': '1', 'opp-ribana': '', 'opp-br': '0', 'opp-data': '', 'opp-obs': '' });
  await api.salvarOP();
  ok('3. excluir a OP 1 e criar outra: nasce a OP 3 (maior + 1)', STATE.ordensPreparo.map(x => x.numero).join() === '2,3', STATE.ordensPreparo.map(x => x.numero));
  // validação
  const antes = STATE.ordensPreparo.length;
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': 'Marinho', 'opp-bt': '0', 'opp-ribana': '', 'opp-br': '0' }); await api.salvarOP();
  preencher({ 'opp-tecido': '', 'opp-cor': 'Marinho', 'opp-bt': '2' }); await api.salvarOP();
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': '', 'opp-bt': '2' }); await api.salvarOP();
  preencher({ 'opp-tecido': 'Malha PV', 'opp-cor': 'Marinho', 'opp-bt': '2', 'opp-ribana': '', 'opp-br': '3' }); await api.salvarOP();
  ok('4. sem bobina, sem tecido, sem cor ou ribana sem nome: não salva', STATE.ordensPreparo.length === antes, STATE.ordensPreparo.length);
  // lista
  api.renderOrdensPreparo();
  const lista = els['op-prep-painel'].innerHTML;
  ok('5. lista: coluna Ações com "Avaliar amostras" e as OP numeradas', /<th class="col-actions">Ações<\/th>/.test(lista) && (lista.match(/>Avaliar amostras<\/button>/g) || []).length === 2 && /OP 0002/.test(lista) && /OP 0003/.test(lista), plano(lista).slice(0, 400));
  ok('6. lista: total de bobinas (3 + 0 = 3) e "A avaliar"', /Total bobinas/.test(lista) && /A avaliar/.test(lista), plano(lista));
  // avaliar → Colorímetro
  api.avaliarAmostrasOP(b.id);
  ok('7. Avaliar amostras abre o Colorímetro com a OP', avaliados[0] === b.id, avaliados);
  // folha
  const c = STATE.ordensPreparo.find(x => x.numero === 3);
  c.bobinasRibana = 2; c.ribanaNome = 'Ribana PV';
  api.verOP(c.id);
  ok('8. "folha" abre a página da folha da OP', pagina === 'print-preparo', pagina);
  api.renderPrintPreparo();
  let f = plano(els['print-sheet-prep'].innerHTML);
  ok('9. folha: título, número, bobinas totais, tecido, cor e ribana', /ORDEM DE PREPARO DE MATÉRIA-PRIMA OP/.test(f) && /OP 0003/.test(f) && /3 Bobinas totais/.test(f)
    && /Tipo de tecido Malha PV/.test(f) && /Cor Marinho/.test(f) && /Ribana Ribana PV · Marinho/.test(f), f);
  ok('10. folha sem relatório: pede Avaliar amostras', /Ainda não avaliado/.test(f), f);
  // relatório do Colorímetro
  const r = await api.salvarRelatorioCorOP(c.id, 'Tecidos compatíveis em cor\nTecido Tom 1 (ref 1: 1 - ribana 1,2)');
  f = plano(els['print-sheet-prep'].innerHTML);
  ok('11. relatório gravado na OP e mostrado na folha', r === true && c.relatorios.length === 1 && c.relatorios[0].por === 'nathaly'
    && /Tecido Tom 1 \(ref 1: 1 - ribana 1,2\)/.test(f) && /Avaliado em/.test(f), f);
  api.renderOrdensPreparo();
  ok('12. lista: a OP avaliada sai como "Avaliada"', /Avaliada/.test(els['op-prep-painel'].innerHTML), '');
  ok('13. relatório vazio ou de OP que não existe: não grava', (await api.salvarRelatorioCorOP(c.id, '  ')) === false && (await api.salvarRelatorioCorOP('nada', 'x')) === false, '');
  // STATUS (12/11/2026, "insira status na OP: em preparo, pronta, cancelada")
  ok('14. OP nova nasce "em preparo" e a lista tem a coluna Status', c.status === 'em-preparo' && /<th>Status<\/th>/.test(els['op-prep-painel'].innerHTML)
    && /Em preparo/.test(els['op-prep-painel'].innerHTML) && /mudarStatusOP\('[^']+','pronta'\)/.test(els['op-prep-painel'].innerHTML), plano(els['op-prep-painel'].innerHTML));
  await api.mudarStatusOP(c.id, 'pronta');
  f = plano(els['print-sheet-prep'].innerHTML);
  ok('15. marcar pronta: status, quem/quando, folha diz Pronta', c.status === 'pronta' && c.statusPor === 'nathaly' && c.historicoStatus.length === 1 && /Status Pronta/.test(f), f);
  respostaPrompt = null; await api.mudarStatusOP(c.id, 'cancelada');
  ok('16. cancelar e desistir no motivo: não cancela', c.status === 'pronta', c.status);
  respostaPrompt = 'tecido devolvido'; await api.mudarStatusOP(c.id, 'cancelada');
  const L = els['op-prep-painel'].innerHTML; f = plano(els['print-sheet-prep'].innerHTML);
  ok('17. cancelada: motivo guardado, folha com carimbo CANCELADO e motivo', c.status === 'cancelada' && c.motivoCancelamento === 'tecido devolvido'
    && els['print-sheet-prep'].classList.contains('folha-cancelada') && /sheet-carimbo-cancelado/.test(els['print-sheet-prep'].innerHTML) && /motivo do cancelamento: tecido devolvido/.test(f), f);
  const linhaC = L.split('<tr').find(t => t.includes('OP 0003')) || '';
  ok('18. lista: cancelada sem Avaliar amostras/editar, com reabrir', /Cancelada/.test(linhaC) && !/Avaliar amostras/.test(linhaC) && !/abrirModalOP/.test(linhaC) && /'em-preparo'\)[^>]*>reabrir/.test(linhaC), linhaC);
  const n0 = avaliados.length; api.avaliarAmostrasOP(c.id); api.abrirModalOP(c.id);
  ok('19. cancelada: Avaliar amostras e editar recusados', avaliados.length === n0 && toasts.slice(-2).every(t => /cancelada/.test(t[0])), toasts.slice(-2));
  await api.mudarStatusOP(c.id, 'em-preparo');
  ok('20. reabrir: volta para em preparo, sem carimbo e sem motivo', c.status === 'em-preparo' && c.motivoCancelamento === '' && !els['print-sheet-prep'].classList.contains('folha-cancelada')
    && !/CANCELADO/.test(els['print-sheet-prep'].innerHTML) && c.historicoStatus.length === 3, c);
  delete b.status; api.renderOrdensPreparo();
  ok('21. OP antiga sem status conta como em preparo', (els['op-prep-painel'].innerHTML.split('<tr').find(t => t.includes('OP 0002')) || '').includes('Em preparo'), '');
  // RELATÓRIO EDITÁVEL NA FOLHA (12/11/2026, "o relatório de tonalidade na folha
  // de OP deve ser aberto para usuário editar o texto, assim como o campo de
  // observação das OS")
  api.renderPrintPreparo();
  let H = els['print-sheet-prep'].innerHTML;
  ok('22. folha: o relatório é um campo que grava ao sair (salvarTextoRelatorioOP)', /<textarea[^>]*op-prep-rel-campo[^>]*onchange="salvarTextoRelatorioOP\('p3', this.value\)"[^>]*>Tecidos compatíveis em cor\nTecido Tom 1/.test(H), H.slice(-900));
  const emAntes = c.relatorios[0].em;
  ok('23. editar o texto corrige o último relatório, mantém a data e marca quem editou',
    (await api.salvarTextoRelatorioOP(c.id, 'Tom 1 (ref 1: 1 - ribana 1,2)\nbobina 2 separada')) === true && c.relatorios.length === 1
    && c.relatorios[0].texto === 'Tom 1 (ref 1: 1 - ribana 1,2)\nbobina 2 separada' && c.relatorios[0].em === emAntes && c.relatorios[0].editadoPor === 'nathaly'
    && /texto editado em/.test(plano(els['print-sheet-prep'].innerHTML)) && /bobina 2 separada/.test(els['print-sheet-prep'].innerHTML), c.relatorios);
  ok('24. mesmo texto: não grava de novo', (await api.salvarTextoRelatorioOP(c.id, c.relatorios[0].texto)) === false, '');
  const sem = STATE.ordensPreparo.find(x => x.numero === 2); api.verOP(sem.id); api.renderPrintPreparo();
  ok('25. OP sem avaliação: campo vazio na folha', /<textarea[^>]*op-prep-rel-campo[^>]*><\/textarea>/.test(els['print-sheet-prep'].innerHTML), els['print-sheet-prep'].innerHTML.slice(-700));
  ok('26. escrever sem ter avaliado vira relatório "escrito à mão"; em branco não grava',
    (await api.salvarTextoRelatorioOP(sem.id, '   ')) === false && (await api.salvarTextoRelatorioOP(sem.id, 'Tom único')) === true
    && sem.relatorios.length === 1 && sem.relatorios[0].manual === true && /Escrito à mão em/.test(plano(els['print-sheet-prep'].innerHTML)), sem.relatorios);
  respostaPrompt = 'x'; await api.mudarStatusOP(sem.id, 'cancelada');
  H = els['print-sheet-prep'].innerHTML;
  ok('27. OP cancelada: relatório só texto, e editar é recusado', !/<textarea/.test(H) && /Tom único/.test(H) && (await api.salvarTextoRelatorioOP(sem.id, 'outro')) === false && sem.relatorios[0].texto === 'Tom único', H.slice(-600));
  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
