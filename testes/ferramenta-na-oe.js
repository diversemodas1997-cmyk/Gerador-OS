/* A FERRAMENTA NA ORDEM DE EXPEDICAO (27/10/2026, Junior: "insira no
   planejamento de OE botao + Ferramenta, ao lado do botao + Aviamento").

   Alocar tira do em estoque da origem na hora; na data da carga a ferramenta
   entra na linha de mesmo tipo da outra unidade (lida das OE, nao gravada);
   tirar da OE devolve. As funcoes sao recortadas do app.js. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function pegaFuncao(nome, async) {
  const ini = src.indexOf('\n' + (async ? 'async ' : '') + 'function ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const linha = re => (src.match(re) || [''])[0];
const bloco = (ini, fim) => src.slice(src.indexOf(ini), src.indexOf(fim, src.indexOf(ini)) + fim.length);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const campos = {};
const documento = { getElementById: id => campos[id] || (campos[id] = { innerHTML: '', textContent: '', value: '' }) };
const STATE = {
  ferramentasCad: [
    { id: 't1', nome: 'Tesoura de corte', desc: '10"', unidade: 'desc', emUso: 2, emEstoque: 5 },
    { id: 'r1', nome: 'Régua', unidade: 'desc', emUso: 0, emEstoque: 1 }
  ],
  ferramentasMov: []
};
const ctxHoje = { hoje: '2026-10-27' };
const api = new Function('document', 'STATE', 'ctxHoje', `
  const esc = s => String(s == null ? '' : s);
  const formatDate = d => d;
  ${linha(/const AVIAMENTO_UNIDADES = \[[\s\S]*?\];/)}
  ${linha(/const _aviUnidadeDe = [^\n]*/)}
  ${linha(/const _aviOrigemDe = [^\n]*/)}
  ${linha(/const _aviDestinoDe = [^\n]*/)}
  ${pegaFuncao('_normNome')}
  ${bloco('const ESTOQUE_ITENS = {', '\n};')}
  ${bloco('const _estItemMedida = ', "\n")}
  ${bloco('const _estArred = ', "Math.round(Number(v) || 0);")}
  const _expDataEfetivaCarga = c => c.data;
  const _aviHoje = () => ctxHoje.hoje;
  let _n = 0; const uid = () => 'id' + (++_n);
  const salvos = []; const saveState = async k => { salvos.push(k); };
  const closeModal = () => {}; const openModal = () => {};
  const toasts = []; const toast = (t, k) => toasts.push(k + ':' + t);
  const exigirEdicao = () => true;
  const confirm = () => true;
  const renderExpedicaoPlano = () => {}; const _oeRerenderFolhaSeAberta = () => {};
  let _estItensCtx = null;
  ${['_ferExpMovs', '_ferDaPerna', '_ferDataChegada', '_ferChegou', '_estItemChegadas', '_estItemEstoque', '_ferEmTransito', '_ferChegadasHist', '_estItensRegistrar', 'abrirModalExpFerramenta'].map(f => pegaFuncao(f)).join('\n')}
  ${pegaFuncao('salvarFerramentaExp', true)}
  ${pegaFuncao('excluirFerramentaExp', true)}
  return { abrirModalExpFerramenta, salvarFerramentaExp, excluirFerramentaExp, _ferDaPerna, _estItemEstoque, _ferEmTransito, _ferChegadasHist,
           ctx: () => _estItensCtx, toasts, salvos };
`)(documento, STATE, ctxHoje);

const preenche = (item, qtd, vol, obs) => {
  documento.getElementById('mef-item').value = item;
  documento.getElementById('mef-qtd').value = String(qtd);
  documento.getElementById('mef-vol').value = String(vol || '');
  documento.getElementById('mef-obs').value = obs || '';
};
const est = id => api._estItemEstoque('ferramentas', STATE.ferramentasCad.find(x => x.id === id));

(async () => {
console.log('-- alocar --');
api.abrirModalExpFerramenta('j1', '2026-10-29', 'ida');
const modal = campos['modal-estoque-item-fields'].innerHTML;
ok('1. o modal lista as ferramentas da origem (Descalvado) com o em estoque', /Tesoura de corte · 10" — 5 em estoque/.test(modal) && /Régua — 1 em estoque/.test(modal), modal.slice(0, 400));
preenche('t1', 6);
await api.salvarFerramentaExp(api.ctx());
ok('2. nao embarca mais do que ha em estoque', api.toasts.pop() === 'err:Só há 5 de Tesoura de corte em estoque nesta unidade' && !STATE.ferramentasMov.length);
preenche('t1', 2, 1, 'mesa 2');
await api.salvarFerramentaExp(api.ctx());
const m = STATE.ferramentasMov[0];
ok('3. sai do em estoque da origem na hora (5 -> 3)', est('t1') === 3, est('t1'));
ok('4. a saida fica no historico com a OE', m && m.tipo === 'saida' && m.motivo === 'expedicao' && m.dEstoque === -2 && m.janelaId === 'j1' && m.oeData === '2026-10-29' && m.perna === 'ida' && m.volumes === 1, m);
const dest = STATE.ferramentasCad.find(x => x.unidade === 'sc');
ok('5. a linha de Sao Carlos nasce zerada com o mesmo tipo e descricao', dest && dest.nome === 'Tesoura de corte' && dest.desc === '10"' && dest.emEstoque === 0 && m.destItemId === dest.id, dest);
ok('6. a perna da OE mostra a ferramenta', api._ferDaPerna('j1', '2026-10-29', 'ida').length === 1 && !api._ferDaPerna('j1', '2026-10-29', 'volta').length);
ok('7. antes da carga: em transito, e Sao Carlos ainda sem nada', api._ferEmTransito('ferramentas', STATE.ferramentasCad[0]) === 2 && est(dest.id) === 0);
ok('8. grava cadastro e historico', api.salvos.includes('ferramentasCad') && api.salvos.includes('ferramentasMov'));

console.log('-- a carga --');
ctxHoje.hoje = '2026-10-29';
ok('9. na data da carga entra em Sao Carlos (0 -> 2) sem gravar nada', est(dest.id) === 2 && dest.emEstoque === 0);
ok('10. e deixa de estar em transito', api._ferEmTransito('ferramentas', STATE.ferramentasCad[0]) === 0);
const ch = api._ferChegadasHist('ferramentas');
ok('11. a chegada aparece no historico de Sao Carlos', ch.length === 1 && ch[0].unidade === 'sc' && ch[0].tipo === 'entrada' && ch[0].dEstoque === 2 && ch[0].data === '2026-10-29', ch);

console.log('-- corrigir e tirar --');
api.abrirModalExpFerramenta('j1', '2026-10-29', 'ida', m.id);
preenche('t1', 3, 1, 'mesa 2');
await api.salvarFerramentaExp(api.ctx());
ok('12. corrigir 2 -> 3 tira so a diferenca da origem (3 -> 2) e Sao Carlos recebe 3', est('t1') === 2 && est(dest.id) === 3 && STATE.ferramentasMov.length === 1, [est('t1'), est(dest.id)]);
await api.excluirFerramentaExp(m.id);
ok('13. tirar da OE devolve a origem (2 -> 5) e Sao Carlos volta a 0', est('t1') === 5 && est(dest.id) === 0 && !STATE.ferramentasMov.length, [est('t1'), est(dest.id)]);

console.log('-- volta --');
ctxHoje.hoje = '2026-10-27';
dest.emEstoque = 1;
api.abrirModalExpFerramenta('j1', '2026-10-29', 'volta');
ok('14. na volta a origem e Sao Carlos', /Unidade São Carlos/.test(campos['modal-estoque-item-fields'].innerHTML) && /Tesoura de corte · 10" — 1 em estoque/.test(campos['modal-estoque-item-fields'].innerHTML));
preenche(dest.id, 1);
await api.salvarFerramentaExp(api.ctx());
ok('15. e a chegada cai na linha que ja existe em Descalvado, sem criar outra', STATE.ferramentasCad.length === 3 && STATE.ferramentasMov[0].destItemId === 't1');

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
})();
