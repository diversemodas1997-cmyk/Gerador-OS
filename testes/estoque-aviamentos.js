/* ESTOQUE DE AVIAMENTOS (24/09/2026, Junior: "Insira um Estoque de aviamentos
   na barra lateral abaixo de Estoque de tecidos ... volume de entrada, volume
   de saída, volume residual (estocado), volume corrente e volume total ...
   quadro para cada tipo: Fio, linha, etiqueta, botão, viés ... cor e peso").

   A conta e recortada do app.js e rodada com lancamentos de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const tiposLinha = (src.match(/const AVIAMENTO_TIPOS = [^\n]*/) || [''])[0];
const unidadeDe = (src.match(/const _aviUnidadeDe = [^\n]*/) || [''])[0];
const origemDe = (src.match(/const _aviOrigemDe = [^\n]*/) || [''])[0];
const destinoDe = (src.match(/const _aviDestinoDe = [^\n]*/) || [''])[0];

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const api = new Function(`
  ${tiposLinha}
  ${unidadeDe}
  ${origemDe}
  ${destinoDe}
  ${pegaFuncao('_aviPernasDaExpedicao')}
  ${pegaFuncao('_normNome')}
  ${pegaFuncao('calcularEstoqueAviamentos')}
  return { calcularEstoqueAviamentos, AVIAMENTO_TIPOS };
`)();

console.log('-- os cinco tipos --');
ok('1. um quadro para cada tipo pedido',
   api.AVIAMENTO_TIPOS.join(',') === 'Fio,Linha,Etiqueta,Botão,Viés', api.AVIAMENTO_TIPOS);

/* Linha preta: 10 kg em agosto (antes do periodo), no periodo de setembro
   entram 5 e saem 3, e em outubro (depois do periodo) saem mais 4. Hoje e
   15/10, entao o corrente ja viu a saida de outubro. */
const mov = [
  { tipo: 'entrada', item: 'Linha', cor: 'Preto', kg: 10, data: '2026-08-20' },
  { tipo: 'entrada', item: 'Linha', cor: 'preto', kg: 5, data: '2026-09-05' },
  { tipo: 'saida', item: 'linha', cor: 'Preto', kg: 3, data: '2026-09-10' },
  { tipo: 'saida', item: 'Linha', cor: 'Preto', kg: 4, data: '2026-10-02' },
  { tipo: 'entrada', item: 'Botão', cor: 'Bege', kg: 1.25, data: '2026-09-30' },
  // Lancamento com data depois de hoje: ainda nao aconteceu.
  { tipo: 'entrada', item: 'Botão', cor: 'Bege', kg: 9, data: '2026-12-01' }
];
const r = api.calcularEstoqueAviamentos(mov, '2026-09-01', '2026-09-30', '2026-10-15');
const linha = r.find(x => x.item === 'Linha');
const botao = r.find(x => x.item === 'Botão');

console.log('-- os cinco volumes --');
ok('2. mesma cor com outra caixa cai na MESMA linha', r.filter(x => x.item === 'Linha').length === 1, r);
ok('3. ENTRADA e SAIDA contam so o periodo', linha.entrada === 5 && linha.saida === 3, linha);
ok('4. RESIDUAL = o de antes + entrada - saida (10 + 5 - 3 = 12)', linha.residual === 12, linha);
ok('5. TOTAL = o que ja havia + o que entrou (10 + 5 = 15)', linha.total === 15, linha);
ok('6. CORRENTE e o de hoje, e ve a saida de outubro (12 - 4 = 8)', linha.corrente === 8, linha);
ok('7. lancamento com data futura nao entra no corrente', botao.corrente === 1.25, botao);
ok('8. com o periodo terminando hoje, residual e corrente se encontram',
   (() => { const x = api.calcularEstoqueAviamentos(mov, '2026-09-01', '2026-10-15', '2026-10-15').find(y => y.item === 'Linha');
     return x.residual === x.corrente && x.corrente === 8; })());

console.log('-- as duas unidades --');
/* 24/09/2026: "separe por Unidade Descalvado e Unidade Sao Carlos". A mesma
   linha preta tem 6 kg em Descalvado (lancamento sem unidade = Descalvado) e
   4 kg em Sao Carlos, e 1 kg sai de Sao Carlos. */
const mu = [
  { tipo: 'entrada', item: 'Linha', cor: 'Preto', kg: 6, data: '2026-09-02' },
  { tipo: 'entrada', unidade: 'sc', item: 'Linha', cor: 'Preto', kg: 4, data: '2026-09-03' },
  { tipo: 'saida', unidade: 'sc', item: 'Linha', cor: 'Preto', kg: 1, data: '2026-09-04' }
];
const ud = api.calcularEstoqueAviamentos(mu, '2026-09-01', '2026-09-30', '2026-09-30', 'desc')[0];
const us = api.calcularEstoqueAviamentos(mu, '2026-09-01', '2026-09-30', '2026-09-30', 'sc')[0];
ok('9a. Descalvado conta so o dele (sem unidade = Descalvado): 6 kg', ud.corrente === 6 && ud.saida === 0, ud);
ok('9b. Sao Carlos conta so o dele: 4 - 1 = 3 kg', us.entrada === 4 && us.saida === 1 && us.corrente === 3, us);
ok('9c. sem unidade pedida, as duas somam (6 + 3 = 9)',
   api.calcularEstoqueAviamentos(mu, '2026-09-01', '2026-09-30', '2026-09-30')[0].corrente === 9);
ok('9d. o lancamento grava a unidade escolhida', /unidade: v\('ma-unidade'\) === 'sc' \? 'sc' : 'desc'/.test(src));
ok('9e. a tela tem as abas das duas unidades',
   /rotulo: 'Unidade Descalvado'/.test(src) && /rotulo: 'Unidade São Carlos'/.test(src) && /_aviTrocarUnidade\('\$\{u\.k\}'\)/.test(src));

console.log('-- a migracao pela Ordem de Expedicao --');
/* 24/09/2026: "os volumes de entrada da Unidade Descalvado sao expedidos para
   Unidade Sao Carlos ... atraves da Ordem de expedicao". Descalvado tem 10 kg
   de linha preta. Em 20/09 alocam-se 4 kg na IDA da carga de 25/09. */
const base = [{ tipo: 'entrada', unidade: 'desc', item: 'Linha', cor: 'Preto', kg: 10, data: '2026-09-01' }];
const ida = { tipo: 'expedicao', janelaId: 'j1', data: '2026-09-25', perna: 'ida', item: 'Linha', cor: 'Preto', kg: 4, dataSaida: '2026-09-20' };
const em = (hoje, un, mv, efetiva) => api.calcularEstoqueAviamentos(mv || base.concat(ida), '2026-09-01', hoje, hoje, un, efetiva)
  .find(x => x.item === 'Linha') || { corrente: 0, entrada: 0, saida: 0 };
ok('14. antes de alocar, Descalvado tem os 10 kg', em('2026-09-19', 'desc').corrente === 10);
ok('15. alocado (22/09): sai de Descalvado na hora (10 - 4 = 6)', em('2026-09-22', 'desc').corrente === 6, em('2026-09-22', 'desc'));
ok('16. e ainda NAO chegou a Sao Carlos: esta em transito', em('2026-09-22', 'sc').corrente === 0, em('2026-09-22', 'sc'));
ok('17. na data da carga (25/09) entra em Sao Carlos: 4 kg', em('2026-09-25', 'sc').corrente === 4, em('2026-09-25', 'sc'));
ok('18. Descalvado segue com 6, e a soma das duas volta aos 10', em('2026-09-25', 'desc').corrente === 6 && em('2026-09-25').corrente === 10);
ok('19. em Sao Carlos a expedicao conta como ENTRADA no periodo', em('2026-09-30', 'sc').entrada === 4);
ok('20. em Descalvado ela conta como SAIDA no periodo', em('2026-09-30', 'desc').saida === 4);
// Ocorrencia remarcada de 25/09 para 28/09: a chegada vai junto.
const remarcada = m => (m.janelaId === 'j1' && m.data === '2026-09-25') ? '2026-09-28' : m.data;
ok('21. carga remarcada (25 -> 28/09): em 26/09 ainda esta em transito',
   em('2026-09-26', 'sc', null, remarcada).corrente === 0 && em('2026-09-28', 'sc', null, remarcada).corrente === 4);
// A VOLTA traz de Sao Carlos para Descalvado.
const volta = { tipo: 'expedicao', janelaId: 'j1', data: '2026-09-29', perna: 'volta', item: 'Linha', cor: 'Preto', kg: 1, dataSaida: '2026-09-29' };
const mvv = base.concat(ida, volta);
ok('22. a VOLTA tira de Sao Carlos e poe em Descalvado (SC 4 - 1 = 3, DESC 6 + 1 = 7)',
   em('2026-09-30', 'sc', mvv).corrente === 3 && em('2026-09-30', 'desc', mvv).corrente === 7);
ok('23. a OE tem o botao de alocar aviamento em cada perna, e a folha o imprime',
   /onclick="abrirModalExpAviamento\(/.test(src) && /\$\{linhas\}\$\{aviPrint\}/.test(src));
ok('24. nao embarca mais do que a unidade tem', /if \(kg > tem \+ 0\.0005\)/.test(src));

console.log('-- entrada por quantidade de unidade --');
/* 24/09/2026: "deve haver entrada por quantidade de unidade". Botao se conta:
   entram 500 un em Descalvado, saem 120, e 100 un embarcam na ida de 25/09. */
const mun = [
  { tipo: 'entrada', unidade: 'desc', item: 'Botão', cor: 'Bege', qtd: 500, data: '2026-09-02' },
  { tipo: 'saida', unidade: 'desc', item: 'Botão', cor: 'Bege', qtd: 120, data: '2026-09-03' },
  { tipo: 'expedicao', janelaId: 'j1', data: '2026-09-25', perna: 'ida', item: 'Botão', cor: 'Bege', qtd: 100, dataSaida: '2026-09-20' }
];
const bd = api.calcularEstoqueAviamentos(mun, '2026-09-01', '2026-09-30', '2026-09-30', 'desc')[0];
const bs = api.calcularEstoqueAviamentos(mun, '2026-09-01', '2026-09-30', '2026-09-30', 'sc')[0];
ok('25. a linha lancada so em unidades conta em unidades, e nao em kg', bd.temUn && !bd.temKg && bd.corrente === 0, bd);
ok('26. os cinco volumes em unidades: entrada 500, saida 120 + 100 = 220, residual e corrente 280',
   bd.un.entrada === 500 && bd.un.saida === 220 && bd.un.residual === 280 && bd.un.corrente === 280 && bd.un.total === 500, bd.un);
ok('27. as unidades viajam na OE: Sao Carlos recebe as 100 un', bs.un.entrada === 100 && bs.un.corrente === 100, bs.un);
// Peso e unidade juntos no mesmo lancamento: as duas contas andam.
const ambos = api.calcularEstoqueAviamentos([{ tipo: 'entrada', unidade: 'desc', item: 'Etiqueta', cor: '', kg: 1.5, qtd: 1000, data: '2026-09-02' }],
  '2026-09-01', '2026-09-30', '2026-09-30', 'desc')[0];
ok('28. peso e quantidade no mesmo lancamento contam os dois', ambos.corrente === 1.5 && ambos.un.corrente === 1000 && ambos.temKg && ambos.temUn, ambos);
ok('29. a janela tem o campo de quantidade e aceita so ela',
   /id="ma-qtd"/.test(src) && /if \(!\(kg > 0\) && !\(qtd > 0\)\) return toast\('Informe a quantidade \(un\) ou o peso \(kg\)'/.test(src));
ok('30. a OE nao embarca mais unidades do que a unidade tem', /if \(qtd > temUn\)/.test(src));

console.log('-- a tela --');
ok('9. o item de menu fica logo abaixo do Estoque de tecidos',
   /data-page="estoque"[^>]*>Estoque de tecidos<\/a>\s*<a class="nav-btn" href="#estoque-aviamentos" data-page="estoque-aviamentos"[^>]*>Estoque de aviamentos<\/a>/.test(html));
ok('10. a pagina e o modal existem',
   /<section class="page hidden" data-page="estoque-aviamentos">/.test(html) && /id="modal-aviamento"/.test(html));
ok('11. a rota desenha a tela', /if \(page === 'estoque-aviamentos'\) renderEstoqueAviamentos\(\);/.test(src));
ok('12. a chave aviamentosMov e sincronizada e carregada (as duas listas de chaves)',
   (src.match(/'compraOCs','aviamentosMov'/g) || []).length === 2, (src.match(/'compraOCs','aviamentosMov'/g) || []).length);
ok('13. o lancamento guarda cor e peso', /cor: v\('ma-cor'\)\.trim\(\),\s*kg: Math\.round/.test(src));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
