/* Rode com:  node testes/tons-mais-menos.js

   AS LINHAS DE TOM DA FOLHA DE OS, agora sem caixa de marcar.

   Junior, 31/08/2026: "a folha de os, inclua opcional de adicionar mais tons,
   como um botao de mais... Usuario pode adicionar ate 7 tons e pode retirar
   todos os tons, exceto o tom 1." E logo depois: "retire as caixas de checklist
   das linhas de tons, pois o que determina essas linhas sao as camadas
   adicionadas na fase corpo 1."

   As duas frases mudam a MESMA coisa: quem escreve `progresso.totalTamanhoTons`.
   Antes eram N caixas independentes, cada uma capaz de deixar o estado furado
   (Tom 3 marcado sem o 2), e por isso `togglarTotalTamanhoTom` carregava uma
   lista de pre-requisitos escrita a mao, tom por tom. Agora sao duas portas — o
   + e o − — que so sabem escrever um PREFIXO CONTIGUO 1..N.

   O que este teste protege:

     · o Tom 1 existe sempre, mesmo em OS que nunca marcou tom nenhum;
     · o + para em 7 e o − para em 1, sem estourar nem sumir com o Tom 1;
     · retirar um tom apaga o RASTRO dele (o V do total e as camadas por tom em
       TODAS as fases) — sem isso o proximo lancamento na fase Corpo 1
       ressuscitaria a linha que acabou de ser recolhida, porque e dali que
       `recalcularDeCamadasPorTom` reescreve os tons;
     · um estado furado herdado do modelo antigo e curado na primeira mexida;
     · sem permissao de editar a folha, nada muda.

   O teste recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(nome) {
  const i = src.indexOf(nome);
  if (i < 0) { console.error('nao achei ' + nome + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + nome); process.exit(1); }
  return src.slice(i, j + 2);
}
const teto = (src.match(/const MAX_TONS = \d+;/) || [])[0];
if (!teto) { console.error('nao achei o MAX_TONS no app.js'); process.exit(1); }

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const printOsAtual = null;
  const document = { querySelector: () => null };
  const saveState = async () => { ctx.salvou++; };
  const exigirEdicaoFolha = () => ctx.podeEditar;
  const propagarVolumesExpedicaoOS = async () => 0;
  const _expSugestaoVolumes = () => '';
  const toast = (msg) => { ctx.avisos.push(msg); };
  const renderExpedicaoPlano = () => {};
  const renderPrintSheet = () => {};
  ${teto}
  ${corta('function tonsEfetivos')}
  ${corta('function nLinhasTomOS')}
  ${corta('async function _definirLinhasTomOS')}
  ${corta('async function adicionarLinhaTomOS')}
  ${corta('async function removerLinhaTomOS')}
  return { MAX_TONS, tonsEfetivos, nLinhasTomOS, adicionarLinhaTomOS, removerLinhaTomOS };
`)(ctx);

const ctxDe = (progresso = {}, podeEditar = true) => {
  const ctx = {
    podeEditar, salvou: 0, avisos: [],
    STATE: { ordens: [{ id: 'os1', os: '0500', progresso }] }
  };
  return { ctx, api: monta(ctx), os: () => ctx.STATE.ordens[0] };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};
const marcados = os => Object.keys(os.progresso.totalTamanhoTons || {}).sort().join(',');

(async () => {
  console.log('-- o Tom 1 e o piso --');
  {
    const t = ctxDe({});
    ok('1. OS sem tom nenhum marcado ja mostra UMA linha (o Tom 1 implicito)',
       t.api.nLinhasTomOS(t.os()) === 1, t.api.nLinhasTomOS(t.os()));
    await t.api.removerLinhaTomOS('os1');
    ok('2. o − nao tira o Tom 1 — e avisa por que',
       t.api.nLinhasTomOS(t.os()) === 1 && /não pode ser retirado/.test(t.ctx.avisos.join('')),
       t.ctx.avisos.join('|'));
    ok('3. e uma recusa nao grava nada', t.ctx.salvou === 0, t.ctx.salvou);
  }

  console.log('-- o + abre linha, ate o teto --');
  {
    const t = ctxDe({});
    await t.api.adicionarLinhaTomOS('os1');
    ok('4. o primeiro + leva a DUAS linhas (Tom 1 editavel + Tom 2 balanceador)',
       t.api.nLinhasTomOS(t.os()) === 2 && marcados(t.os()) === '1,2', marcados(t.os()));
    for (let i = 0; i < 10; i++) await t.api.adicionarLinhaTomOS('os1');
    ok('5. o + para no teto de ' + t.api.MAX_TONS,
       t.api.nLinhasTomOS(t.os()) === t.api.MAX_TONS, t.api.nLinhasTomOS(t.os()));
    ok('6. e o que ficou gravado e um prefixo contiguo 1..7',
       marcados(t.os()) === '1,2,3,4,5,6,7', marcados(t.os()));
    ok('7. o + recusado avisou o teto, e nao gravou de novo',
       t.ctx.avisos.length === 5 && t.ctx.salvou === 6,
       'avisos=' + t.ctx.avisos.length + ' gravacoes=' + t.ctx.salvou);
  }

  console.log('-- o − apaga o rastro do tom retirado --');
  {
    const t = ctxDe({
      totalTamanhoTons: { 1: true, 2: true, 3: true },
      totalTamanhoTomValor: { 1: 24, 2: 12, 3: 8 },
      // Duas fases com camadas por tom: a principal (Corpo 1) e a da ribana.
      enfestosTons: { 1: { 1: '12', 2: '6', 3: '4' }, 3: { 1: '2', 3: '1' } }
    });
    await t.api.removerLinhaTomOS('os1');
    ok('8. retirar deixa duas linhas', t.api.nLinhasTomOS(t.os()) === 2 && marcados(t.os()) === '1,2',
       marcados(t.os()));
    ok('9. o V do tom retirado sai do total por tamanho',
       t.os().progresso.totalTamanhoTomValor[3] === undefined,
       JSON.stringify(t.os().progresso.totalTamanhoTomValor));
    ok('10. e as camadas dele saem de TODAS as fases, nao so da principal',
       t.os().progresso.enfestosTons[1][3] === undefined
       && t.os().progresso.enfestosTons[3][3] === undefined,
       JSON.stringify(t.os().progresso.enfestosTons));
    ok('11. o que ficou nao foi tocado',
       t.os().progresso.totalTamanhoTomValor[1] === 24
       && t.os().progresso.enfestosTons[1][2] === '6'
       && t.os().progresso.enfestosTons[3][1] === '2',
       JSON.stringify(t.os().progresso));
  }

  console.log('-- estado furado do modelo antigo --');
  {
    // Isto so existia porque as caixas eram independentes: alguem marcou o Tom 3
    // com o 2 desmarcado. tonsEfetivos sempre leu isso como [1]; agora a primeira
    // mexida tambem CURA o dado.
    const t = ctxDe({ totalTamanhoTons: { 1: true, 3: true } });
    ok('12. um Tom 3 solto nunca contou como tonalidade',
       t.api.nLinhasTomOS(t.os()) === 1, t.api.nLinhasTomOS(t.os()));
    await t.api.adicionarLinhaTomOS('os1');
    ok('13. e o primeiro + limpa o furo em vez de somar em cima dele',
       marcados(t.os()) === '1,2', marcados(t.os()));
  }

  console.log('-- permissao --');
  {
    const t = ctxDe({ totalTamanhoTons: { 1: true, 2: true } }, false);
    await t.api.adicionarLinhaTomOS('os1');
    await t.api.removerLinhaTomOS('os1');
    ok('14. sem permissao de editar a folha, o numero de tons nao muda',
       marcados(t.os()) === '1,2' && t.ctx.salvou === 0, marcados(t.os()) + ' / ' + t.ctx.salvou);
  }

  console.log(falhas === 0 ? '\nTODOS OS TESTES PASSARAM' : `\n${falhas} FALHA(S)`);
  process.exit(falhas === 0 ? 0 : 1);
})();
