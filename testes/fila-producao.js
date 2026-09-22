/* Rode com:  node testes/fila-producao.js

   A FILA DE PRODUCAO (22/09/2026, Junior: "insira a capacidade do usuario
   determinar a ordem das OS com status nao iniciado em 1a, 2a, 3a, 4a, 5a, etc.
   Essa capacidade deve ser concedida para o usuario pelo admin").

   O que este teste guarda:

     - so OS NAO INICIADA entra na fila. A que ja comecou nao tem ordem a
       definir, e enfileirar o que ja saiu da mesa encheria a tela de decisoes
       vencidas;
     - sem ninguem ter ordenado, a fila e a ordem natural da casa: numero menor
       (OS mais antiga) primeiro;
     - quem tem lugar marcado vem antes de quem nao tem, e uma OS nova entra no
       FIM sem empurrar ninguem;
     - escrever a posicao poe a OS naquele lugar e ACOMODA as outras; numero
       fora da faixa nao recusa - 0 vira a primeira, 99 vira a ultima;
     - a OS que sai do "nao iniciado" GUARDA O LUGAR (o id fica na lista), e so
       o id de OS apagada some de vez. Era isso que fazia uma OS devolvida para
       "nao iniciada" voltar para onde estava em vez de cair no fim;
     - sem a area concedida pelo admin, nada muda.

   Recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
function pegaAsync(nome) {
  const ini = src.indexOf('\nasync function ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao async ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
/* Anda por LINHA ate o fim da declaracao, e nao ate o primeiro ";" com quebra
   colada: o app.js e gravado ora com LF, ora com CRLF (o git converte no
   checkout), e um corte que procura ";\n" devolve VAZIO no dia em que o arquivo
   esta com CRLF — a constante some do motor e o teste morre dizendo que ela nao
   existe. O comentario do fim da linha tambem sai, que ele pode ter
   ponto-e-virgula no meio da frase. */
function pegaConst(nome) {
  const i = src.search(new RegExp('^const ' + nome + ' = ', 'm'));
  if (i < 0) { console.error('nao achei a constante ' + nome); process.exit(1); }
  const out = [];
  for (const l of src.slice(i).split(/\r?\n/)) {
    out.push(l);
    if (l.replace(/\/\/[^\r\n]*$/, '').trimEnd().endsWith(';')) return out.join('\n');
  }
  console.error('nao achei o fim da constante ' + nome);
  process.exit(1);
}

/* O status da OS entra DUBLADO (ele vem do checklist da folha, com teste
   proprio), e o mesmo vale para a gravacao e a permissao — o que se prova aqui
   e a ORDEM. `M.pode` e a area "Fila de producao" concedida pelo admin. */
function monta(M) {
  return new Function('M', `
    var STATE = M.STATE;
    var window = {};
    function _statusOS(o) { return M.status[o.id] || 'nao-iniciado'; }
    function exigirEdicao(acao) { M.pedidos.push(acao); return !!M.pode; }
    async function saveState(k) { M.salvou.push(k); }
    function desfazerNomearAcao() {}
    function renderListaOS() {}
    function toast() {}
    function temAcesso() { return !!M.pode; }
    function esc(s) { return String(s == null ? '' : s); }
    ${pegaConst('FILA_OS_ABERTA_CHAVE')}
    ${pegaFuncao('numeroOSordenacao')}
    ${pegaFuncao('_filaLista')}
    ${pegaFuncao('_osNaoIniciadas')}
    ${pegaFuncao('podeMexerFilaOS')}
    ${pegaFuncao('filaDeProducao')}
    ${pegaFuncao('_filaPosicoes')}
    ${pegaConst('_filaOrdinal')}
    ${pegaAsync('_filaGravar')}
    ${pegaAsync('moverNaFila')}
    ${pegaAsync('definirPosicaoFila')}
    return { filaDeProducao, _filaPosicoes, moverNaFila, definirPosicaoFila,
             _filaOrdinal, podeMexerFilaOS, fila: () => STATE.meta.filaOS };
  `)(M);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ----------------------
   Cinco OS: quatro nao iniciadas (0101, 0102, 0103, 0104) e uma cortando. */
const mundo = (pode) => ({
  pode: pode !== false,
  pedidos: [], salvou: [],
  STATE: {
    meta: {},
    ordens: [
      { id: 'a', os: '0103' }, { id: 'b', os: '0101' },
      { id: 'c', os: '0104' }, { id: 'd', os: '0102' },
      { id: 'e', os: '0099' }
    ]
  },
  status: { e: 'cortando' }
});
const numeros = api => api.filaDeProducao().map(o => o.os).join(' ');

(async () => {
  /* ---------- 1. quem entra e em que ordem ---------- */
  let M = mundo(); let api = monta(M);
  ok('1. so as nao iniciadas entram (a 0099 esta cortando e fica de fora)',
     numeros(api) === '0101 0102 0103 0104', numeros(api));
  ok('2. sem ninguem ter ordenado, a ordem e a natural: a OS mais antiga primeiro',
     api.filaDeProducao()[0].os === '0101', numeros(api));
  ok('3. as posicoes sao 1, 2, 3, 4', [...api._filaPosicoes().values()].join(',') === '1,2,3,4',
     [...api._filaPosicoes().entries()]);
  ok('4. o ordinal sai em portugues (1a, 2a)',
     api._filaOrdinal(1) === '1ª' && api._filaOrdinal(12) === '12ª', api._filaOrdinal(1));

  /* ---------- 2. mover uma casa ---------- */
  await api.moverNaFila('c', -1);       // a 0104 sobe de 4a para 3a
  ok('5. a seta sobe uma casa', numeros(api) === '0101 0102 0104 0103', numeros(api));
  await api.moverNaFila('c', -1);
  await api.moverNaFila('c', -1);
  ok('6. tres subidas poem a 0104 em primeiro', numeros(api) === '0104 0101 0102 0103', numeros(api));
  await api.moverNaFila('c', -1);
  ok('7. subir a primeira nao faz nada (nao da a volta)', numeros(api) === '0104 0101 0102 0103', numeros(api));
  ok('8. gravou em meta, e so a chave meta', M.salvou.every(k => k === 'meta') && M.salvou.length > 0, M.salvou);

  /* ---------- 3. escrever a posicao ---------- */
  M = mundo(); api = monta(M);
  await api.definirPosicaoFila('c', 1);   // 0104 para a 1a
  ok('9. escrever "1" poe a OS em primeiro e acomoda as outras',
     numeros(api) === '0104 0101 0102 0103', numeros(api));
  await api.definirPosicaoFila('c', 99);
  ok('10. numero maior que a fila vira a ultima posicao',
     numeros(api) === '0101 0102 0103 0104', numeros(api));
  await api.definirPosicaoFila('c', 0);
  ok('11. zero vira a primeira', numeros(api) === '0104 0101 0102 0103', numeros(api));
  const antes = numeros(api);
  await api.definirPosicaoFila('c', '');
  ok('12. campo apagado nao mexe na fila', numeros(api) === antes, numeros(api));

  /* ---------- 4. a OS que comeca guarda o lugar ---------- */
  M = mundo(); api = monta(M);
  await api.definirPosicaoFila('c', 1);            // 0104 em primeiro
  M.status.c = 'enfestando';                       // a 0104 comecou
  ok('13. a OS que comecou sai da fila', numeros(api) === '0101 0102 0103', numeros(api));
  ok('14. mas o lugar dela fica guardado na lista', api.fila().includes('c'), api.fila());
  await api.moverNaFila('d', -1);                  // mexe na fila sem ela
  ok('15. e continua guardado depois de a fila ser mexida', api.fila().includes('c'), api.fila());
  M.status.c = 'nao-iniciado';
  ok('16. voltando para "nao iniciada", ela volta para o lugar de antes',
     api.filaDeProducao()[0].os === '0104', numeros(api));
  // OS apagada: o id some de vez na proxima gravacao.
  M.STATE.ordens = M.STATE.ordens.filter(o => o.id !== 'c');
  await api.moverNaFila('b', 1);
  ok('17. id de OS apagada some da lista', !api.fila().includes('c'), api.fila());

  /* ---------- 5. a permissao vem do admin ---------- */
  M = mundo(false); api = monta(M);
  const ordemAntes = numeros(api);
  await api.moverNaFila('c', -1);
  await api.definirPosicaoFila('c', 1);
  ok('18. sem a area concedida, a ordem nao muda', numeros(api) === ordemAntes, numeros(api));
  ok('19. e nada e gravado', M.salvou.length === 0, M.salvou);
  ok('20. a recusa passa pela area certa (exigirEdicao com a acao da fila)',
     M.pedidos.every(a => a.includes('fila de produção')) && M.pedidos.length === 2, M.pedidos);
  ok('21. quem nao tem a area continua LENDO a fila inteira',
     api.filaDeProducao().length === 4 && api.podeMexerFilaOS() === false, numeros(api));

  console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
