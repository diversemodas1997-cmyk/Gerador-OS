/* Rode com:  node testes/numero-os-servidor.js

   O NÚMERO DA OS VEM DO SERVIDOR (28/10/2026, Jones: "conceder capacidade de
   criação de OS para outro usuário além do admin, de forma que não aconteça
   conflito de dois usuários criarem ordens de serviço diferente com número
   igual"). Toda OS nova nasce no planejamento; a aprovação relê as OS do
   servidor, calcula o próximo número sobre todas e reserva um bloco no banco
   (reservar_numeros_os), que trava a linha e nunca entrega o mesmo número duas
   vezes. O que este teste protege:
     · duas aprovações ao mesmo tempo, com a MESMA lista na memória, saem com
       números diferentes;
     · a releitura do servidor conta a OS que outro computador acabou de gerar;
     · grade que conjuga ocupa dois números (a OS em cima, a conjugada embaixo);
     · "3 grades iguais" sai num bloco seguido;
     · servidor sem a função: o número sai da releitura, como antes;
     · planejamento aprovado em outro computador não vira OS de novo;
     · número do bloco já tomado: não gera nada;
     · quem não tem a área Aprovar planejamento não aprova.
   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n}', i) + 2);
}
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^\\n]*;$', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

/* O BANCO DE MENTIRA: a mesma conta da função SQL — o maior entre `minimo` e
   "último entregue + 1", e o último anda o tamanho do bloco. Cada chamada é
   atômica aqui, como a trava do banco a faz lá. */
const novoBanco = (semFuncao) => ({
  ultimo: 0, chamadas: [],
  async rpc(nome, args) {
    if (semFuncao) return { data: null, error: { code: 'PGRST202', message: 'Could not find the function' } };
    this.chamadas.push(Object.assign({ nome }, args));
    const inicio = Math.max(this.ultimo + 1, args.minimo || 1, 1);
    this.ultimo = inicio + args.qtd - 1;
    return { data: inicio, error: null };
  }
});

const monta = (ctx) => new Function('ctx', `
  let STATE = ctx.STATE;
  const supa = ctx.banco;
  const currentUser = { id: 'u1', email: ctx.login };
  let _cloudLoadErro = false;
  // A releitura do servidor: o que ctx.servidor tiver é o que a tela passa a ver.
  const cloudLoad = async () => { ctx.releu = (ctx.releu || 0) + 1; _cloudLoadErro = !!ctx.servidorFora; };
  const loadState = async () => { if (ctx.servidor) ctx.servidor(STATE); };
  const toast = (m, t) => ctx.toasts.push((t || '') + ':' + m);
  const confirm = () => true;
  const temAcesso = (k) => k === 'aprovar-plan' && !!ctx.aprova;
  const _recusarSomenteLeitura = () => false;
  let n = 0;
  const uid = () => 'os' + (++n) + '-' + ctx.nome;
  const _obsQuemSou = () => ctx.login;
  const _congelarDesenhoNaOS = () => {};
  const faltaDeTecidoParaOS = () => [];
  const _textoFaltaDeTecido = () => '';
  const _aplicarTonsPadrao = (o) => o;
  const saveState = async (k) => { ctx.gravou.push(k); };
  const atualizarCounterOS = async () => {};
  const aplicarBaixaEstoqueOS = async () => {};
  const aplicarRegraConjugadaSeAplicavel = async (o) => {
    if (!gradeConjugaAlguma(o.gradeId)) return null;
    const c = { id: 'c-' + o.id, os: numeroParaConjugada(o), conjugadaPaiId: o.id };
    STATE.ordens.push(c);
    return c;
  };
  const salvarPdfEtiquetasAuto = () => {};
  const dadosEtiquetaParaOS = () => ({});
  const atualizarContagensNav = () => {};
  const renderPlanejamentoProducao = () => {};
  const sanitizeForFilename = (s) => s;
  const _planResumo = () => '';
  let _salvandoOS = false;
  ${constante('NUMERO_OS_MAX')}
  ${constante('OS_SALTO_ESTRANHO')}
  ${recorte('async function _comTravaDeSalvar', 'a trava de salvar')}
  ${recorte('function _ehFuncaoAusente', 'a funcao ausente')}
  ${recorte('function formatarNumeroOS', 'o formato do numero')}
  ${recorte('function _numeroOSCanonico', 'o numero canonico')}
  ${recorte('function _numerosOSExistentes', 'os numeros existentes')}
  ${recorte('function _topoDaFilaOS', 'o topo da fila')}
  ${recorte('function proximoNumeroOS()', 'o proximo numero')}
  ${recorte('function gradeConjugaAlguma', 'a grade que conjuga')}
  ${recorte('function numeroOSLivre', 'o numero livre')}
  ${recorte('function proximoNumeroOSParaGrade', 'o numero da grade')}
  ${recorte('function numeroParaConjugada', 'o numero da conjugada')}
  ${recorte('function _planLista', 'a lista de planejamentos')}
  ${recorte('function _planVezes', 'as grades iguais')}
  ${recorte('function _planAguardando', 'o planejamento em espera')}
  ${recorte('function podeAprovarPlanejamento', 'quem aprova')}
  ${recorte('async function _reservarNumerosOS', 'a reserva no servidor')}
  ${recorte('async function aprovarPlanejamento', 'a aprovacao')}
  return { aprovarPlanejamento, STATE };
`)(ctx);

const novo = (estado, extra) => {
  const ctx = Object.assign({ nome: 'a', login: 'nathaly@diverse.local', aprova: true, toasts: [], gravou: [],
    banco: novoBanco(), STATE: estado }, extra || {});
  return { ctx, api: monta(ctx) };
};
const estado = (ordens, planos, grades) => ({ ordens: ordens || [], osPlanejadas: planos || [], grades: grades || [] });
const numeros = (st) => st.ordens.map(o => o.os).sort();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

(async () => {
  console.log('-- duas aprovacoes ao mesmo tempo nao repetem numero --');
  {
    const banco = novoBanco();
    const ordens = [{ id: 'x', os: '0700' }];
    // Dois computadores com a MESMA lista na memória (nenhum viu a OS do outro).
    const a = novo(estado(JSON.parse(JSON.stringify(ordens)), [{ id: 'p1', gradeId: 'g' }]), { banco, nome: 'a' });
    const b = novo(estado(JSON.parse(JSON.stringify(ordens)), [{ id: 'p2', gradeId: 'g' }]), { banco, nome: 'b', login: 'admin@diverse.local' });
    await Promise.all([a.api.aprovarPlanejamento('p1'), b.api.aprovarPlanejamento('p2')]);
    const na = a.ctx.STATE.ordens.find(o => o.planejamento).os, nb = b.ctx.STATE.ordens.find(o => o.planejamento).os;
    ok('1. as duas saem com numero', !!na && !!nb, [na, nb]);
    ok('2. e com numeros DIFERENTES (0701 e 0702)', na !== nb && [na, nb].sort().join() === '0701,0702', [na, nb]);
    ok('3. cada uma pediu ao banco a partir do topo que viu (0701)', banco.chamadas.every(c => c.minimo === 701 && c.qtd === 1), banco.chamadas);
    ok('4. e releu o servidor antes', a.ctx.releu === 1 && b.ctx.releu === 1);
    ok('5. o planejamento sai da lista ao virar OS', !a.ctx.STATE.osPlanejadas.length && a.ctx.gravou.includes('osPlanejadas'));
    ok('6. a OS guarda quem planejou e quem aprovou', a.ctx.STATE.ordens.find(o => o.planejamento).planejamento.aprovadoPor === 'nathaly@diverse.local');
  }

  console.log('\n-- a releitura conta a OS que outro computador acabou de gerar --');
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), {
      servidor: (st) => { if (!st.ordens.some(o => o.os === '0701')) st.ordens.push({ id: 'y', os: '0701' }); }
    });
    await t.api.aprovarPlanejamento('p1');
    ok('7. a tela via 0700, o servidor ja tinha 0701: sai 0702', t.ctx.STATE.ordens.some(o => o.planejamento && o.os === '0702'), numeros(t.ctx.STATE));
  }

  console.log('\n-- grade que conjuga ocupa dois numeros --');
  {
    const grades = [{ id: 'g1', conjugadaGradeId: 'g2' }, { id: 'g2' }];
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g1' }], grades));
    await t.api.aprovarPlanejamento('p1');
    ok('8. pede dois numeros ao banco', t.ctx.banco.chamadas[0].qtd === 2, t.ctx.banco.chamadas);
    ok('9. a OS fica em cima (0702) e a conjugada embaixo (0701)', numeros(t.ctx.STATE).join() === '0700,0701,0702'
       && t.ctx.STATE.ordens.find(o => o.planejamento).os === '0702', numeros(t.ctx.STATE));
  }

  console.log('\n-- tres grades iguais saem num bloco seguido --');
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g', repeticoes: 3 }]));
    await t.api.aprovarPlanejamento('p1');
    ok('10. um pedido de 3 numeros', t.ctx.banco.chamadas.length === 1 && t.ctx.banco.chamadas[0].qtd === 3, t.ctx.banco.chamadas);
    ok('11. 0701, 0702 e 0703', numeros(t.ctx.STATE).join() === '0700,0701,0702,0703', numeros(t.ctx.STATE));
    // Grade que conjuga x 2: quatro números, OS em cima de cada par.
    const g = [{ id: 'g1', conjugadaGradeId: 'g2' }, { id: 'g2' }];
    const u = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g1', repeticoes: 2 }], g));
    await u.api.aprovarPlanejamento('p1');
    ok('12. conjugada x2: 0702 e 0704 sao as OS, 0701 e 0703 as conjugadas',
       u.ctx.STATE.ordens.filter(o => o.planejamento).map(o => o.os).join() === '0702,0704'
       && numeros(u.ctx.STATE).join() === '0700,0701,0702,0703,0704', numeros(u.ctx.STATE));
  }

  console.log('\n-- o banco lembra o que ja entregou --');
  {
    const banco = novoBanco();
    banco.ultimo = 705;   // 0701..0705 entregues a aprovacoes que ainda nao chegaram
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), { banco });
    await t.api.aprovarPlanejamento('p1');
    ok('13. a lista diz 0701, o banco ja entregou ate 0705: sai 0706', t.ctx.STATE.ordens.some(o => o.os === '0706'), numeros(t.ctx.STATE));
  }

  console.log('\n-- servidor sem a funcao: o numero sai da releitura --');
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), { banco: novoBanco(true) });
    await t.api.aprovarPlanejamento('p1');
    ok('14. sai 0701, sem erro', t.ctx.STATE.ordens.some(o => o.os === '0701') && !t.ctx.toasts.some(x => x.startsWith('err')), t.ctx.toasts);
  }

  console.log('\n-- o que nao gera nada --');
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), {
      servidor: (st) => { st.osPlanejadas = []; st.ordens.push({ id: 'y', os: '0701', planejamento: { id: 'p1' } }); }
    });
    await t.api.aprovarPlanejamento('p1');
    ok('15. aprovado em outro computador: nao vira OS de novo', t.ctx.STATE.ordens.length === 2
       && t.ctx.toasts.some(x => /já foi aprovado/.test(x)), t.ctx.toasts);
  }
  {
    const banco = { chamadas: [], async rpc() { return { data: 701, error: null }; } };   // banco que erra
    const t = novo(estado([{ id: 'x', os: '0700' }, { id: 'z', os: '0701' }, { id: 'w', os: '0702' }], [{ id: 'p1', gradeId: 'g', repeticoes: 2 }]), { banco });
    await t.api.aprovarPlanejamento('p1');
    ok('16. numero do bloco ja tomado: nao gera NADA', t.ctx.STATE.ordens.length === 3 && t.ctx.STATE.osPlanejadas.length === 1
       && t.ctx.toasts.some(x => /já é de outra OS/.test(x)), t.ctx.toasts);
  }
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), { servidorFora: true });
    await t.api.aprovarPlanejamento('p1');
    ok('17. sem conseguir reler o servidor: nao gera', t.ctx.STATE.ordens.length === 1 && t.ctx.toasts.some(x => /nada foi gerado/.test(x)), t.ctx.toasts);
  }
  {
    const t = novo(estado([{ id: 'x', os: '0700' }], [{ id: 'p1', gradeId: 'g' }]), { aprova: false });
    await t.api.aprovarPlanejamento('p1');
    ok('18. sem a area Aprovar planejamento: nao aprova', t.ctx.STATE.ordens.length === 1 && !t.ctx.banco.chamadas.length
       && t.ctx.toasts.some(x => /não inclui aprovar/.test(x)), t.ctx.toasts);
  }

  console.log('\n-- toda OS nova nasce no planejamento (conferido no app.js) --');
  ok('19. "Nova OS" abre o planejamento quando nao e uma OS que ja existe',
     /_osPlanModo = _osPlanAbrir \|\| !_formAbreOSExistente\(\)/.test(src));
  ok('20. e o "Salvar OS" recusa OS nova',
     /async function salvarOS\(\) \{[\s\S]{0,300}if \(!_formAbreOSExistente\(\)\)/.test(src));
  ok('21. o botao Aprovar abre para a area concedida', (src.match(/admin-estrito aprova-plan/g) || []).length === 2
     && /body\.pode-aprovar-plan \.admin-estrito\.aprova-plan/.test(fs.readFileSync(path.join(__dirname, '..', 'styles.css'), 'utf8')));
  ok('22. a funcao do banco existe em sql/', fs.existsSync(path.join(__dirname, '..', 'sql', 'supabase-reservar-numeros-os.sql')));

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
