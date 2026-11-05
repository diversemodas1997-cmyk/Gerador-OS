/* Rode com:  node testes/assistente-status.js

   O ASSISTENTE MUDA STATUS DE OS (05/11/2026, Junior: "crie a ferramenta para
   o assistente mudar status" e "o assistente deve ser capaz de alterar o que
   for impossível do usuário fazer").

   Recorta a mudarStatus de verdade do assistente.js e roda contra um programa
   de mentira (só as funções de status que ela chama). Prende as travas:

     1. sem confirmado, NADA grava — só mostra de → para e põe o botão;
     2. quem pede precisa da área "Status da OS"; a conta do assistente também;
     3. o que o seletor oferece vai pelo mudarStatusOS (o mesmo do clique);
     4. o que o seletor NÃO oferece (voltar etapa) vai pelo carimbo do
        programa, com gravação, estoque e as conjugadas junto;
     5. a observação assinada fica na OS;
     6. o prompt do servidor conhece a ferramenta e não a proíbe mais. */
const fs = require('fs');
const path = require('path');

const raiz = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(raiz, 'assistente.js'), 'utf8');
const ts = fs.readFileSync(path.join(raiz, 'servidor', 'funcoes', 'assistente', 'index.ts'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue); process.exit(1); }
  const j = src.indexOf(ate, i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + ate.length);
}
const fAlvo = recorte('function _statusAlvoDoPedido', '\n  }\n', '_statusAlvoDoPedido');
const fMudar = recorte('async function mudarStatus(p)', '\n  }\n', 'mudarStatus');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const STATUS_OS = [
  { k: 'nao-iniciado', rotulo: 'Não iniciado' },
  { k: 'materia-prima', rotulo: 'Preparando matéria-prima' },
  { k: 'enfestando', rotulo: 'Enfestando' },
  { k: 'cortando', rotulo: 'Cortando' },
  { k: 'parado', rotulo: 'Parado' },
  { k: 'cancelado', rotulo: 'Cancelado' }
];

// Um programa de mentira: sequência linear, seletor só para a frente.
function montar(op) {
  op = op || {};
  const reg = { gravou: 0, estoque: [], seletor: [], botoes: [], carimbos: [] };
  const o = { id: 'a', os: '0704', statusOS: 'enfestando' };
  const c = { id: 'b', os: '0705', statusOS: 'enfestando', segue: 'a' };
  const STATE = { ordens: [o, c] };
  const seqIds = ['materia-prima', 'enfestando', 'cortando'];
  const env = {
    STATE, STATUS_OS,
    STATUS_FORA_SEQ: ['nao-iniciado', 'parado', 'cancelado'],
    STATUS_SEMPRE_OS: ['parado', 'cancelado'],
    CONTA_ASSISTENTE: 'claude.assistent',
    norm: s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(),
    numOS: s => String(s || '').replace(/\D/g, '').replace(/^0+/, ''),
    rotuloStatus: k => (STATUS_OS.find(x => x.k === k) || {}).rotulo || k,
    statusDe: x => x.statusOS || 'nao-iniciado',
    login: () => 'junior',
    ehAdmin: () => !!op.admin,
    contaTemAcesso: (area, conta) => area === 'os-status' && (conta !== 'claude.assistent' || op.assistenteTem !== false),
    podeGravar: () => true,
    podeMudarStatusOS: () => op.pessoaTem !== false,
    _statusSeqDaOS: () => seqIds.map(k => ({ k })),
    _statusSeqExpandida: seq => seq.map(x => ({ k: x.k, p: 1, id: x.k })),
    _statusRepSeq: () => 1,
    _statusRepetidosSeq: () => [],
    _statusIdPasso: (k, p) => (p > 1 ? k + '#' + p : k),
    _statusDoId: id => ({ k: String(id).split('#')[0], p: 1 }),
    _statusNoOS: x => x.statusOS || 'nao-iniciado',
    _statusOS: x => x.statusOS || 'nao-iniciado',
    _statusDoChecklistOS: () => op.folha || '',
    // Seletor: o de agora, o seguinte, Parado e Cancelado.
    _statusOpcoesOS: x => {
      const cam = ['nao-iniciado'].concat(seqIds);
      const i = cam.indexOf(x.statusOS || 'nao-iniciado');
      return [cam[i], cam[i + 1], 'parado', 'cancelado'].filter(Boolean);
    },
    _conjugadasQueSeguemStatus: (x, alvo) => STATE.ordens.filter(y => y.segue === x.id && (y.statusOS || 'nao-iniciado') !== alvo),
    _obsNotas: x => x.obsNotas || [],
    _carimbarStatusOS: (x, alvo) => { reg.carimbos.push(x.os + ':' + alvo); if (alvo === 'nao-iniciado') delete x.statusOS; else x.statusOS = alvo; },
    mudarStatusOS: async (id, alvo) => { reg.seletor.push(id + ':' + alvo); const x = STATE.ordens.find(y => y.id === id); x.statusOS = alvo; reg.gravou++; },
    saveState: async () => { reg.gravou++; },
    _estoqueSeguirStatusOS: async (x, alvo) => { reg.estoque.push(x.os + ':' + alvo); },
    renderListaOS() {}, renderStatusFolhaOS() {}, _redesenharCampoAtivo() {},
    oferecerAcao: (rotulo, fn) => reg.botoes.push({ rotulo, fn }),
    anotarNoChat: () => {}
  };
  const nomes = Object.keys(env);
  const mudarStatus = new Function(...nomes, `${fAlvo}\n${fMudar}\nreturn mudarStatus;`)(...nomes.map(n => env[n]));
  return { mudarStatus, reg, o, c };
}

(async () => {
  console.log('-- sem confirmado, nada grava --');
  {
    const t = montar();
    const r = await t.mudarStatus({ os: '704', status: 'Não iniciado' });
    ok('1. devolve de -> para e aguarda', r.aguardando_confirmacao && r.de === 'Enfestando' && r.para === 'Não iniciado', JSON.stringify(r));
    ok('2. nada gravado, status intacto', t.reg.gravou === 0 && t.o.statusOS === 'enfestando' && !t.o.obsNotas, JSON.stringify(t.reg));
    ok('3. um botão no chat', t.reg.botoes.length === 1 && /0704: Enfestando → Não iniciado/.test(t.reg.botoes[0].rotulo), JSON.stringify(t.reg.botoes));
    ok('4. diz que é para trás e que o seletor não deixa', r.sentido === 'para trás' && r.pelo_seletor === false, JSON.stringify(r));
    ok('5. avisa a conjugada que vai junto', r.conjugadas_que_vao_junto.length === 1 && /0705/.test(r.conjugadas_que_vao_junto[0]), JSON.stringify(r.conjugadas_que_vao_junto));
    // o clique no botão grava
    await t.reg.botoes[0].fn();
    ok('6. o clique do botão grava', !t.o.statusOS && t.reg.gravou > 0, JSON.stringify(t.o));
  }

  console.log('');
  console.log('-- voltar etapa (o que a tela não deixa) --');
  {
    const t = montar();
    const r = await t.mudarStatus({ os: '0704', status: 'nao-iniciado', confirmado: true, motivo: 'carimbada errada' });
    ok('7. mudou para Não iniciado', r.mudado && r.agora === 'Não iniciado' && !t.o.statusOS, JSON.stringify(r));
    ok('8. pelo carimbo do programa, não pelo seletor', t.reg.seletor.length === 0 && t.reg.carimbos.includes('0704:nao-iniciado'), JSON.stringify(t.reg));
    ok('9. a conjugada foi junto', t.reg.carimbos.includes('0705:nao-iniciado') && !t.c.statusOS, JSON.stringify(t.reg.carimbos));
    ok('10. gravou e acertou o estoque das duas', t.reg.gravou === 1 && t.reg.estoque.join() === '0704:nao-iniciado,0705:nao-iniciado', JSON.stringify(t.reg));
    const n = (t.o.obsNotas || [])[0] || {};
    ok('11. observação assinada com de -> para e motivo', n.login === 'junior' && /Enfestando → Não iniciado/.test(n.texto) && /carimbada errada/.test(n.texto), JSON.stringify(n));
  }

  console.log('');
  console.log('-- o que o seletor oferece vai pelo clique de verdade --');
  {
    const t = montar();
    const r = await t.mudarStatus({ os: '704', status: 'cortando', confirmado: true });
    ok('12. pelo mudarStatusOS', t.reg.seletor.join() === 'a:cortando' && t.reg.carimbos.length === 0, JSON.stringify(t.reg));
    ok('13. resultado diz o status de agora', r.mudado && r.agora === 'Cortando' && r.pelo_seletor === true, JSON.stringify(r));
  }

  console.log('');
  console.log('-- as travas de acesso --');
  {
    const t = montar({ pessoaTem: false });
    const r = await t.mudarStatus({ os: '704', status: 'nao-iniciado', confirmado: true });
    ok('14. quem pede sem "Status da OS": sem_acesso e nada muda', r.sem_acesso && t.o.statusOS === 'enfestando' && t.reg.gravou === 0, JSON.stringify(r));
  }
  {
    const t = montar({ pessoaTem: false, admin: true });
    const r = await t.mudarStatus({ os: '704', status: 'nao-iniciado', confirmado: true });
    ok('15. admin passa', r.mudado === true, JSON.stringify(r));
  }
  {
    const t = montar({ assistenteTem: false });
    const r = await t.mudarStatus({ os: '704', status: 'nao-iniciado', confirmado: true });
    ok('16. conta do assistente sem a área: sem_acesso', r.sem_acesso && /claude\.assistent/.test(r.mensagem) && t.reg.gravou === 0, JSON.stringify(r));
  }

  console.log('');
  console.log('-- pedidos tortos --');
  {
    const t = montar();
    const r1 = await t.mudarStatus({ os: '704', status: 'voando' });
    ok('17. status desconhecido: erro com a lista', r1.erro && Array.isArray(r1.status_validos), JSON.stringify(r1));
    const r2 = await t.mudarStatus({ os: '9999', status: 'Parado' });
    ok('18. OS que não existe: erro', /9999/.test(r2.erro || ''), JSON.stringify(r2));
    const r3 = await t.mudarStatus({ os: '704', status: 'Enfestando' });
    ok('19. já está lá (e a conjugada também): nada a fazer', r3.ja_estava === true && t.reg.botoes.length === 0, JSON.stringify(r3));
  }
  {
    const t = montar({ folha: 'cortando' });
    const r = await t.mudarStatus({ os: '704', status: 'Não iniciado' });
    ok('20. avisa quando a folha tem etapa marcada', r.avisos.some(a => /etapa marcada/.test(a)), JSON.stringify(r.avisos));
  }

  console.log('');
  console.log('-- o servidor (prompt e lista de ferramentas) --');
  ok('21. a ferramenta está na lista do servidor', /name: 'mudar_status_os'/.test(ts), '');
  const nunca = (ts.match(/O que você NUNCA faz[^\n]*/) || [''])[0];
  ok('22. "mudar status" saiu do NUNCA', nunca && !/mudar status/i.test(nunca), nunca.slice(0, 160));
  ok('23. e o navegador executa a ferramenta', /mudar_status_os\(p\)\s*\{\s*return mudarStatus\(p\)/.test(src), '');

  console.log('');
  console.log(falhas ? `${falhas} FALHA(S)` : 'tudo ok');
  process.exit(falhas ? 1 : 0);
})();
