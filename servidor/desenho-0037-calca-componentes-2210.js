/*
 * Desenho 0037 (Calça Moletom, CJM): tira as peças da BLUSA e põe tudo em
 * Preto Moletom. Corrige também a OS 0676, que nasceu dele.
 *
 * POR QUE ISTO EXISTE
 *
 * O desenho 0037 tinha nove componentes, e três eram da Blusa Moletom Básica:
 * "Frente Blusa Moletom Básica", "Mangas Blusa Moletom Básica" e "Costas Blusa
 * Moletom Básica". Os três já tinham sido apagados do cadastro de componentes,
 * mas continuavam ligados ao desenho. E, tirando a Frente D, todos estavam na
 * cor "Preto Malha Algodão", em tecido Moletom. A OS 0676 copiou isso tudo
 * (a 0589, emitida antes, tinha cinco componentes).
 *
 * Pedido do Junior em 22/10/2026, junto com a correção das etiquetas: a calça
 * imprimia manga, barra, punhos — peças que ela não tem.
 *
 * COMO RODAR
 *
 *   node servidor/desenho-0037-calca-componentes-2210.js            só relata
 *   node servidor/desenho-0037-calca-componentes-2210.js --gravar   grava
 *
 * Antes de gravar salva o blob em backups/, e a escrita confere o carimbo.
 *
 * ATENÇÃO: recarregue a página (F5) depois de rodar. Aba aberta desde antes
 * regrava a chave inteira com o estado velho e desfaz isto.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const DESENHO = '0037';
const OS_CORRIGIR = ['0676'];
const DA_BLUSA = /blusa/i;           // as três peças têm "Blusa" no nome
const COR_ALVO = 'Preto Moletom';

async function lerBlob() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at, cab };
}

async function gravarBlob(cab, data, updatedAt) {
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
}

(async () => {
  const { data, updatedAt, cab } = await lerBlob();
  const le = k => (typeof data[k] === 'string' ? JSON.parse(data[k] || '[]') : (data[k] || []));
  const desenhos = le('desenhos'), ordens = le('ordens'), cores = le('cores');

  const cor = cores.find(c => c.nome === COR_ALVO);
  if (!cor) throw new Error('nao achei a cor "' + COR_ALVO + '" no cadastro');
  const nomeCor = id => (cores.find(c => c.id === id) || {}).nome || '(sem cor)';

  const d = desenhos.find(x => x.codigo === DESENHO);
  if (!d) throw new Error('nao achei o desenho ' + DESENHO);

  let mudou = 0;
  console.log('Desenho ' + DESENHO + ':');
  const antes = d.componentes || [];
  const fora = antes.filter(c => DA_BLUSA.test(c.nome || ''));
  fora.forEach(c => console.log('   tira  "' + c.nome + '"'));
  d.componentes = antes.filter(c => !DA_BLUSA.test(c.nome || ''));
  mudou += fora.length;
  if (Array.isArray(d.componentesIds)) {
    const idsFora = new Set(fora.map(c => c.componenteId));
    const n = d.componentesIds.length;
    d.componentesIds = d.componentesIds.filter(id => !idsFora.has(id));
    if (d.componentesIds.length !== n) console.log('   componentesIds: ' + n + ' -> ' + d.componentesIds.length);
  }
  d.componentes.forEach(c => {
    console.log('   (fica) "' + c.nome + '" em ' + nomeCor(c.corId) + ' [' + c.corId + ']');
    if (c.corId === cor.id) return;
    console.log('   "' + c.nome + '": ' + nomeCor(c.corId) + ' -> ' + COR_ALVO);
    c.corId = cor.id;
    mudou++;
  });

  OS_CORRIGIR.forEach(num => {
    const o = ordens.find(x => x.os === num);
    if (!o) { console.log('\nOS ' + num + ': nao achei, pulada'); return; }
    if (o.desenhoId !== d.id) { console.log('\nOS ' + num + ': nao e do desenho ' + DESENHO + ', pulada'); return; }
    console.log('\nOS ' + num + ':');
    const comps = o.componentes || [];
    comps.filter(c => DA_BLUSA.test(c.nome || '')).forEach(c => { console.log('   tira  "' + c.nome + '"'); mudou++; });
    o.componentes = comps.filter(c => !DA_BLUSA.test(c.nome || ''));
    o.componentes.forEach(c => {
      if (c.cor === cor.id && c.corNome === cor.nome) return;
      console.log('   "' + c.nome + '": ' + (c.corNome || '(sem cor)') + ' -> ' + COR_ALVO);
      c.cor = cor.id;
      c.corNome = cor.nome;
      mudou++;
    });
  });

  const outras = ordens.filter(o => o.desenhoId === d.id && !OS_CORRIGIR.includes(o.os));
  outras.forEach(o => console.log('\n(nao mexida) OS ' + o.os + ': '
    + (o.componentes || []).map(c => c.nome + ' [' + (c.corNome || '') + ']').join(', ')));

  console.log('\nalteracoes: ' + mudou);
  if (!mudou) { console.log('Nada a corrigir.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-desenho-0037-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  const escreve = (k, v) => { data[k] = typeof data[k] === 'string' ? JSON.stringify(v) : v; };
  escreve('desenhos', desenhos);
  escreve('ordens', ordens);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
