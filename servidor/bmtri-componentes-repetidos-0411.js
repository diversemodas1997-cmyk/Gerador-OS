/*
 * BM.TRI 0024 E 0025 SEM OS COMPONENTES REPETIDOS.
 *
 * Junior, 04/11/2026: "pode limpar os componentes repetidos do 0024 e 0025".
 *
 * Os desenhos 0024 (Verde/Preto/Bege) e 0025 (Preto/Mostarda/Off-White) tem os
 * 15 componentes de hoje ("Frente parte 1"...) e, alem deles, os 9 de abril
 * ("Frente/Costa/Mangas PARTE 1/2/3 Blusa Moletom Tricolor") com a mesma cor,
 * tecido e quantidade. Ver bmtri-azul-componentes-0411.js, que tirou os mesmos
 * 9 do 0049.
 *
 * SO SAI O QUE TEM GEMEO IDENTICO. Para cada antigo, procura o componente de
 * hoje de mesmo nome ("Frente PARTE 2 Blusa..." = "Frente parte 2") e o tira
 * so se cor, tecido, por peca e quantidades por tamanho forem iguais. Isso
 * preserva as 10 OS antigas (0287, 0342, 0350, 0356, 0405, 0423, 0435, 0436,
 * 0463, 0488) que tem SO os 9 antigos para Frente, Costa e Mangas: neles os
 * antigos sao a peca, nao repeticao. Se uma OS tiver antigo sem gemeo, ela
 * inteira fica como esta e o script diz qual.
 *
 * Nao muda estoque nem producao: produtosOS e produtosPorTecidoCorOS pegam o
 * MAIOR componente de cada tamanho / tecido+cor, nao a soma, e o gemeo
 * identico fica.
 *
 *   node servidor/bmtri-componentes-repetidos-0411.js            (so mostra)
 *   node servidor/bmtri-componentes-repetidos-0411.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const DESENHOS = ['0024', '0025'];
const VELHO = /^(Frente|Costa|Mangas) PARTE ([123]) Blusa Moletom Tricolor$/;
const norm = s => (s || '').toString().toLowerCase().replace(/\s+/g, ' ').trim();
const gemeoNome = nome => { const m = VELHO.exec(nome || ''); return m ? norm(m[1] + ' parte ' + m[2]) : null; };
const igualJSON = (a, b) => JSON.stringify(a || {}) === JSON.stringify(b || {});

async function conectar() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  return { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
}

// Tira os antigos com gemeo identico. `mesmo(a, b)` diz se dois componentes
// batem. Devolve a lista nova, ou null se algum antigo nao tiver gemeo.
function limpar(comps, mesmo) {
  const velhos = comps.filter(c => VELHO.test(c.nome || ''));
  if (!velhos.length) return { lista: comps, saem: 0 };
  const novos = comps.filter(c => !VELHO.test(c.nome || ''));
  const sem = velhos.filter(v => !novos.some(n => norm(n.nome) === gemeoNome(v.nome) && mesmo(n, v)));
  if (sem.length) return { lista: null, sem };
  return { lista: novos, saem: velhos.length };
}

(async () => {
  const cab = await conectar();
  const l = await (await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  const data = l[0].data, updatedAt = l[0].updated_at;
  const ler = k => (typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]);
  const pôr = (k, v) => { data[k] = (typeof data[k] === 'string') ? JSON.stringify(v) : v; };

  const desenhos = ler('desenhos');
  const ordens = ler('ordens');
  let mudou = 0;

  const ids = new Set();
  for (const cod of DESENHOS) {
    const d = desenhos.find(x => x.codigo === cod);
    if (!d) throw new Error('desenho ' + cod + ' nao encontrado');
    ids.add(d.id);
    const r = limpar(d.componentes || [], (n, v) =>
      n.corId === v.corId && n.tecidoId === v.tecidoId && n.qtdPorPeca === v.qtdPorPeca);
    if (!r.lista) { console.log(`desenho ${cod}: FICA (antigo sem gemeo: ${r.sem.map(c => c.nome).join(', ')})`); continue; }
    console.log(`desenho ${cod} ${d.desc}: ${d.componentes.length} -> ${r.lista.length} componentes`);
    if (r.saem) { d.componentes = r.lista; mudou++; }
  }

  const ficam = [];
  ordens.filter(o => ids.has(o.desenhoId)).forEach(o => {
    const r = limpar(o.componentes || [], (n, v) =>
      n.corNome === v.corNome && n.material === v.material && n.qtdPorPeca === v.qtdPorPeca
      && igualJSON(n.qtdPorTamanho, v.qtdPorTamanho) && (Number(n.qtdTotal) || 0) === (Number(v.qtdTotal) || 0));
    if (!r.lista) { ficam.push(o.os); return; }
    if (!r.saem) return;
    console.log(`  OS ${o.os} (${o.statusOS}): ${o.componentes.length} -> ${r.lista.length}`);
    o.componentes = r.lista; mudou++;
  });
  if (ficam.length) console.log(`  ficam como estao (antigos sao a peca, sem gemeo): ${ficam.join(', ')}`);

  if (!mudou) { console.log('\nnada a mudar'); return; }
  if (!GRAVAR) { console.log(`\n${mudou} registro(s) a mudar (so mostrou; rode com --gravar para gravar)`); return; }

  const bk = path.join(RAIZ, 'backups', `shared_data-antes-bmtri-componentes-repetidos-${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
  fs.writeFileSync(bk, JSON.stringify(l[0]));
  console.log('backup:', bk);

  pôr('desenhos', desenhos);
  pôr('ordens', ordens);
  const r = await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
    body: JSON.stringify({ data, updated_at: new Date().toISOString() })
  });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  try {
    await fetch(`${SUPA}/rest/v1/sync_signal?on_conflict=id`, {
      method: 'POST',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates' }),
      body: JSON.stringify({ id: 'main', updated_at: new Date().toISOString(), device_id: 'script-bmtri-repetidos' })
    });
  } catch (e) { /* o polling do app acha a mudanca */ }
  console.log('gravado.');
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
