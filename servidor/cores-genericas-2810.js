/*
 * AS CORES PASSAM A SER GENERICAS, DE TODOS OS TECIDOS.
 *
 * Jones, 28/10/2026: "corrija o cadastro de cores, para que haja apenas cores
 * genericas: preto, branco, verde, etc. Mantenha as siglas do SKU nas cores,
 * mas retire o tecido, pois as cores servirao para todos os tecidos" e
 * "Atualize todos os cadastros que utilizam os cadastros de cores se for
 * necessario".
 *
 * Antes havia uma cor por tecido: "Preto Malha Algodao", "Preto Moletom",
 * "Preto Ribana Moletom", "Preto Rugao"... (53 cadastros). Agora fica uma so,
 * "Preto", com a Sigla SKU. O tecido vem da linha (OS, desenho, estoque).
 *
 * O QUE O SCRIPT FAZ
 *   1. Agrupa as cores pelo nome sem o tecido (o mesmo corte do app:
 *      corBaseNome, com o nome curto do Texturizado — "Rugao", "Prime").
 *   2. Cada grupo vira UMA cor, que fica com o id do cadastro de Malha Algodao
 *      (ou o primeiro do grupo): nome generico, hex e codigo dele, sem tecido.
 *      A Sigla SKU e a mais usada no grupo (sem traco — a com traco e SKU
 *      inteiro, nao sigla); Off-White fica OFF (a dos SKUs COT.*-OFF) e
 *      Grafite, que nao tinha, fica GRAF (a do CM.LISA-GRAF).
 *   3. O que o cadastro antigo sabia e nao cabe na cor generica fica guardado
 *      nela: `legado` (id, nome, tecido, sigla, hex, codigo de cada antiga) e
 *      `fornecedorPorTecido` (o codigo/nome da cor no fornecedor, que e de um
 *      tecido so — "PRETO ENXOFRE 020104" da Ribana Moletom).
 *   4. Em TODOS os cadastros e lancamentos, o id antigo vira o id da generica
 *      e o nome antigo vira o nome generico — so valor IGUAL ao inteiro (nunca
 *      pedaco de texto): desenhos (cores e componentes), OS (variantes,
 *      tecidos, componentes, fases, enfesto, foto do desenho), estoque de
 *      tecidos, compras, rolos, fios/linhas e o snapshot da Contabilidade.
 *      O estoque de tecido continua separado por tecido + cor: "Malha Algodao
 *      · Preto Malha Algodao" vira "Malha Algodao · Preto".
 *   Fica de fora o Diario (e o historico do que aconteceu, com os nomes de
 *   entao).
 *
 *   node servidor/cores-genericas-2810.js            (so mostra)
 *   node servidor/cores-genericas-2810.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const FORA = new Set(['cores', 'diario', '_device']);
// A sigla que o grupo tem de ter (o resto sai da mais usada).
const SIGLA_FIXA = { 'off-white': 'OFF', grafite: 'GRAF' };

const normNome = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();

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
async function lerBlob(cab) {
  const l = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  return { data: l[0].data, updatedAt: l[0].updated_at };
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
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  }
}
// Avisa os programas abertos para relerem (a mesma tabela-sinal do app).
async function sinalizar(cab) {
  try {
    await fetch(`${SUPA}/rest/v1/sync_signal?on_conflict=id`, {
      method: 'POST',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates' }),
      body: JSON.stringify({ id: 'main', updated_at: new Date().toISOString(), device_id: 'script-cores-genericas' })
    });
  } catch (e) { /* sem sinal, o polling do app acha a mudanca */ }
}

// O nome sem o tecido: o maior tecido que casa como sufixo (com o nome curto
// do Texturizado, o que vem depois da barra). Devolve as palavras originais.
function semTecido(nome, tecidos) {
  const n = normNome(nome);
  let suf = '';
  tecidos.forEach(t => {
    const tn = normNome(t.nome);
    const curto = tn.includes('|') ? tn.split('|').pop().trim() : '';
    [tn, curto].forEach(x => { if (x && x.length > suf.length && n.endsWith(' ' + x)) suf = x; });
  });
  if (!suf) return String(nome).trim();
  const p = String(nome).trim().split(/\s+/);
  return p.slice(0, Math.max(1, p.length - suf.split(' ').length)).join(' ');
}

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return v; };
  const cores = ler('cores') || [];
  const tecidos = ler('tecidos') || [];
  const tecNome = id => ((tecidos.find(t => t.id === id) || {}).nome || '');
  if (!cores.length) throw new Error('cadastro de cores vazio');

  // 1) os grupos
  const grupos = new Map();
  cores.forEach(c => {
    const base = semTecido(c.nome, tecidos);
    const k = normNome(base);
    if (!grupos.has(k)) grupos.set(k, { k, base, membros: [] });
    grupos.get(k).membros.push(c);
  });
  if ([...grupos.values()].every(g => g.membros.length === 1 && normNome(g.membros[0].nome) === g.k && !g.membros[0].tecidoId)) {
    console.log('As cores ja sao genericas. Nada a fazer.'); return;
  }

  // 2) a cor generica de cada grupo
  const mapaId = new Map(), mapaNome = new Map();
  const novas = [];
  const titulo = s => s.split(/(\s+|-)/).map(w => (/^[a-zà-ú]/i.test(w) ? w.charAt(0).toUpperCase() + w.slice(1) : w)).join('');
  [...grupos.values()].forEach(g => {
    const m = g.membros;
    const canon = m.find(c => normNome(tecNome(c.tecidoId)) === 'malha algodao' || /malha algodao$/.test(normNome(c.nome))) || m[0];
    const conta = new Map();
    m.forEach(c => { const s = String(c.siglaSku || '').trim().toUpperCase(); if (s && !s.includes('-')) conta.set(s, (conta.get(s) || 0) + 1); });
    const maisUsada = [...conta.entries()].sort((a, b) => b[1] - a[1] || (a[0] === String(canon.siglaSku || '').toUpperCase() ? -1 : 1))[0];
    const sigla = SIGLA_FIXA[g.k] || (maisUsada ? maisUsada[0] : '');
    const nome = titulo(semTecido(canon.nome, tecidos));
    const forn = m.filter(c => c.fornecedorId || c.codigoFornecedor || c.nomeFornecedor)
      .map(c => ({ corAntiga: c.nome, tecidoId: c.tecidoId || '', tecidoNome: tecNome(c.tecidoId) || semTecido(c.nome, []).slice(nome.length).trim(),
        fornecedorId: c.fornecedorId || '', codigoFornecedor: c.codigoFornecedor || '', nomeFornecedor: c.nomeFornecedor || '' }));
    const nova = { id: canon.id, nome, hex: canon.hex || '', codigo: canon.codigo || '', siglaSku: sigla,
      fornecedorId: '', codigoFornecedor: '', nomeFornecedor: '' };
    if (forn.length) nova.fornecedorPorTecido = forn;
    nova.legado = m.map(c => ({ id: c.id, nome: c.nome, tecidoId: c.tecidoId || '', siglaSku: c.siglaSku || '', hex: c.hex || '', codigo: c.codigo || '' }));
    novas.push(nova);
    m.forEach(c => { if (c.id !== nova.id) mapaId.set(c.id, nova.id); if (c.nome !== nova.nome) mapaNome.set(c.nome, nova.nome); });
    const hexes = [...new Set(m.map(c => String(c.hex || '').toLowerCase()).filter(Boolean))];
    console.log(`  ${nome.padEnd(10)} sigla ${String(sigla || '—').padEnd(9)} ${hexes.length > 1 ? 'hex ' + nova.hex + ' (havia ' + hexes.join(' ') + ')' : 'hex ' + nova.hex}  <- ${m.map(c => c.nome + (c.siglaSku ? ' [' + c.siglaSku + ']' : '')).join(', ')}`);
  });
  novas.sort((a, b) => String(a.codigo).localeCompare(String(b.codigo), 'pt-BR', { numeric: true }));

  // 3) as referencias, em todas as colecoes
  const trocas = {};
  const anda = (v, k) => {
    if (typeof v === 'string') {
      const n = mapaId.get(v) || mapaNome.get(v);
      if (n !== undefined) { trocas[k] = (trocas[k] || 0) + 1; return n; }
      return v;
    }
    if (Array.isArray(v)) { for (let i = 0; i < v.length; i++) v[i] = anda(v[i], k); return v; }
    if (v && typeof v === 'object') { for (const q of Object.keys(v)) v[q] = anda(v[q], k); return v; }
    return v;
  };
  const novoData = Object.assign({}, data);
  Object.keys(data).forEach(k => {
    if (FORA.has(k)) return;
    const eraTexto = typeof data[k] === 'string';
    let v;
    try { v = eraTexto ? JSON.parse(data[k]) : JSON.parse(JSON.stringify(data[k])); } catch (e) { return; }
    const antes = trocas[k] || 0;
    v = anda(v, k);
    if ((trocas[k] || 0) !== antes) novoData[k] = eraTexto ? JSON.stringify(v) : v;
  });
  novoData.cores = typeof data.cores === 'string' ? JSON.stringify(novas) : novas;

  console.log(`\n${cores.length} cores -> ${novas.length} genericas.`);
  console.log('Referencias trocadas: ' + (Object.entries(trocas).map(([k, n]) => k + ' ' + n).join(' · ') || 'nenhuma'));

  // Conferencia: nada pode sobrar apontando para cor que nao existe mais.
  const idsNovos = new Set(novas.map(c => c.id)), nomesVelhos = new Set(mapaNome.keys()), idsVelhos = new Set(mapaId.keys());
  let sobra = 0;
  const confere = v => {
    if (typeof v === 'string') { if (idsVelhos.has(v) || nomesVelhos.has(v)) sobra++; return; }
    if (Array.isArray(v)) v.forEach(confere); else if (v && typeof v === 'object') Object.values(v).forEach(confere);
  };
  Object.keys(novoData).forEach(k => { if (FORA.has(k)) return; try { confere(typeof novoData[k] === 'string' ? JSON.parse(novoData[k]) : novoData[k]); } catch (e) {} });
  console.log('Sobrou referencia a cor antiga: ' + sobra + ' · cores genericas: ' + idsNovos.size);
  if (sobra) throw new Error('sobrou referencia a cor antiga -- nada gravado');

  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cores-genericas-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, novoData, updatedAt);
  await sinalizar(cab);
  const d2 = (await lerBlob(cab)).data;
  const c2 = typeof d2.cores === 'string' ? JSON.parse(d2.cores) : d2.cores;
  console.log('gravado. Cores no servidor: ' + c2.length + ' (' + c2.map(c => c.nome).join(', ') + ')');
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
