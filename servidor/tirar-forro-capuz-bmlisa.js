/*
 * TIRA O "FORRO DO CAPUZ" DAS OS DE BM.LISA.
 *
 * Junior, 25/09/2026: "BM.LISA NAO TEM FORRO DE CAPUZ" e "sim, tire o forro
 * de capuz dessas OS". Varias OS de BM.LISA traziam o componente "Forro do
 * capuz" (2 por peca) -- resto de cadastro. Ele nao sai na folha de OE, mas
 * entra no que le os componentes da OS (produtos por tecido+cor dos estoques,
 * custo). So a BM.TRI tem forro de capuz.
 *
 * Mexe em DUAS coisas:
 *   - o componente cujo nome comeca com "forro" nas OS cuja linha de SKU e
 *     BM.LISA (o.componentes);
 *   - o mesmo componente nos DESENHOS TECNICOS de BM.LISA, se houver -- senao
 *     o "Repor componentes do desenho" traria o forro de volta.
 * Nada mais: fases, grade, estoque e cargas ficam como estao.
 *
 *   node servidor/tirar-forro-capuz-bmlisa.js            (so mostra)
 *   node servidor/tirar-forro-capuz-bmlisa.js --gravar   (grava)
 *
 * Antes de gravar salva o blob inteiro em backups/. A escrita so passa se
 * ninguem gravou no servidor desde a leitura (carimbo updated_at).
 * Depois de gravar: F5 no programa, senao aba aberta regrava o estado velho.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const normNome = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const ehForro = c => /^forro/.test(normNome(c && c.nome));

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

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { return typeof data[k] === 'string' ? JSON.parse(data[k]) : (data[k] || []); } catch (e) { return []; } };
  const ordens = ler('ordens'), desenhos = ler('desenhos'), modelos = ler('modelos');

  // A linha do SKU, como o app (_skuBaseDaOS): override > desenho > modelo, sem a cor.
  const linhaOS = o => {
    const d = desenhos.find(x => x.id === o.desenhoId);
    const m = modelos.find(x => x.id === o.modeloId);
    const sku = String(o.skuOverride || (d && d.skuLinha) || (m && m.skuLinha) || '').trim().toUpperCase();
    return sku.split('-')[0];
  };
  const linhaDesenho = d => String((d && d.skuLinha) || '').trim().toUpperCase().split('-')[0];

  let nOS = 0, nDes = 0;
  console.log('OS de BM.LISA com forro:');
  ordens.forEach(o => {
    if (linhaOS(o) !== 'BM.LISA') return;
    const f = (o.componentes || []).filter(ehForro);
    if (!f.length) return;
    nOS++;
    console.log('  OS ' + o.os + '  -> tira ' + f.map(c => '"' + c.nome + '" (' + c.qtdPorPeca + '/pc, total ' + (c.qtdTotal || 0) + ')').join(', '));
    o.componentes = o.componentes.filter(c => !ehForro(c));
  });
  // Desenhos: o campo de componentes pode ter outro nome -- procura qualquer
  // lista de objetos com `nome` que tenha forro.
  console.log('\nDesenhos tecnicos de BM.LISA com forro:');
  desenhos.forEach(d => {
    if (linhaDesenho(d) !== 'BM.LISA') return;
    Object.keys(d).forEach(k => {
      if (!Array.isArray(d[k]) || !d[k].some(x => x && typeof x === 'object' && 'nome' in x)) return;
      const f = d[k].filter(ehForro);
      if (!f.length) return;
      nDes++;
      console.log('  desenho ' + (d.codigo || d.nome || d.id) + ' [' + k + '] -> tira ' + f.map(c => '"' + c.nome + '"').join(', '));
      d[k] = d[k].filter(c => !ehForro(c));
    });
  });
  console.log('\n' + nOS + ' OS e ' + nDes + ' lista(s) de desenho a corrigir.');
  if (!nOS && !nDes) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-tirar-forro-bmlisa-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);

  const novo = Object.assign({}, data);
  const volta = (k, v) => (typeof data[k] === 'string' ? JSON.stringify(v) : v);
  if (nOS) novo.ordens = volta('ordens', ordens);
  if (nDes) novo.desenhos = volta('desenhos', desenhos);
  await gravarBlob(cab, novo, updatedAt);

  // Confere lendo de volta.
  const { data: d2 } = await lerBlob(cab);
  const o2 = typeof d2.ordens === 'string' ? JSON.parse(d2.ordens) : d2.ordens;
  const resta = o2.filter(o => linhaOS(o) === 'BM.LISA' && (o.componentes || []).some(ehForro)).length;
  console.log('gravado. OS de BM.LISA ainda com forro: ' + resta);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
