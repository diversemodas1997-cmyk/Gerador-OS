/*
 * OS 0617, "Acabamento de mangas" de 15/10: 670 -> 1340.
 *
 * Junior, 15/10/2026: "corrija para 1340". O 670 foi o numero que o campo
 * trazia pronto (os produtos da OS) antes de as mangas contarem o dobro
 * (_relCostTotalOp no app.js); as 670 pecas tem 1340 mangas.
 *
 *   node servidor/costura-mangas-0617-1510.js            (so mostra)
 *   node servidor/costura-mangas-0617-1510.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();

async function req(method, caminho, body) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: { apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' },
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t;
}
const arr = v => { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; };

(async () => {
  const linha = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = linha.data;
  const lista = arr(data.costuraProd);
  const alvo = lista.filter(l => l.osNumero === '0617' && l.operacao === 'Acabamento de mangas'
    && l.data === '2026-10-15' && Number(l.qtd) === 670);
  if (alvo.length !== 1) {
    console.log(`  achei ${alvo.length} lancamento(s) com 670 -- nada feito`);
    lista.filter(l => l.osNumero === '0617').forEach(l => console.log('   ', l.etapa, '|', l.operacao, l.qtd));
    return;
  }
  const l = alvo[0];
  console.log(`  ${l.pessoaNome} · OS ${l.osNumero} · ${l.etapa} · ${l.operacao} · ${l.qtd} -> 1340`);
  l.qtd = 1340;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const em = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-costura-mangas-0617-' + em.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(linha.data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, { costuraProd: JSON.stringify(lista) });
  const volta = JSON.parse(await req('PATCH',
    `/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(linha.updated_at)}`,
    { data: novo, updated_at: em }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU ENTRE A LEITURA E A ESCRITA -- nada alterado. Rode de novo.');
  const conf = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.costuraProd)
    .find(x => x.id === l.id);
  console.log(`gravado. Agora: ${conf.operacao} = ${conf.qtd}`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
