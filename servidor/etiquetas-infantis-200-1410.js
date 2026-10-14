/*
 * ETIQUETAS 2 AO 16 VOLTAM A 200 EM DESCALVADO.
 *
 * Junior, 14/10/2026: "Corrija as entradas de etiqueta tamanho 2 até o 16,
 * para que contenha 200 em estoque" — "unidade descalvado".
 *
 * Em 30/09 entraram 200 de cada tamanho (etiquetas-infantis-3009.js). Em
 * 14/10, às 11:53, entrou mais uma ENTRADA de 400 em sete tamanhos (2, 4, 6,
 * 8, 10, 12 e 16; o 14 ficou de fora), e o estoque desses foi a 600. Nenhuma
 * OS infantil chegou à costura — as ensacadas e as não iniciadas só reservam —,
 * então o estoque de cada tamanho é a soma das entradas. Tirar essas sete
 * entradas de 400 deixa os oito tamanhos em 200.
 *
 * Só apaga entrada de Etiqueta, unidade Descalvado, tamanho 2 a 16, de 400 un,
 * com data 2026-10-14. Se alguma não existir mais, segue com as outras.
 *
 *   node servidor/etiquetas-infantis-200-1410.js            (so mostra)
 *   node servidor/etiquetas-infantis-200-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
// O Kong por dentro da máquina: o endereço da rede mudou com o roteador
// (14/10/2026) e o localhost não muda nunca.
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const TAMS = ['2', '4', '6', '8', '10', '12', '14', '16'];

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

const arr = s => (typeof s === 'string' ? JSON.parse(s || '[]') : (s || []));
const ehEtiquetaInfantil = m => m.item === 'Etiqueta' && TAMS.includes(String(m.tam)) && (m.unidade || 'desc') === 'desc';
const ehAlvo = m => ehEtiquetaInfantil(m) && m.tipo === 'entrada' && Number(m.qtd) === 400 && m.data === '2026-10-14';

function saldos(movs) {
  const s = {};
  TAMS.forEach(t => { s[t] = 0; });
  movs.filter(ehEtiquetaInfantil).forEach(m => {
    if (m.tipo === 'entrada') s[m.tam] += Number(m.qtd) || 0;
    else if (m.tipo === 'saida') s[m.tam] -= Number(m.qtd) || 0;
  });
  return TAMS.map(t => `${t}: ${s[t]}`).join(' · ');
}

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const movs = arr(data.aviamentosMov);
  const alvo = movs.filter(ehAlvo);
  console.log('Entradas de etiqueta 2 ao 16 em Descalvado (soma por tamanho):');
  console.log('  antes:  ' + saldos(movs));
  if (!alvo.length) { console.log('  Nenhuma entrada de 400 de 14/10 — nada a fazer.'); return; }
  console.log(`  a apagar (${alvo.length}): ` + alvo.map(m => `tam ${m.tam} +${m.qtd} (${m.em || m.data}, ${m.por || ''})`).join('; '));
  const saida = movs.filter(m => !ehAlvo(m));
  console.log('  depois: ' + saldos(saida));
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-etiquetas-infantis-200-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));

  const novo = Object.assign({}, data, {
    aviamentosMov: typeof data.aviamentosMov === 'string' ? JSON.stringify(saida) : saida,
    _device: 'etiquetas-infantis-200-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');

  const m2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.aviamentosMov);
  console.log('gravado. Agora: ' + saldos(m2) + ` (${m2.length} lançamentos no blob)`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
