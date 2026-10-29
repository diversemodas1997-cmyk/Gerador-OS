/*
 * O MODELO "SHORT FEMININO" PERDE O SKU SF.BERM.SEM.
 *
 * Junior, 29/10/2026: "Retire o SKU SF.BERM.SEM". O SKU estava só no cadastro
 * do modelo; os dois desenhos dele (0063 preto, 0064 azul)
 * já foram para SF.BERM.FEM. Por causa da diferença, o Estoque de produtos
 * mostrava a linha SF.BERM.SEM sozinha, sem cor, ao lado das dos desenhos.
 * Nenhuma OS nem lançamento de estoque usa o SF.BERM.SEM. Aqui o modelo passa
 * ao SKU dos seus desenhos — o modelo e os desenhos ficam.
 *
 *   node servidor/sku-short-feminino-2910.js            (só mostra)
 *   node servidor/sku-short-feminino-2910.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const ANTIGO = 'SF.BERM.SEM';

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

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  // Fora do cadastro do modelo, o SKU antigo não pode estar em lugar nenhum:
  // se estiver, trocar só o modelo deixaria a outra ponta solta.
  const fora = Object.keys(data).filter(k => k !== 'modelos' && k !== '_device'
    && (typeof data[k] === 'string' ? data[k] : JSON.stringify(data[k] || '')).includes(ANTIGO));
  if (fora.length) throw new Error(ANTIGO + ' também aparece em: ' + fora.join(', ') + ' -- nada foi alterado.');
  const modelos = arr(data.modelos);
  const alvo = modelos.filter(m => String(m.skuLinha || '').trim().toUpperCase() === ANTIGO);
  if (alvo.length !== 1) throw new Error('esperava 1 modelo com ' + ANTIGO + ', achei ' + alvo.length);
  const m = alvo[0];
  const desenhos = arr(data.desenhos).filter(d => d.modeloId === m.id);
  // O SKU novo é o dos desenhos do modelo — tem de ser um só.
  const dos = [...new Set(desenhos.map(d => String(d.skuLinha || '').trim().toUpperCase()).filter(Boolean))];
  if (dos.length !== 1 || dos[0] === ANTIGO) throw new Error('os desenhos do modelo não têm um SKU único diferente de ' + ANTIGO + ': ' + dos.join(', '));
  const NOVO = dos[0];
  console.log('modelo   : "' + m.nome + '" (' + m.id + ')  ' + m.skuLinha + ' -> ' + NOVO);
  console.log('desenhos : ' + desenhos.map(d => d.codigo + ' ' + d.desc + ' [' + d.skuLinha + ']').join(' · '));
  m.skuLinha = NOVO;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-sku-short-feminino-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    modelos: typeof data.modelos === 'string' ? JSON.stringify(modelos) : modelos,
    _device: 'sku-short-feminino-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const m2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.modelos)
    .find(x => x.id === m.id);
  console.log('gravado: "' + m2.nome + '" agora é ' + m2.skuLinha);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
