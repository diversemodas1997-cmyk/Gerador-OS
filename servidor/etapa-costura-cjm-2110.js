/* ETAPA "Costura CLM | Descalvado" → "Costura CJM | Descalvado" (21/10/2026,
   Junior: "Renomeie a etapa para Costura CJM também"), junto da troca do SKU da
   calça moletom (sku-clm-para-cjm-2110.js). O nome da etapa é a chave com que o
   desenho e a OS a guardam, então troca nos três lugares onde ele está: o
   cadastro de etapas, o desenho 0037 (etapasNomes) e a OS 0589 (etapas, e o
   progresso, se tiver a etapa marcada).

   Rode:  node servidor/etapa-costura-cjm-2110.js            (só mostra)
          node servidor/etapa-costura-cjm-2110.js --gravar   (grava, com backup) */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const VELHO = 'Costura CLM | Descalvado';
const NOVO = 'Costura CJM | Descalvado';
const CHAVES = ['etapas', 'desenhos', 'ordens'];

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
const texto = v => (typeof v === 'string' ? v : JSON.stringify(v));

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  // O nome só pode estar nas três chaves de sempre — em outra, é algo que este
  // script não conhece, e é melhor parar.
  Object.keys(data).filter(k => !CHAVES.includes(k) && texto(data[k]).includes(VELHO))
    .forEach(k => { throw new Error('o nome aparece também em ' + k); });
  if (Object.keys(data).some(k => texto(data[k]).includes(NOVO))) throw new Error('já existe "' + NOVO + '"');
  const json = JSON.stringify(VELHO).slice(1, -1), jsonNovo = JSON.stringify(NOVO).slice(1, -1);
  const novo = Object.assign({}, data);
  CHAVES.forEach(k => {
    const s = texto(data[k]);
    const n = s.split(json).length - 1;
    console.log(`${k}: ${n} ocorrência(s)`);
    if (!n) return;
    const trocado = s.split(json).join(jsonNovo);
    novo[k] = typeof data[k] === 'string' ? trocado : JSON.parse(trocado);
  });
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-etapa-costura-cjm-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  novo._device = 'etapa-costura-cjm-' + Date.now();
  const r = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!r.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const conf = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const resta = Object.keys(conf).filter(k => texto(conf[k]).includes(VELHO));
  console.log(resta.length ? 'AINDA COM O NOME VELHO: ' + resta.join(', ') : 'gravado: nenhum "' + VELHO + '" sobrou');
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
