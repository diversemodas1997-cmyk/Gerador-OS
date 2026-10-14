/*
 * ETIQUETAS DE SÃO CARLOS: O QUE JÁ HAVIA LÁ ANTES DE 25/09.
 *
 * Junior, 14/10/2026: "Não foram dadas as entradas das quantidades de
 * etiquetas que já havia na unidade antes da entrada do dia 25. Cadastre as
 * entradas para que atualmente tenha etiquetas tamanho P 2000, M 6000,
 * G 6000, GG 6000, G1 9400, G2 9600, G3 8000."
 *
 * O saldo de hoje em São Carlos, lido do programa às 16:40 de 14/10 (entrada
 * da OE de 25/09, baixas das OS, ajuste das 23 OS), era:
 *   P 3192 · M 936 · G −1924 · GG 604 · G1 1290 · G2 1754 · G3 1626
 * A diferença para o pedido entra como ENTRADA em 24/09 — o estoque que já
 * estava na prateleira antes da caixa de 25/09. O P é a exceção: o pedido
 * (2000) é MENOR do que o saldo, e uma entrada não diminui nada; ele recebe uma
 * SAÍDA de ajuste em 14/10 com a diferença.
 *
 * Não repete: lançamento já gravado com o mesmo id fica como está.
 *
 *   node servidor/etiquetas-sc-saldo-inicial-1410.js            (so mostra)
 *   node servidor/etiquetas-sc-saldo-inicial-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();

const SALDO_HOJE = { P: 3192, M: 936, G: -1924, GG: 604, G1: 1290, G2: 1754, G3: 1626 };
const PEDIDO = { P: 2000, M: 6000, G: 6000, GG: 6000, G1: 9400, G2: 9600, G3: 8000 };

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
  const movs = arr(data.aviamentosMov);
  const agora = new Date().toISOString();
  const novos = Object.keys(PEDIDO).map(tam => {
    const d = PEDIDO[tam] - SALDO_HOJE[tam];
    if (!d) return null;
    const entrada = d > 0;
    return {
      id: 'saldo-inicial-sc-1410:Etiqueta:' + tam, tipo: entrada ? 'entrada' : 'saida', unidade: 'sc',
      item: 'Etiqueta', tam, cor: 'Preto', kg: 0, qtd: Math.abs(d), data: entrada ? '2026-09-24' : '2026-10-14',
      obs: entrada ? 'Estoque que já havia em São Carlos antes da OE de 25/09 (lançado em 14/10/2026)'
                   : 'Ajuste para o saldo informado em 14/10/2026',
      por: 'Junior', em: agora
    };
  }).filter(Boolean).filter(m => !movs.some(y => y.id === m.id));
  novos.forEach(m => console.log(`  ${m.tam.padEnd(3)} ${m.tipo.padEnd(7)} ${String(m.qtd).padStart(5)}  em ${m.data}  → fica ${PEDIDO[m.tam]}`));
  if (!novos.length) { console.log('Nada a fazer.'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-etiquetas-sc-saldo-inicial-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('backup: ' + path.relative(RAIZ, arq));
  const saida = movs.concat(novos);
  const novo = Object.assign({}, data, {
    aviamentosMov: typeof data.aviamentosMov === 'string' ? JSON.stringify(saida) : saida,
    _device: 'etiquetas-sc-saldo-inicial-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  console.log('gravado. Agora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
