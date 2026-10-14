/*
 * FIOS E LINHAS DE SÃO CARLOS = A CONTAGEM DE 14/10/2026.
 *
 * Junior, 14/10/2026: "ajuste a quantidade em estoque de linha unidade são
 * carlos para 4 linhas Bege, 13 linhas e 5 fios branco, 25 linhas e 12 fios
 * preto, 15 linhas e 11 fios amarelo, 4 linhas roxo, 2 linhas grafite, 2
 * linhas e 2 fios marrom, 4 linhas off-white, 18 linhas e 13 fios rosa" — e é
 * a contagem COMPLETA: o que não está na lista fica em zero.
 *
 * Para cada linha do estoque de São Carlos (item + modelo + cor + código), um
 * lançamento de ENTRADA ou SAÍDA com a diferença entre o saldo de hoje e o
 * contado. Cor contada que ainda não existia lá entra com o modelo e o código
 * que a mesma cor tem em Descalvado ou, sem ela lá, no cadastro de fios e
 * linhas (o Fio Amarelo entrou primeiro sem o código 668 e foi corrigido).
 *
 * O saldo de hoje é a soma dos lançamentos gravados (entradas, saídas e as
 * pernas das OE). Fio e linha de São Carlos não têm baixa automática de OS
 * hoje (conferido antes de rodar): se um dia tiverem, este script não vale mais.
 *
 *   node servidor/fios-linhas-sc-1410.js            (so mostra)
 *   node servidor/fios-linhas-sc-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const HOJE = '2026-10-14';
const OBS = 'Contagem 14/10/2026 São Carlos';
const MODELO = { Linha: 'BR02C', Fio: 'TP200' };
const CONTADO = {
  Linha: { 'Bege': 4, 'Branco': 13, 'Preto': 25, 'Amarelo': 15, 'Roxo': 4, 'Grafite': 2, 'Marrom': 2, 'Off White': 4, 'Rosa': 18 },
  Fio: { 'Branco': 5, 'Preto': 12, 'Amarelo': 11, 'Marrom': 2, 'Rosa': 13 }
};

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
const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[\s-]+/g, ' ').trim();

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const movs = arr(data.aviamentosMov);
  if (movs.some(m => String(m.obs || '') === OBS)) { console.log('A contagem de 14/10 já foi lançada. Nada a fazer.'); return; }

  // Saldo por linha, numa unidade.
  const saldo = unidade => {
    const linhas = new Map();
    movs.forEach(m => {
      if (m.item !== 'Fio' && m.item !== 'Linha') return;
      const pernas = m.tipo === 'expedicao'
        ? [{ tipo: 'saida', unidade: m.perna === 'volta' ? 'sc' : 'desc', data: m.dataSaida || m.data },
           { tipo: 'entrada', unidade: m.perna === 'volta' ? 'desc' : 'sc', data: m.data }]
        : [{ tipo: m.tipo, unidade: m.unidade === 'sc' ? 'sc' : 'desc', data: m.data }];
      pernas.forEach(p => {
        if (p.unidade !== unidade || String(p.data || HOJE) > HOJE) return;
        const k = [m.item, norm(m.modelo), norm(m.cor), String(m.corCodigo || '').trim()].join('|');
        const cur = linhas.get(k) || { item: m.item, modelo: m.modelo || '', cor: m.cor, corCodigo: String(m.corCodigo || '').trim(), un: 0 };
        cur.un += (p.tipo === 'saida' ? -1 : 1) * (Number(m.qtd) || 0);
        linhas.set(k, cur);
      });
    });
    return [...linhas.values()];
  };
  const sc = saldo('sc'), desc = saldo('desc');

  const novos = [];
  const agora = new Date().toISOString();
  const lancar = (l, delta) => {
    if (!delta) return;
    novos.push({ id: 'contagem-sc-1410:' + l.item + ':' + norm(l.cor) + ':' + (l.corCodigo || '-'), tipo: delta > 0 ? 'entrada' : 'saida',
      unidade: 'sc', item: l.item, modelo: l.modelo, tam: '', cor: l.cor, corCodigo: l.corCodigo, kg: 0, qtd: Math.abs(delta),
      data: HOJE, obs: OBS, por: 'Junior', em: agora });
  };
  const usados = new Set();
  console.log('São Carlos — hoje → contado:');
  sc.forEach(l => {
    const alvoCor = Object.keys(CONTADO[l.item]).find(c => norm(c) === norm(l.cor));
    // Mais de uma linha com a mesma cor: o contado vai para a primeira, as outras zeram.
    const alvo = alvoCor && !usados.has(l.item + alvoCor) ? CONTADO[l.item][alvoCor] : 0;
    if (alvoCor) usados.add(l.item + alvoCor);
    console.log(`  ${l.item.padEnd(5)} ${l.cor.padEnd(14)} ${String(l.corCodigo || '-').padEnd(7)} ${String(l.un).padStart(4)} → ${alvo}`);
    lancar(l, alvo - l.un);
  });
  Object.entries(CONTADO).forEach(([item, cores]) => Object.entries(cores).forEach(([cor, qtd]) => {
    if (usados.has(item + cor)) return;
    const ref = desc.find(l => l.item === item && norm(l.cor) === norm(cor))
      || arr(data.aviamentoTipos).map(t => ({ item: t.item, modelo: t.codigo, cor: t.corNome, corCodigo: String(t.corCodigo || '').trim() }))
        .find(t => t.item === item && norm(t.cor) === norm(cor));
    const l = { item, modelo: (ref && ref.modelo) || MODELO[item], cor, corCodigo: (ref && ref.corCodigo) || '' };
    console.log(`  ${item.padEnd(5)} ${cor.padEnd(14)} ${String(l.corCodigo || '-').padEnd(7)}  nova → ${qtd}`);
    lancar(l, qtd);
  }));
  console.log(`\n${novos.length} lançamentos (${novos.filter(m => m.tipo === 'entrada').length} entradas, ${novos.filter(m => m.tipo === 'saida').length} saídas).`);
  if (!GRAVAR) { console.log('Nada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-fios-linhas-sc-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('backup: ' + path.relative(RAIZ, arq));
  const saida = movs.concat(novos);
  const novo = Object.assign({}, data, {
    aviamentosMov: typeof data.aviamentosMov === 'string' ? JSON.stringify(saida) : saida,
    _device: 'fios-linhas-sc-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  console.log('gravado. Agora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
