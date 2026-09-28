/*
 * CORRIGE AS CORES DO FIO TP200 E DA LINHA BR02C PELA LISTA DA FABRICA.
 *
 * Junior, 28/09/2026: "Analise a imagem lista de cores fios e linhas diverse na
 * pasta download e corrija os cadastros de fios e linhas do estoque de
 * aviamentos". A lista (caderno, "Linha" e "Fio") substitui as cores que
 * tinham entrado so pelos codigos das etiquetas.
 *
 * LINHA: 14 cores, todas com codigo. O codigo vai com 6 digitos, como sai na
 * etiqueta do cone ("cor 000242"), e os tres que ja estavam (000242, 000105,
 * 000154) ganham nome: Bege, Preto, Roxo.
 * FIO: 12 cores, 6 com codigo e 6 so com o nome. Os codigos 1500 e 9900, que
 * vieram das etiquetas, nao estao na lista e saem.
 * Os nomes vao como no cadastro de cores da fabrica (Preto, Branco, Roxo...),
 * nao no feminino da lista ("Preta", "Branca"), para casar com o resto.
 *
 * Lancamento que use uma cor que sai e avisado (hoje nao ha nenhum).
 *
 *   node servidor/corrigir-cores-fio-linha.js            (so mostra)
 *   node servidor/corrigir-cores-fio-linha.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

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


const p6 = c => String(c).padStart(6, '0');
const CORES = {
  'Linha|BR02C': [
    ['701', 'Branco'], ['105', 'Preto'], ['285', 'Grafite'], ['154', 'Roxo'], ['2026', 'Vermelho'],
    ['242', 'Bege'], ['714', 'Verde'], ['130', 'Azul'], ['484', 'Off White'], ['264', 'Marrom'],
    ['572', 'Mostarda'], ['668', 'Amarelo'], ['680', 'Rosa'], ['757', 'Bege Moletom']
  ].map(([c, n]) => ({ codigo: p6(c), nome: n })),
  'Fio|TP200': [
    ['238', 'Verde'], ['', 'Grafite'], ['', 'Roxo'], ['', 'Vermelho'], ['526268', 'Bege'],
    ['523679', 'Azul Marinho'], ['632', 'Mostarda'], ['637', 'Marrom'], ['', 'Preto'], ['', 'Branco'],
    ['668', 'Amarelo'], ['', 'Rosa']
  ].map(([c, n]) => ({ codigo: c, nome: n }))
};
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const tipos = ler('aviamentoTipos'), movs = ler('aviamentosMov');
  const agora = new Date().toISOString();
  let n = 0;
  Object.entries(CORES).forEach(([k, cores]) => {
    const [item, codigo] = k.split('|');
    const t = tipos.find(x => normNome(x.item) === normNome(item) && normNome(x.codigo) === normNome(codigo));
    if (!t) { console.log('  NAO ACHEI o tipo ' + item + ' ' + codigo); return; }
    const antes = (t.cores || []).map(c => c.codigo + (c.nome ? '=' + c.nome : '')).join(', ');
    // Quem usa uma cor que sai (pelo codigo ou pelo nome)?
    const fica = c => cores.some(x => (x.codigo && normNome(x.codigo) === normNome(c)) || normNome(x.nome) === normNome(c));
    const usos = movs.filter(m => normNome(m.item) === normNome(item) && normNome(m.modelo) === normNome(codigo) && m.cor && !fica(m.cor));
    if (usos.length) console.log('  AVISO: ' + usos.length + ' lancamento(s) de ' + item + ' ' + codigo + ' usam cor que sai: ' + [...new Set(usos.map(m => m.cor))].join(', '));
    t.cores = cores; t.atualizadoPor = 'Junior'; t.atualizadoEm = agora;
    n++;
    console.log('  ' + item + ' ' + codigo + '\n     antes: ' + antes + '\n     agora: ' + cores.map(c => (c.codigo || '(sem codigo)') + '=' + c.nome).join(', '));
  });
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cores-fio-linha-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos) }), updatedAt);
  const t2 = JSON.parse((await lerBlob(cab)).data.aviamentoTipos);
  t2.forEach(t => console.log('gravado: ' + t.item + ' ' + t.codigo + ' com ' + t.cores.length + ' cores (' + t.cores.filter(c => !c.codigo).length + ' sem codigo)'));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
