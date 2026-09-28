/*
 * AGULHAS: CADASTRO EM CADASTROS > PECAS E UMA DE CADA TIPO EM USO.
 *
 * Junior, 28/09/2026:
 *   - "insira Pecas na barra lateral, na pasta Cadastros. Depois, cadastre em
 *     Pecas todos os tipos de agulhas" -> os 11 tipos que estao no Estoque de
 *     pecas viram cadastros da categoria 'peca' em STATE.materiais, com o
 *     proximo codigo livre de 3 digitos. Descricao = o tipo ("Agulha DBx1
 *     no 11"); Tipo = "Agulha" + a maquina.
 *   - "Insira na coluna em uso do estoque pecas, uma agulha de cada tipo" ->
 *     em cada linha de agulha do Estoque de pecas (as duas unidades), 1 sai do
 *     estoque e entra em uso, pelo lancamento "Posta em uso" do app.
 * Rodar de novo nao repete nenhuma das duas coisas.
 *
 *   node servidor/agulhas-cadastro-e-em-uso.js            (so mostra)
 *   node servidor/agulhas-cadastro-e-em-uso.js --gravar   (grava)
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


const OBS = 'Uma de cada tipo na maquina';
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const mats = ler('materiais'), pecas = ler('pecasCad'), movs = ler('pecasMov');
  const agora = new Date().toISOString();
  const hoje = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  const agulhas = pecas.filter(p => /^agulha/i.test(p.nome || ''));

  // 1) o cadastro: um por tipo (o nome), pela linha de Descalvado
  let prox = Math.max(0, ...mats.map(m => parseInt(String(m.codigo || '').replace(/\D/g, ''), 10) || 0)) + 1;
  let nCad = 0;
  const tipos = [...new Map(agulhas.filter(a => a.unidade !== 'sc').map(a => [normNome(a.nome), a])).values()];
  tipos.forEach((a, i) => {
    if (mats.some(m => normNome(m.desc) === normNome(a.nome))) { console.log('  ja cadastrada: ' + a.nome); return; }
    const maq = ((String(a.desc || '').match(/máquina:\s*(.+)$/i) || [])[1] || '').trim();
    const codigo = String(prox++).padStart(3, '0');
    mats.push({ id: 'id_' + Date.now() + '_c' + i + String(Math.floor(Math.random() * 1000)),
      codigo, desc: a.nome, tipo: 'Agulha' + (maq ? ' · ' + maq : ''), categoria: 'peca' });
    nCad++;
    console.log('  cadastro ' + codigo + ' · ' + a.nome + ' (Agulha' + (maq ? ' · ' + maq : '') + ')');
  });

  // 2) uma de cada em uso
  let nUso = 0;
  agulhas.forEach((x, i) => {
    if (movs.some(m => m.itemId === x.id && m.obs === OBS)) { console.log('  ja em uso: ' + x.nome + ' [' + (x.unidade || 'desc') + ']'); return; }
    const est = Math.round(Number(x.emEstoque) || 0), uso = Math.round(Number(x.emUso) || 0);
    if (est < 1) { console.log('  sem estoque: ' + x.nome + ' [' + (x.unidade || 'desc') + ']'); return; }
    x.emEstoque = est - 1; x.emUso = uso + 1;
    x.atualizadoPor = 'Junior'; x.atualizadoEm = agora;
    movs.push({ id: 'id_' + Date.now() + '_u' + i + String(Math.floor(Math.random() * 1000)),
      itemId: x.id, nome: x.nome, desc: x.desc || '', unidade: x.unidade === 'sc' ? 'sc' : 'desc', data: hoje,
      obs: OBS, tipo: 'saida', motivo: 'uso', qtd: 1, dUso: 1, dEstoque: -1, por: 'Junior', em: agora });
    nUso++;
  });
  console.log('\n' + nCad + ' cadastro(s) em Pecas; ' + nUso + ' linha(s) do estoque com 1 em uso.');
  if (!nCad && !nUso) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-agulhas-cadastro-uso-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiais: JSON.stringify(mats), pecasCad: JSON.stringify(pecas), pecasMov: JSON.stringify(movs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const m2 = JSON.parse(d2.materiais), p2 = JSON.parse(d2.pecasCad);
  const ag2 = p2.filter(p => /^agulha/i.test(p.nome || ''));
  console.log('gravado. Pecas no cadastro: ' + m2.filter(m => m.categoria === 'peca').length
    + ' · agulhas no estoque: ' + ag2.length + ' linhas, em uso ' + ag2.reduce((a, p) => a + (p.emUso || 0), 0)
    + ', em estoque ' + ag2.reduce((a, p) => a + (p.emEstoque || 0), 0));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
