/*
 * PAPEL KRAFT E FILME: DESCRICAO GENERICA, METROS, BAIXA PELAS OS E A CONTAGEM.
 *
 * Junior, 28/09/2026:
 *   - "As caracteristicas do cadastro de papel devem informar o basico sobre o
 *     material, pois os fornecedores podem mudar" -> sai o que e do fornecedor
 *     (marca, codigo do produto, CNPJ, EAN); fica gramatura, largura, rolo,
 *     espessura.
 *   - "papel e filme em metros" e a baixa automatica pelas OS (ver _matDasOS
 *     no app.js) -> medida 'm' e baixaOS ligada.
 *   - "Atualmente, existe 4 bobinas e meia de filme e 4 bobinas e meia de
 *     papel" -> kraft: 4,5 x 250 m = 1.125 m. Filme: a etiqueta da o PESO
 *     (25,8 kg), nao o comprimento; pelo tamanho da folha (2,60 m x 0,03 mm) e
 *     a densidade do polietileno (~920 kg/m3), 1 m pesa ~0,0718 kg e a bobina
 *     rende ~360 m -> 4,5 x 360 = 1.620 m. ESTIMATIVA: corrigir quando se
 *     souber o comprimento da bobina.
 *   - "Papel e filme plastico so devem constar no estoque de materiais da
 *     unidade descalvado" -> apaga copia que houver em Sao Carlos.
 *
 * A quantidade entra como AJUSTE no historico (como o "editar" do app), com o
 * que mexeu em estoque. Rodar de novo nao soma outra vez: so ajusta se a
 * contagem gravada for diferente da de agora.
 *
 *   node servidor/materiais-enfesto-metros.js            (so mostra)
 *   node servidor/materiais-enfesto-metros.js --gravar   (grava)
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

const FILME_KG_POR_M = 2.6 * 0.00003 * 920;          // ~0,0718 kg por metro
const FILME_M_POR_BOBINA = Math.round(25.8 / FILME_KG_POR_M);   // ~360 m
const MATERIAIS = [
  { nome: 'Papel kraft aerado',
    desc: 'Papel kraft aerado · 32 g/m² · rolo de 250 m × 1,60 m · 100% celulose',
    obs: '1 rolo = 250 m',
    emEstoque: 4.5 * 250 },
  { nome: 'Filme plástico',
    desc: 'Filme plástico em bobina · folha de 2,60 m de largura · 0,03 mm de espessura',
    obs: '1 bobina ≈ ' + FILME_M_POR_BOBINA + ' m (estimado pelo peso de 25,8 kg)',
    emEstoque: Math.round(4.5 * FILME_M_POR_BOBINA) }
];

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
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  let cad = ler('materiaisEstCad');
  const movs = ler('materiaisEstMov');
  const agora = new Date().toISOString();
  const hoje = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  let mexeu = 0;
  MATERIAIS.forEach((m, i) => {
    const doNome = cad.filter(c => normNome(c.nome) === normNome(m.nome));
    // So Descalvado: copia em Sao Carlos sai.
    doNome.filter(c => c.unidade === 'sc').forEach(c => {
      console.log('  apaga a copia de Sao Carlos: ' + c.nome);
      cad = cad.filter(x => x.id !== c.id); mexeu++;
    });
    const x = doNome.find(c => c.unidade !== 'sc');
    if (!x) { console.log('  NAO ACHEI em Descalvado: ' + m.nome); return; }
    const antes = Number(x.emEstoque) || 0;
    console.log('  ' + m.nome + ':\n      desc  "' + x.desc + '"\n         -> "' + m.desc + '"\n      obs   "' + (x.obs || '') + '" -> "' + m.obs + '"'
      + '\n      medida ' + (x.medida || 'un') + ' -> m, baixa pelas OS ' + (!!x.baixaOS) + ' -> true'
      + '\n      em estoque ' + antes + ' -> ' + m.emEstoque + ' m');
    x.desc = m.desc; x.obs = m.obs; x.medida = 'm'; x.baixaOS = true;
    x.atualizadoPor = 'Junior'; x.atualizadoEm = agora;
    if (antes !== m.emEstoque) {
      const d = Math.round((m.emEstoque - antes) * 100) / 100;
      x.emEstoque = m.emEstoque;
      movs.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
        itemId: x.id, nome: x.nome, desc: x.desc, unidade: 'desc', setor: x.setor || '', data: hoje,
        obs: 'Contagem: 4 bobinas e meia', tipo: 'ajuste', motivo: 'correcao', qtd: 0, dUso: 0, dEstoque: d,
        por: 'Junior', em: agora });
    }
    mexeu++;
  });
  if (!mexeu) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-materiais-enfesto-metros-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiaisEstCad: JSON.stringify(cad), materiaisEstMov: JSON.stringify(movs) }), updatedAt);

  const { data: d2 } = await lerBlob(cab);
  const c2 = JSON.parse(d2.materiaisEstCad);
  c2.filter(c => MATERIAIS.some(m => normNome(m.nome) === normNome(c.nome))).forEach(c =>
    console.log('gravado: ' + c.nome + ' [' + (c.unidade || 'desc') + '] ' + c.emEstoque + ' ' + c.medida + ', baixa pelas OS: ' + c.baixaOS));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
