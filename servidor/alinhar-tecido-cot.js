/*
 * Poe cada fase de uma grade COT no tecido que o NOME dela declara.
 *
 * POR QUE ISTO EXISTE
 *
 * A familia Texturizado tem tres panos — Jaguar, Prime e Rugao — e a grade diz
 * qual no proprio nome: COT.JAC, COT.PRI, COT.RUG. Nas grades PRI e RUG todas
 * as fases estao no pano da linha delas; nas JAC, nao: uma tinha o Vies em
 * Rugao e a outra sem tecido nenhum.
 *
 * A origem foi uma copia: a `P-M-G-GG-G1 | COT.JAC | 157cm` nasceu da grade
 * COT.RUG (as medidas coincidem de verdade — os dois riscos dao 5,1006 m x
 * 1,57 m, porque as duas bobinas sao de 157 cm) e o tecido nao foi trocado. O
 * Junior corrigiu o "Corpo + Gola" das duas a mao; o Vies ficou.
 *
 * POR QUE VALE A REGRA
 *
 * Nas COT o vies e cortado da SOBRA do mesmo enfesto do corpo — e por isso usa
 * o mesmo pano. Nao e regra geral do programa: numa blusa de moletom o vies e
 * de Malha Algodao enquanto o corpo e Moletom, e ali as fases usam panos
 * diferentes de proposito. Por isso este script olha SO as grades COT.
 *
 * Fase sem tecido tambem e preenchida: fase sem pano reserva zero, e zero num
 * campo de pano principal e o tipo de erro que ninguem ve ate faltar tecido.
 *
 * COMO RODAR
 *
 *   node servidor/alinhar-tecido-cot.js             so relata, nao grava
 *   node servidor/alinhar-tecido-cot.js --gravar    grava no servidor
 *
 * Antes de gravar salva o blob em backups/, e a escrita confere o carimbo.
 *
 * ATENCAO: recarregue a pagina (F5) depois de rodar. Aba aberta desde antes
 * regrava a chave inteira com o estado velho e desfaz isto — aconteceu tres
 * vezes em 28/08 (ver project_restauracao_credenciada).
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

// A sigla no nome da grade e o tecido que ela pede. O nome do tecido nao entra
// aqui inteiro: e resolvido no cadastro, para o script nao ter a sua propria
// ideia de como o pano se chama.
const LINHAS = [
  { sigla: 'COT.JAC', casaTecido: /jaguar/i },
  { sigla: 'COT.PRI', casaTecido: /prime/i },
  { sigla: 'COT.RUG', casaTecido: /rug/i }
];

async function lerBlob() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at, cab };
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
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
}

(async () => {
  const { data, updatedAt, cab } = await lerBlob();
  const le = k => JSON.parse(data[k] || '[]');
  const grades = le('grades'), tecidos = le('tecidos'), ordens = le('ordens');
  const tN = id => (tecidos.find(t => t.id === id) || {}).nome || '(sem tecido)';

  let mudou = 0;
  grades.forEach(g => {
    const linha = LINHAS.find(l => String(g.nome || '').toUpperCase().includes(l.sigla));
    if (!linha) return;
    // Um tecido, e só um, casa a linha da grade. Dois casando seria cadastro
    // ambíguo, e escolher no escuro poe pano errado na conta.
    const alvos = tecidos.filter(t => /textur/i.test(t.nome) && linha.casaTecido.test(t.nome));
    if (alvos.length !== 1) {
      console.log('  PULADA ' + g.nome + ': ' + alvos.length + ' tecidos casam "' + linha.sigla + '"');
      return;
    }
    const alvo = alvos[0];
    const trocas = (g.fases || []).filter(f => f.tecidoId !== alvo.id);
    if (!trocas.length) return;
    const nOS = ordens.filter(o => o.gradeId === g.id).length;
    console.log('  ' + g.nome + '   (' + nOS + ' OS emitidas)');
    trocas.forEach(f => {
      console.log('     fase "' + f.nome + '": ' + tN(f.tecidoId) + '  ->  ' + alvo.nome);
      f.tecidoId = alvo.id;
      mudou++;
    });
  });

  console.log('');
  console.log('fases alinhadas: ' + mudou);
  if (!mudou) { console.log('Nada a corrigir.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-tecido-cot-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.grades = JSON.stringify(grades);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
