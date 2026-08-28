/*
 * Junta a fase de ribana DUPLICADA da grade: leva a medida para a fase que ja
 * existia e apaga a copia.
 *
 * POR QUE ISTO EXISTE
 *
 * A importacao de risco de ribana criou fase NOVA em vez de reconhecer a que ja
 * estava na grade, porque o programa normaliza o nome mas nao reordena as
 * palavras: "Barra/Punhos" vira `barra punhos` e "Punhos/Barra" vira
 * `punhos barra` — para ele, duas fases diferentes.
 *
 * O estrago e na PROXIMA OS, nao nas emitidas: a OS copia as fases da grade no
 * momento em que e emitida. Medido numa OS nova da 2M-4G-2GG | BM.TRI | 177.5cm
 * (antes de o Junior conserta-la a mao): a ribana saia 85,4 kg em vez de 42,7,
 * e a duplicata ainda entrava no planejamento com 9 operacoes proprias, um dia
 * de posto para um enfesto que nao existe.
 *
 * O QUE ELE FAZ, E EM QUE DIRECAO
 *
 * Nas tres grades que sobraram em 28/08 a duplicata e a UNICA que tem a medida
 * — a original esta em branco. Apagar a copia, como se fez nas duas primeiras,
 * jogaria fora o unico encaixe medido e a ribana passaria a reservar zero.
 * Entao a direcao aqui e a inversa: a copia ENTREGA o que tem e sai.
 *
 * Ele NUNCA sobrescreve um valor que a fase original ja tenha. So preenche o
 * que esta vazio — medida, tecido, bobinas, unidades, excedente e o `risco` (o
 * PDF que provou a medida, ver project_coluna_riscos_grade). Onde as duas
 * responderam, quem manda e a original: ela e a que as OS ja emitidas conhecem.
 *
 * COMO ELE DECIDE QUE SAO A MESMA FASE
 *
 * Mesmas palavras no nome, em qualquer ordem (`barra punhos` == `punhos barra`),
 * E o mesmo tecido — ou uma das duas sem tecido. Sem a trava do tecido, duas
 * fases de nomes parecidos e panos diferentes seriam fundidas.
 *
 * Casos de nome DIFERENTE ("Gola" e "Ribana" na camiseta) ele nao pega, de
 * proposito: ali so quem conhece a peca sabe que sao a mesma coisa.
 *
 * COMO RODAR
 *
 *   node servidor/juntar-fase-duplicada.js             so relata, nao grava
 *   node servidor/juntar-fase-duplicada.js --gravar    grava no servidor
 *
 * Antes de gravar ele salva o blob inteiro em backups/, e a escrita e
 * read-modify-write com CONFERENCIA DE CARIMBO: se alguem gravou no servidor
 * entre a leitura e a escrita, aborta em vez de passar por cima.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

// Os campos que a copia pode entregar. `nome` e `ordem` ficam de fora: o nome
// certo e o da original (e o que as OS emitidas carregam), e a ordem dela e a
// posicao na corrente do enfesto.
const CAMPOS = ['comp', 'larg', 'tecidoId', 'bobinas', 'unidades', 'excedente', 'risco',
                'enfestoMin', 'corteMin'];

const vazio = v => v == null || String(v).trim() === '';
// Mesmas palavras, em qualquer ordem — e a unica diferenca entre "Barra/Punhos"
// e "Punhos/Barra".
const chaveNome = s => String(s || '').toLowerCase().normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '').replace(/[+/&,;-]+/g, ' ')
  .split(/\s+/).filter(Boolean).sort().join(' ');

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
  const grades = le('grades');
  const ordens = le('ordens');
  const tecidos = le('tecidos');
  const tN = id => (tecidos.find(t => t.id === id) || {}).nome || '(sem tecido)';

  let juntadas = 0, recusadas = 0;
  grades.forEach(g => {
    const fases = g.fases || [];
    const porChave = new Map();
    const aRemover = [];
    fases.forEach(f => {
      const k = chaveNome(f.nome);
      if (!k) return;
      const orig = porChave.get(k);
      if (!orig) { porChave.set(k, f); return; }

      // Mesma fase? So se o pano bate — ou se uma das duas ainda nao respondeu.
      const mesmoTecido = vazio(orig.tecidoId) || vazio(f.tecidoId) || orig.tecidoId === f.tecidoId;
      if (!mesmoTecido) {
        recusadas++;
        console.log('  RECUSADA (panos diferentes): ' + g.nome
          + ' -> "' + orig.nome + '" (' + tN(orig.tecidoId) + ') e "' + f.nome + '" (' + tN(f.tecidoId) + ')');
        return;
      }

      const levou = [];
      CAMPOS.forEach(c => {
        if (vazio(orig[c]) && !vazio(f[c])) { orig[c] = f[c]; levou.push(c); }
      });
      const perdidos = CAMPOS.filter(c => !vazio(orig[c]) && !vazio(f[c])
        && String(orig[c]) !== String(f[c]));
      aRemover.push({ f, orig, levou, perdidos });
    });

    if (!aRemover.length) return;

    // Apagar so pode nao deixar buraco na ordem: a `ordem` casa a fase da grade
    // com a fase que a OS copiou, e renumerar aqui bagunçaria as OS emitidas.
    const restantes = fases.filter(x => !aRemover.some(r => r.f === x));
    const ords = restantes.map(x => Number(x.ordem)).sort((a, b) => a - b);
    const contiguo = ords.every((v, i) => v === i + 1);
    if (!contiguo) {
      recusadas += aRemover.length;
      console.log('  RECUSADA (apagar deixaria buraco na ordem): ' + g.nome
        + ' -> ordens que sobrariam: ' + ords.join(','));
      return;
    }

    const osDaGrade = ordens.filter(o => o.gradeId === g.id);
    console.log('');
    console.log('  ' + g.nome + '   (' + osDaGrade.length + ' OS emitidas)');
    aRemover.forEach(r => {
      console.log('     "' + r.f.nome + '"  ->  "' + r.orig.nome + '"');
      r.levou.forEach(c => console.log('        leva ' + c + ': ' + JSON.stringify(r.orig[c])));
      if (!r.levou.length) console.log('        (nada a levar — a original ja tinha tudo)');
      r.perdidos.forEach(c => console.log('        MANTIDO o da original em ' + c
        + ': ' + JSON.stringify(r.orig[c]) + '  (a copia dizia ' + JSON.stringify(r.f[c]) + ')'));
      console.log('        apaga a copia (ordem ' + r.f.ordem + ')');
      juntadas++;
    });
    g.fases = restantes;
  });

  console.log('');
  console.log('fases juntadas: ' + juntadas + (recusadas ? '   |   recusadas: ' + recusadas : ''));
  if (!juntadas) { console.log('Nada a fazer.'); return; }

  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-juntar-fase-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.grades = JSON.stringify(grades);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
