/*
 * Reaponta `fase.risco` para o PDF que MUDOU DE PASTA — sem trocar de PDF.
 *
 * POR QUE ISTO EXISTE
 *
 * Em 04/09/2026, cinco grades da CM.TRI tiveram a pasta de largura renomeada de
 * "116.5 cm" para "116.5 cm - MAPAS IMPRESSOS". Os PDFs não se moveram e nem
 * mudaram de nome: só o endereço deles mudou. O cadastro, porém, guarda em
 * `fase.risco` o caminho COMPLETO de quando a medida foi importada — e ele
 * continuou apontando para a pasta antiga.
 *
 * O estrago é o de 01/09/2026 com a OS 0508, e é caro justamente por ser mudo:
 * a coluna Riscos da grade fica vazia, não há de onde reimportar a medida, e
 * dias depois uma OS nova diz "a medida desta grade não veio de um risco" — o
 * aviso está certo e parece mentira.
 *
 * A REGRA: RESOLVE PELO NOME DO ARQUIVO, E SÓ QUANDO NÃO HÁ DÚVIDA
 *
 * Não há aqui nenhuma regra de "troque 116.5 cm por 116.5 cm - MAPAS
 * IMPRESSOS". Isso consertaria hoje e não serviria da próxima vez. O que ele
 * faz é: para cada risco que não resolve mais, pega o NOME do arquivo e procura
 * na lista de PDFs (dados/riscos-pdf.json).
 *
 *   · achou EXATAMENTE UM  -> reaponta para o caminho novo;
 *   · achou mais de um     -> NÃO ESCOLHE. Relata e deixa para uma pessoa.
 *   · não achou nenhum     -> relata. O PDF sumiu de verdade, e isso é outro
 *                             problema, que um reaponte esconderia.
 *
 * Adivinhar aqui custa tecido: risco errado numa grade vira enfesto errado.
 * Por isso o empate NUNCA é desempatado por conta própria.
 *
 * A comparação de "resolve ou não" é a MESMA do app (sufixo), porque
 * `fase.risco` às vezes guarda o caminho inteiro e às vezes só o nome do
 * arquivo — ver a nota de _riscoRelResolvido no app.js. Comparar por igualdade
 * acusaria metade do cadastro como quebrada.
 *
 * COMO RODAR
 *
 *   node servidor/corrigir-risco-de-pasta-renomeada.js             só relata
 *   node servidor/corrigir-risco-de-pasta-renomeada.js --gravar    grava
 *
 * Antes de gravar ele salva o blob inteiro em backups/, e a escrita é
 * read-modify-write com CONFERÊNCIA DE CARIMBO: se alguém gravou no servidor
 * entre a leitura e a escrita, aborta em vez de passar por cima (o blob é uma
 * linha só, compartilhada — ver project_restauracao_credenciada).
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const norm = s => String(s || '').split('\\').join('/').trim();
const base = s => norm(s).split('/').pop();

// A mesma regra do app: sufixo, nos dois sentidos.
const casa = (lista, r) => lista.some(e => e === r || e.endsWith('/' + r) || r.endsWith('/' + e));

/* ---- o servidor da fabrica ---- */
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
  const idx = JSON.parse(fs.readFileSync(path.join(RAIZ, 'dados', 'riscos-pdf.json'), 'utf8'))
    .arquivos.map(norm);
  // nome do arquivo -> caminhos que terminam nele
  const porNome = new Map();
  idx.forEach(e => {
    const n = base(e);
    if (!porNome.has(n)) porNome.set(n, []);
    porNome.get(n).push(e);
  });

  const { data, updatedAt, cab } = await lerBlob();
  const grades = typeof data.grades === 'string' ? JSON.parse(data.grades) : data.grades;

  const arrumados = [], empatados = [], sumidos = [];
  for (const g of grades) {
    for (const f of (g.fases || [])) {
      const r = norm(f.risco);
      if (!r || casa(idx, r)) continue;
      let cand = porNome.get(base(r)) || [];
      /* EMPATE DESFEITO PELO QUE NAO MUDOU.
         O mesmo nome de arquivo aparece em mais de uma pasta de grade — a
         ribana v1 do P ao G3 existe na pasta P-M-G-GG-G1-G2-G3 e na
         P-M-2G-GG-G1-G2-G3, e sao encaixes diferentes. Quando o risco antigo
         era um caminho, a pasta da GRADE nele nao mudou (o que mudou foi a de
         largura, embaixo dela): entao dela sai o desempate, e isso nao e
         chutar — e usar a parte da informacao que continua valendo.
         Sobrando mais de um mesmo assim, continua sem escolha. */
      if (cand.length > 1 && r.includes('/')) {
        const pastaDaGrade = norm(r).split('/').slice(0, 2).join('/') + '/';
        const mesmos = cand.filter(c => c.startsWith(pastaDaGrade));
        if (mesmos.length === 1) cand = mesmos;
      }
      if (cand.length === 1) { arrumados.push({ g, f, de: r, para: cand[0] }); }
      else if (cand.length > 1) { empatados.push({ g, f, de: r, cand }); }
      else { sumidos.push({ g, f, de: r }); }
    }
  }

  console.log(`grades: ${grades.length} · PDFs na lista: ${idx.length}\n`);

  console.log(`=== REAPONTADOS — mesmo arquivo, endereco novo (${arrumados.length}) ===`);
  arrumados.forEach(a => console.log(`  ${a.g.nome}  [${a.f.nome}]\n      de:   ${a.de}\n      para: ${a.para}`));

  if (empatados.length) {
    console.log(`\n=== EMPATE, NAO ESCOLHI (${empatados.length}) ===`);
    empatados.forEach(e => {
      console.log(`  ${e.g.nome}  [${e.f.nome}]  -> ${e.de}`);
      e.cand.forEach(c => console.log(`        candidato: ${c}`));
    });
  }
  if (sumidos.length) {
    console.log(`\n=== O PDF NAO EXISTE MAIS EM LUGAR NENHUM (${sumidos.length}) ===`);
    sumidos.forEach(s => console.log(`  ${s.g.nome}  [${s.f.nome}]  -> ${s.de}`));
  }

  if (!arrumados.length) { console.log('\nNada a reapontar.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  arrumados.forEach(a => { a.f.risco = a.para; });
  data.grades = JSON.stringify(grades);

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-risco-pasta-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + arrumados.length + ' fases reapontadas.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
