/*
 * Move o texto da CAIXA ANTIGA de observacao (`os.obs`, uma por OS, sem dono)
 * para uma NOTA ASSINADA (`os.obsNotas`), no login que for passado.
 *
 * POR QUE
 *
 * Ate 20/08/2026 a folha tinha uma caixa de observacao so, sem autor: so o
 * admin escrevia, entao a pergunta "de quem e este recado?" nunca aparecia.
 * Nesse dia a folha passou a ter UMA OBSERVACAO POR PESSOA, e o texto velho
 * ficou de fora, aparecendo como um bloco a parte assinado "admin".
 *
 * Junior confirmou que TODAS as anotacoes que ja existem, em todas as OS, sao
 * dele - e pediu para carimba-las com o login dele. Este script faz isso: o
 * texto vira uma nota como qualquer outra, e o bloco a parte some da folha.
 *
 * O QUE ELE NAO INVENTA
 *
 * A caixa antiga nao guardava HORA. A nota migrada nasce com `anterior: true` e
 * sem `em`, e a folha mostra "anterior a 20/08/2026" no lugar da data. Carimbar
 * um dia que ninguem sabe seria transformar suposicao em registro.
 *
 * COMO RODAR
 *
 *   node servidor/migrar-obs-para-nota.js <login>            so relata
 *   node servidor/migrar-obs-para-nota.js <login> --gravar   grava no servidor
 *
 * A gravacao e read-modify-write com CONFERENCIA DE CARIMBO: se alguem gravar
 * no servidor entre a leitura e a escrita, o PATCH nao acha a linha e o script
 * aborta em vez de passar por cima do blob de todo mundo.
 */
const fs = require('fs');
const path = require('path');

const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';

const args = process.argv.slice(2);
const GRAVAR = args.includes('--gravar');
const LOGIN = (args.find(a => !a.startsWith('--')) || '').trim().toLowerCase();

if (!LOGIN) {
  console.error('Falta o login. Ex.: node servidor/migrar-obs-para-nota.js fulano@empresa.com');
  process.exit(1);
}

async function servidor() {
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

(async () => {
  const { data, updatedAt, cab } = await servidor();
  const ordens = JSON.parse(data.ordens || '[]');

  const comTexto = ordens.filter(o => String(o.obs || '').trim());
  console.log(`OS no total: ${ordens.length}`);
  console.log(`com texto na caixa antiga: ${comTexto.length}`);
  console.log(`destino: ${LOGIN}`);
  console.log('');

  let migradas = 0, jaTinha = 0;
  comTexto.forEach(o => {
    const texto = String(o.obs).trim();
    const notas = Array.isArray(o.obsNotas) ? o.obsNotas : [];
    // Ja existe nota DESTE login nesta OS? Nao junta os dois textos por conta
    // propria - o que a pessoa escreveu depois e mais recente, e colar o velho
    // em cima ou embaixo e decisao dela, nao do script.
    if (notas.some(n => String(n.login || '').trim().toLowerCase() === LOGIN)) {
      jaTinha++;
      console.log(`  ! OS ${o.os || o.id}: ja tem nota deste login — DEIXADA COMO ESTA`);
      console.log(`      caixa antiga: ${JSON.stringify(texto.slice(0, 70))}`);
      return;
    }
    migradas++;
    console.log(`  OS ${String(o.os || o.id).padEnd(6)} ${JSON.stringify(texto.slice(0, 66))}`);
    if (!GRAVAR) return;
    o.obsNotas = notas.concat([{ login: LOGIN, texto, anterior: true }]);
    o.obs = '';
  });

  console.log('');
  console.log(`a migrar: ${migradas}` + (jaTinha ? ` | deixadas como estao: ${jaTinha}` : ''));

  if (!GRAVAR) { console.log('\n(nada foi gravado — rode com --gravar para valer)'); return; }
  if (!migradas) { console.log('nada a fazer.'); return; }

  data.ordens = JSON.stringify(ordens);
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
  console.log('\nGRAVADO no servidor.');
})().catch(e => { console.error('ERRO: ' + (e && e.message || e)); process.exit(1); });
