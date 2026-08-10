/* Carrega os dados no servidor da fábrica, a partir de um backup local.
   Não depende da nuvem — serve inclusive com o projeto da nuvem restrito.

   Uso (na máquina do servidor, com o Supabase local já de pé e o schema.sql
   já rodado):

     node servidor/migrar-do-backup.js \
       --url  http://localhost:8000 \
       --key  <SERVICE_ROLE_KEY> \
       --arq  backups/BACKUP-COMPLETO-2026-08-07T19-57-40.json

   A SERVICE_ROLE_KEY passa por cima das permissões — por isso este script roda
   no servidor, na mão, e nunca no navegador. */

const fs = require('fs');

const arg = n => {
  const i = process.argv.indexOf('--' + n);
  return i > 0 ? process.argv[i + 1] : null;
};
const URL_ = (arg('url') || '').replace(/\/+$/, '');
const KEY = arg('key');
const ARQ = arg('arq');
if (!URL_ || !KEY || !ARQ) {
  console.error('Faltou --url, --key ou --arq. Veja o cabeçalho deste arquivo.');
  process.exit(1);
}

// O backup guarda cada chave como VALOR REAL (array, objeto, número). No blob
// do servidor cada chave é uma STRING JSON — é o formato que o app grava
// (saveState faz JSON.stringify) e o que ele espera ao ler. Converter aqui
// errado faz o app abrir vazio sem dizer por quê.
const META_DO_ARQUIVO = new Set(['exportadoEm', '__snapshot']);

async function principal() {
  const bruto = JSON.parse(fs.readFileSync(ARQ, 'utf8'));
  const blob = {}, versoes = {};
  const carimbo = new Date().toISOString();
  let chaves = 0;
  for (const [k, v] of Object.entries(bruto)) {
    if (META_DO_ARQUIVO.has(k) || k.startsWith('_')) continue;
    blob[k] = JSON.stringify(v);
    versoes[k] = carimbo;
    chaves++;
  }
  if (!chaves) { console.error('O arquivo não tem nenhuma chave de dados.'); process.exit(1); }

  const cabecalhos = {
    'apikey': KEY,
    'Authorization': 'Bearer ' + KEY,
    'Content-Type': 'application/json',
    'Prefer': 'resolution=merge-duplicates'
  };

  // Trava de segurança: nunca gravar por cima de um servidor que JÁ tem dados,
  // a não ser que se peça na marra. Rodar a migração duas vezes por engano
  // desfaria tudo o que foi feito depois da primeira.
  const jaTem = await fetch(`${URL_}/rest/v1/shared_data?id=eq.main&select=id`, {
    headers: { apikey: KEY, Authorization: 'Bearer ' + KEY }
  }).then(r => r.json()).catch(() => []);
  if (Array.isArray(jaTem) && jaTem.length && !process.argv.includes('--sobrescrever')) {
    console.error('\nO servidor JÁ tem dados em shared_data.\n'
      + 'Se a intenção é mesmo substituir tudo, repita com --sobrescrever.\n');
    process.exit(1);
  }

  const r1 = await fetch(`${URL_}/rest/v1/shared_data`, {
    method: 'POST', headers: cabecalhos,
    body: JSON.stringify({ id: 'main', data: blob, updated_at: carimbo })
  });
  if (!r1.ok) { console.error('Falha ao gravar shared_data:', r1.status, await r1.text()); process.exit(1); }

  // O mapa de versões vai junto: sem ele a primeira leitura de cada máquina
  // seria completa de qualquer forma, mas com ele o servidor já nasce pronto
  // para as leituras parciais.
  const r2 = await fetch(`${URL_}/rest/v1/sync_signal`, {
    method: 'POST', headers: cabecalhos,
    body: JSON.stringify({ id: 'main', updated_at: carimbo, device_id: 'migracao', key_versions: versoes })
  });
  if (!r2.ok) console.warn('Aviso: sync_signal não foi gravado:', r2.status, await r2.text());

  console.log(`\n✅ Migrado: ${chaves} chaves em shared_data.`);
  console.log(`   Origem: ${ARQ}`);
  console.log(`   Confira em ${URL_} → Table Editor → shared_data.`);
  console.log(`\n   FALTA COPIAR OS DESENHOS: node servidor/copiar-desenhos.js (veja o README).\n`);
}

principal().catch(e => { console.error(e); process.exit(1); });
