/* Pacote de recuperação do servidor da fábrica.
   Roda NO SERVIDOR, uma vez por dia (ver Agendador de Tarefas no README).

   Gera UM arquivo cifrado com tudo o que é preciso para levantar um servidor
   idêntico noutra máquina:

     - o banco inteiro (pg_dump): dados, contas de login com as senhas,
       papéis, metadados do Storage;
     - os arquivos das imagens dos desenhos;
     - o .env com as CHAVES ORIGINAIS.

   O .env é o detalhe que muda tudo. Restaurando as mesmas chaves, o servidor
   novo nasce com a mesma ANON_KEY — e as máquinas da fábrica voltam a funcionar
   SEM reconfigurar nada, uma a uma. É a diferença entre recuperar em minutos e
   passar a manhã de máquina em máquina.

   Uso:
     node servidor\backup-servidor.js ^
       --docker  C:\supabase\docker ^
       --destino "J:\Meu Drive\Backup Gerador-OS" ^
       --senha   <senha-do-backup> ^
       [--manter 14] [--com-snapshots]
*/
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { cifrar, decifrar } = require('./cofre');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const DOCKER = arg('docker');
const DESTINO = arg('destino');
const SENHA = arg('senha');
const MANTER = parseInt(arg('manter') || '14', 10);
if (!DOCKER || !DESTINO || !SENHA) {
  console.error('Faltou --docker, --destino ou --senha. Veja o cabeçalho deste arquivo.');
  process.exit(1);
}

const log = m => console.log(`[${new Date().toISOString().slice(11, 19)}] ${m}`);

function lerEnv(arquivo) {
  const env = {};
  for (const linha of fs.readFileSync(arquivo, 'utf8').split(/\r?\n/)) {
    const m = linha.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)$/);
    if (m) env[m[1]] = m[2];
  }
  return env;
}

// Percorre a pasta do Storage e traz cada arquivo em base64. Sem depender de
// tar/zip: no Windows isso evita uma peça a mais que pode faltar na máquina.
function lerArquivos(raiz) {
  const saida = [];
  if (!fs.existsSync(raiz)) return saida;
  (function anda(dir) {
    for (const nome of fs.readdirSync(dir)) {
      const p = path.join(dir, nome);
      const st = fs.statSync(p);
      if (st.isDirectory()) anda(p);
      else saida.push({
        caminho: path.relative(raiz, p).split(path.sep).join('/'),
        b64: fs.readFileSync(p).toString('base64')
      });
    }
  })(raiz);
  return saida;
}

function principal() {
  const envArq = path.join(DOCKER, '.env');
  if (!fs.existsSync(envArq)) { console.error('Não achei o .env em ' + DOCKER); process.exit(1); }

  log('Lendo as chaves do servidor…');
  const env = lerEnv(envArq);

  log('Exportando o banco (pg_dump)…');
  // Os snapshots diários ficam de fora por padrão: são ~30 cópias do blob e
  // sozinhos inchariam o pacote em dezenas de MB, todo dia, numa pasta que
  // sincroniza. O que importa para levantar o servidor é o estado atual — e
  // este pacote JÁ é um backup.
  const argsDump = ['exec', 'supabase-db', 'pg_dump', '-U', 'postgres', '-d', 'postgres',
    '--clean', '--if-exists', '--quote-all-identifiers'];
  if (!process.argv.includes('--com-snapshots')) {
    argsDump.push('--exclude-table-data=public.shared_data_backups');
  }
  let dump;
  try {
    dump = execFileSync('docker', argsDump, { maxBuffer: 1024 * 1024 * 1024 }).toString('utf8');
  } catch (e) {
    console.error('pg_dump falhou. O Docker está rodando e o container supabase-db de pé?');
    console.error(String(e.stderr || e.message).slice(0, 800));
    process.exit(1);
  }
  if (!/CREATE TABLE/i.test(dump)) {
    console.error('O pg_dump saiu sem nenhuma tabela — não vou gravar um backup vazio.');
    process.exit(1);
  }

  log('Lendo as imagens dos desenhos…');
  const imagens = lerArquivos(path.join(DOCKER, 'volumes', 'storage'));

  const pacote = {
    formato: 1,
    gerado_em: new Date().toISOString(),
    env,
    banco: dump,
    imagens,
    resumo: { tabelas: (dump.match(/CREATE TABLE/gi) || []).length, imagens: imagens.length }
  };

  fs.mkdirSync(DESTINO, { recursive: true });
  const nome = `servidor-gerador-os-${new Date().toISOString().slice(0, 10)}.bkp`;
  const alvo = path.join(DESTINO, nome);
  log('Cifrando e gravando…');
  const cifrado = cifrar(pacote, SENHA);
  fs.writeFileSync(alvo, cifrado);

  // CONFERE O QUE ACABOU DE GRAVAR. Um backup nunca aberto não é backup: é uma
  // suposição. Relê do disco e decifra — pega senha errada, disco cheio e
  // gravação truncada agora, e não no dia do desastre.
  log('Conferindo o arquivo gravado…');
  try {
    const volta = decifrar(fs.readFileSync(alvo), SENHA);
    if (!volta.banco || volta.imagens.length !== imagens.length || !volta.env.JWT_SECRET) {
      throw new Error('o pacote reaberto veio incompleto');
    }
  } catch (e) {
    console.error('FALHA NA CONFERÊNCIA: ' + e.message);
    try { fs.unlinkSync(alvo); } catch (_) {}
    console.error('O arquivo foi apagado para não passar por um backup bom.');
    process.exit(1);
  }

  // Retenção
  const antigos = fs.readdirSync(DESTINO)
    .filter(f => /^servidor-gerador-os-.*\.bkp$/.test(f)).sort();
  let apagados = 0;
  while (antigos.length > MANTER) {
    fs.unlinkSync(path.join(DESTINO, antigos.shift())); apagados++;
  }

  const mb = (cifrado.length / 1048576).toFixed(1);
  log(`OK — ${nome} (${mb} MB): ${pacote.resumo.tabelas} tabelas, ${imagens.length} imagens.`);
  if (apagados) log(`${apagados} pacote(s) antigo(s) removido(s); mantendo ${MANTER}.`);
  log(`Destino: ${DESTINO}`);
}

principal();
