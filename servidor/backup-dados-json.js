/*
 * Backup diario dos dados do Gerador-OS puxando DIRETO do banco do servidor.
 * Faz login com a conta admin, le a linha shared_data.id='main' (blob com todos
 * os cadastros e OS) e grava um arquivo diretamente importavel no app, mais as
 * imagens dos desenhos. Commit/push so quando muda.
 *
 * NAO substitui o servidor\backup-diario.ps1, que gera o pacote cifrado de 50 MB
 * para restaurar a maquina inteira. Este aqui e a outra ponta: JSON legivel e
 * versionado em git, para recuperar UMA coisa (um desenho, uma OS) sem precisar
 * restaurar o servidor todo, e para enxergar o historico do que mudou.
 *
 * Credenciais ficam em supa-creds.json, na pasta de destino (nunca versionado).
 * Executado pela Tarefa Agendada "Gerador-OS Backup Dados Diario", 16:00.
 *
 * ONDE ELE MORA E ONDE ELE ESCREVE sao lugares diferentes, de proposito:
 *
 *   mora aqui  -> servidor\backup-dados-json.js, versionado com o resto do app
 *   escreve em -> J:\Meu Drive\Backup ERP Diverse\Gerador-OS-backup-dados
 *
 * Ate 11/08/2026 ele morava DENTRO da pasta de destino, no Google Drive, onde o
 * .gitignore de la excluia justamente ele: o unico arquivo sem historico era o
 * que fazia o backup. Chamava-se backup-supabase.js e ficou com o nome errado
 * quando a nuvem saiu de cena. A copia velha continua la, marcada como obsoleta;
 * a tarefa agendada aponta para ESTE arquivo. Mexer no de la nao muda nada.
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

// Pasta de DESTINO (e o repositorio git dos dados), nao a pasta deste script.
const REPO = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados';
const GIT = 'C:/Program Files/Git/mingw64/bin/git.exe';

// Servidor da fabrica, nao mais a nuvem.
//
// A nuvem foi restringida por cota de trafego em 10/08/2026 e este backup passou
// a falhar TODO DIA as 16:00 com "login falhou (402)" — 402 e Payment Required,
// o proprio Supabase barrando o projeto. Ficou 2 dias assim antes de alguem ver.
//
// Vai por localhost:8000 (o Kong, dentro da maquina) e nao por
// https://193.168.0.200 de proposito: este script roda NO servidor, e o Node nao
// usa a loja de certificados do Windows — pelo endereco https ele recusaria o
// certificado proprio da fabrica e o backup morreria por um motivo bobo.
const SUPA_URL = 'http://localhost:8000';

// A chave anonima NAO fica escrita aqui, de proposito. O instalador ja a publica
// em servidor\tls\servidor-local.json — o mesmo arquivo que o app le para se
// conectar sozinho. Lendo de la: trocar a chave do servidor nao quebra o backup,
// e nao existe uma segunda copia dela largada dentro do Google Drive.
const LOCAL = 'C:/Users/Pichau/Desktop/Gerador-OS/servidor/tls/servidor-local.json';
let ANON = '';

const CREDS = path.join(REPO, 'supa-creds.json');
const OUT = path.join(REPO, 'dados-supabase.json');
const IMGDIR = path.join(REPO, 'desenhos-imagens');
const LOG = path.join(REPO, 'backup.log');

function log(m) {
  fs.appendFileSync(LOG, `${new Date().toISOString()}  ${m}\n`);
}
function git(args) {
  return execFileSync(GIT, args, { cwd: REPO, stdio: ['ignore', 'pipe', 'pipe'] });
}

(async () => {
  try {
    if (!fs.existsSync(CREDS)) { log('ERRO: supa-creds.json nao encontrado'); process.exit(1); }
    let email, password;
    try { ({ email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'))); }
    catch (e) { log('ERRO: supa-creds.json invalido (JSON malformado)'); process.exit(1); }
    if (!email || !password || /COLOQUE_SUA_SENHA/.test(password)) {
      log('ERRO: preencha a senha em supa-creds.json'); process.exit(1);
    }

    // Chave anonima do servidor local (ver comentario no topo).
    if (!fs.existsSync(LOCAL)) {
      log(`ERRO: nao achei ${LOCAL} — o servidor local foi reinstalado?`); process.exit(1);
    }
    try {
      // O arquivo e gravado com BOM; sem tirar, o JSON.parse morre no 1o byte.
      // Escrito como \uFEFF e nao como o caractere: um BOM literal no meio do
      // codigo e invisivel no editor e vira caca ao fantasma quando quebra.
      ANON = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
    } catch (e) {
      log('ERRO: servidor-local.json ilegivel: ' + ((e && e.message) || e)); process.exit(1);
    }
    if (!ANON) { log('ERRO: servidor-local.json sem a chave'); process.exit(1); }

    // 1) Login -> token temporario
    const authRes = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, {
      method: 'POST',
      headers: { apikey: ANON, 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    });
    const auth = await authRes.json();
    if (!auth.access_token) {
      log(`ERRO: login falhou (${auth.error_code || auth.msg || authRes.status})`); process.exit(1);
    }

    // 2) Le o blob compartilhado
    const dataRes = await fetch(`${SUPA_URL}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, {
      headers: { apikey: ANON, Authorization: `Bearer ${auth.access_token}` },
    });
    const rows = await dataRes.json();
    if (!Array.isArray(rows) || !rows.length || !rows[0].data) {
      log(`ERRO: leitura de shared_data falhou (HTTP ${dataRes.status})`); process.exit(1);
    }
    const blob = rows[0].data;

    // O Supabase guarda cada chave como STRING JSON (mesmo formato do backup
    // local). Parseia p/ arrays/objetos reais -> arquivo diretamente importavel
    // no app (importarDados espera Array.isArray(data[k])).
    const clean = {};
    for (const k of Object.keys(blob)) {
      const v = blob[k];
      if (typeof v === 'string') {
        try { clean[k] = JSON.parse(v); } catch (e) { clean[k] = v; }
      } else {
        clean[k] = v;
      }
    }
    const nOrd = Array.isArray(clean.ordens) ? clean.ordens.length : '?';
    const nDes = Array.isArray(clean.desenhos) ? clean.desenhos.length : '?';

    // Sanidade: nao sobrescrever backup bom com blob vazio/quebrado
    if (nOrd === '?' && nDes === '?') { log('ERRO: blob sem ordens/desenhos — abortando p/ nao corromper backup'); process.exit(1); }

    // 3) Grava arquivo importavel (metadados _fonte/_updated_at sao ignorados no import)
    const out = { _fonte: 'servidor local shared_data id=main', _updated_at: rows[0].updated_at, ...clean };
    fs.writeFileSync(OUT, JSON.stringify(out));

    // 3b) Baixa as IMAGENS dos desenhos (bucket publico 'desenhos' do Storage).
    // O blob so guarda a URL — se o Storage se perder, o desenho vai junto.
    // Baixa so o que ainda nao existe na pasta (as imagens nao mudam: nome tem timestamp).
    let novas = 0, falhas = 0;
    if (Array.isArray(clean.desenhos)) {
      if (!fs.existsSync(IMGDIR)) fs.mkdirSync(IMGDIR, { recursive: true });
      for (const d of clean.desenhos) {
        const img = d && typeof d.img === 'string' ? d.img : '';
        if (!img) continue;
        // No servidor local o blob guarda so o NOME do arquivo
        // ("1786379660987_qm2ow9.png"); na nuvem guardava a URL inteira.
        // Aceitar as duas formas: com a regra antiga (exigir URL absoluta) o
        // passo de imagens era pulado EM SILENCIO e o backup saia sem desenho
        // nenhum, com cara de completo — o pior tipo de backup que existe.
        const url = /^https?:\/\//.test(img)
          ? img
          : `${SUPA_URL}/storage/v1/object/public/desenhos/${encodeURIComponent(img)}`;
        const nome = decodeURIComponent(url.split('/').pop().split('?')[0]).replace(/[\\/:*?"<>|]/g, '_');
        const dest = path.join(IMGDIR, nome);
        if (fs.existsSync(dest)) continue;
        try {
          const r = await fetch(url);
          if (!r.ok) throw new Error('HTTP ' + r.status);
          fs.writeFileSync(dest, Buffer.from(await r.arrayBuffer()));
          novas++;
        } catch (e) {
          falhas++;
          log(`AVISO: falha ao baixar imagem ${nome} (${(e && e.message) || e})`);
        }
      }
    }

    // 4) Commit/push so se mudou
    git(['add', 'dados-supabase.json', 'desenhos-imagens']);
    let changed = true;
    try { git(['diff', '--cached', '--quiet']); changed = false; } catch (e) { changed = true; }
    const sufImg = `, imagens novas=${novas}${falhas ? `, falhas=${falhas}` : ''}`;
    if (!changed) { log(`OK: sem mudancas (OS=${nOrd}, desenhos=${nDes}${sufImg}); nada a enviar`); process.exit(0); }

    const stamp = new Date().toISOString().slice(0, 19).replace('T', ' ');
    git(['commit', '-q', '-m', `Backup servidor local ${stamp} (OS=${nOrd}, desenhos=${nDes}${sufImg})`]);
    git(['push', '-q', 'origin', 'main']);
    log(`OK: backup enviado (OS=${nOrd}, desenhos=${nDes}${sufImg})`);
  } catch (e) {
    log('ERRO: ' + ((e && e.message) || e));
    process.exit(1);
  }
})();
