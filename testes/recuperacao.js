/* Rode com:  node testes/recuperacao.js
   Ensaio de recuperação de desastre: monta um pacote, cifra, e restaura numa
   pasta limpa — como aconteceria numa máquina nova. Verifica que as CHAVES
   voltam iguais (é o que dispensa reconfigurar as máquinas da fábrica), que as
   imagens voltam byte a byte e que o banco fica pronto para ser aplicado.

   A parte que depende do Docker (pg_dump) não é exercitada aqui — não há Docker
   neste ambiente. O que se testa é tudo o que acontece depois dele. */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { cifrar } = require('../servidor/cofre');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'geros-recup-'));
const dockerDir = path.join(tmp, 'docker');
fs.mkdirSync(dockerDir, { recursive: true });

const SENHA = 'senha-do-backup-2026';
const IMG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 1, 2, 253, 254, 255]);
const PACOTE = {
  formato: 1,
  gerado_em: '2026-08-10T12:00:00.000Z',
  env: {
    JWT_SECRET: 'segredo-original-da-fabrica',
    ANON_KEY: 'eyJhbGciOiJIUzI1NiJ9.anon.assinatura',
    SERVICE_ROLE_KEY: 'eyJhbGciOiJIUzI1NiJ9.svc.assinatura',
    POSTGRES_PASSWORD: 'senha#do$banco'
  },
  banco: 'DROP TABLE IF EXISTS "shared_data";\nCREATE TABLE "shared_data" (id text);\n',
  imagens: [
    { caminho: 'stub/desenhos/1777290341246_1bsz6o.png', b64: IMG.toString('base64') },
    { caminho: 'stub/desenhos/outro.png', b64: IMG.toString('base64') }
  ],
  resumo: { tabelas: 1, imagens: 2 }
};

const arqBkp = path.join(tmp, 'servidor-gerador-os-2026-08-10.bkp');
fs.writeFileSync(arqBkp, cifrar(PACOTE, SENHA));

const roda = (args, esperaErro) => {
  try {
    return { saida: execFileSync(process.execPath,
      [path.join(__dirname, '..', 'servidor', 'restaurar-servidor.js')].concat(args),
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }), erro: null };
  } catch (e) {
    if (!esperaErro) console.log('    (saida do erro) ' + String(e.stderr || '').slice(0, 200));
    return { saida: String(e.stdout || ''), erro: String(e.stderr || e.message) };
  }
};

// 1. Conferência não escreve nada — é o modo para testar o backup sem risco.
let r = roda(['--arq', arqBkp, '--senha', SENHA, '--conferir']);
ok('1. --conferir abre o pacote', /imagens\s*:\s*2/.test(r.saida) && !r.erro, r.saida.slice(0, 200));
ok('1b. --conferir avisa que a ANON_KEY está lá', /ANON_KEY presente:\s*sim/.test(r.saida));
ok('1c. --conferir não escreveu nada', fs.readdirSync(dockerDir).length === 0);

// 2. Senha errada não restaura nada.
r = roda(['--arq', arqBkp, '--senha', 'senha-errada', '--docker', dockerDir], true);
ok('2. senha errada aborta', !!r.erro && /Senha errada/.test(r.erro));
ok('2b. e não deixou lixo para trás', fs.readdirSync(dockerDir).length === 0);

// 3. Restauração de verdade.
r = roda(['--arq', arqBkp, '--senha', SENHA, '--docker', dockerDir]);
ok('3. restaurou sem erro', !r.erro);

const env = fs.readFileSync(path.join(dockerDir, '.env'), 'utf8');
ok('3b. ANON_KEY volta IDÊNTICA (máquinas não precisam ser tocadas)',
   env.includes('ANON_KEY=eyJhbGciOiJIUzI1NiJ9.anon.assinatura'));
ok('3c. JWT_SECRET volta idêntico', env.includes('JWT_SECRET=segredo-original-da-fabrica'));
ok('3d. senha do banco com caracteres especiais intacta',
   env.includes('POSTGRES_PASSWORD=senha#do$banco'));

const img1 = fs.readFileSync(path.join(dockerDir, 'volumes', 'storage',
  'stub', 'desenhos', '1777290341246_1bsz6o.png'));
ok('3e. imagem volta byte a byte', img1.equals(IMG), 'tamanho ' + img1.length);
ok('3f. as duas imagens voltaram', fs.existsSync(
  path.join(dockerDir, 'volumes', 'storage', 'stub', 'desenhos', 'outro.png')));

const sql = fs.readFileSync(path.join(dockerDir, 'restaurar-banco.sql'), 'utf8');
ok('3g. banco gravado para aplicar depois', sql.includes('CREATE TABLE "shared_data"'));
ok('3h. instruções dizem os comandos que faltam', /docker compose up -d/.test(r.saida)
   && /psql -U postgres/.test(r.saida));

// 4. Restaurar por cima não pode destruir o .env que já estava lá sem cópia.
r = roda(['--arq', arqBkp, '--senha', SENHA, '--docker', dockerDir]);
ok('4. .env anterior é preservado numa cópia',
   fs.existsSync(path.join(dockerDir, '.env.antes-da-restauracao')));

fs.rmSync(tmp, { recursive: true, force: true });
console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
