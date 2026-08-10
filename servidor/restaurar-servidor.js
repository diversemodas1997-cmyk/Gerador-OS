/* Levanta um servidor novo a partir de um pacote de recuperação.
   Roda na MÁQUINA NOVA, depois de instalar o Docker e clonar o Supabase.

   Uso:
     node servidor\restaurar-servidor.js ^
       --arq     "J:\Meu Drive\Backup Gerador-OS\servidor-gerador-os-2026-08-10.bkp" ^
       --docker  C:\supabase\docker ^
       --senha   <senha-do-backup> ^
       [--conferir]

   Com --conferir ele só abre o pacote e mostra o que tem dentro, sem escrever
   nada. Use de vez em quando, numa máquina qualquer: é o único jeito de saber
   que o backup presta ANTES de precisar dele.
*/
const fs = require('fs');
const path = require('path');
const { decifrar } = require('./cofre');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const ARQ = arg('arq');
const DOCKER = arg('docker');
const SENHA = arg('senha');
const SO_CONFERIR = process.argv.includes('--conferir');
if (!ARQ || !SENHA || (!DOCKER && !SO_CONFERIR)) {
  console.error('Faltou --arq, --senha ou --docker. Veja o cabeçalho deste arquivo.');
  process.exit(1);
}

const pacote = decifrar(fs.readFileSync(ARQ), SENHA);   // erra aqui se a senha ou o arquivo estiverem ruins

console.log(`
Pacote de ${pacote.gerado_em}
  tabelas no banco : ${pacote.resumo ? pacote.resumo.tabelas : '?'}
  imagens          : ${pacote.imagens.length}
  chaves do .env   : ${Object.keys(pacote.env).length}
  ANON_KEY presente: ${pacote.env.ANON_KEY ? 'sim' : 'NÃO — as máquinas precisariam ser reconfiguradas'}
`);

if (SO_CONFERIR) {
  console.log('Conferência apenas — nada foi escrito.\n');
  process.exit(0);
}

// 1. As chaves ORIGINAIS. É o que faz as máquinas da fábrica voltarem a
//    funcionar sem serem tocadas uma a uma.
const envArq = path.join(DOCKER, '.env');
if (fs.existsSync(envArq)) {
  const salvo = envArq + '.antes-da-restauracao';
  fs.copyFileSync(envArq, salvo);
  console.log('.env que já existia foi guardado em ' + path.basename(salvo));
}
fs.writeFileSync(envArq,
  Object.entries(pacote.env).map(([k, v]) => `${k}=${v}`).join('\n') + '\n');
console.log('✔ .env restaurado (mesmas chaves de antes)');

// 2. As imagens dos desenhos, de volta ao lugar de onde o Storage as serve.
const raiz = path.join(DOCKER, 'volumes', 'storage');
for (const img of pacote.imagens) {
  const destino = path.join(raiz, img.caminho.split('/').join(path.sep));
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  fs.writeFileSync(destino, Buffer.from(img.b64, 'base64'));
}
console.log(`✔ ${pacote.imagens.length} imagens restauradas`);

// 3. O banco fica num .sql ao lado, para ser aplicado DEPOIS que os containers
//    subirem — não dá para restaurar num banco que ainda não existe.
const sqlArq = path.join(DOCKER, 'restaurar-banco.sql');
fs.writeFileSync(sqlArq, pacote.banco, 'utf8');
console.log('✔ banco gravado em ' + sqlArq);

console.log(`
=====================================================================
 FALTAM 3 COMANDOS, nesta ordem, dentro de ${DOCKER}
=====================================================================

 1) Subir os containers com as chaves restauradas:

      docker compose up -d

    Espere uns 2 minutos na primeira vez.

 2) Aplicar o banco:

      type restaurar-banco.sql | docker exec -i supabase-db psql -U postgres -d postgres

    Vai passar muita coisa na tela. Avisos de "does not exist" no começo
    são normais — sao os DROP do dump limpando o que ainda nao existe.

 3) Reiniciar, para tudo reler o banco restaurado:

      docker compose restart

 DEPOIS, CONFIRA:
   - http://localhost:8000  ->  Table Editor -> shared_data tem a linha 'main'
   - http://localhost:8000  ->  Authentication -> as contas estao la
   - abra o Gerador-OS numa maquina da fabrica: ela deve voltar sozinha,
     SEM reconfigurar, porque a ANON_KEY e a mesma.

 SE O IP DA MAQUINA NOVA FOR DIFERENTE do servidor antigo, aí sim é
 preciso passar nas máquinas e corrigir o endereço em Configurações ->
 Servidor da fábrica. Dá para evitar isso dando à máquina nova o mesmo
 IP fixo da antiga.
=====================================================================
`);
