/* Copia o estado do servidor da fábrica para a nuvem. SEMPRE de mão única.
   Roda NO SERVIDOR, de tempos em tempos (ver Agendador de Tarefas no README).

   Para que serve: a nuvem passa a ter uma cópia viva, que serve para consultar
   de fora da fábrica e como segunda cópia dos dados. Ninguém edita por lá — o
   app abre a nuvem em modo consulta —, então não há dois lados para conciliar.

   Uso (sem nada: usa os endereços e as chaves de sempre):
     node servidor\espelhar-para-nuvem.js [--forcar] [--sem-imagens] [--sem-mensagens]

   ONDE ELE ACHA AS CHAVES, quando não vêm na linha de comando:
     · a da FÁBRICA, no .env do Supabase local (C:\supabase\docker\.env), que é
       de onde todos os outros scripts do servidor já a leem;
     · a da NUVEM, no arquivo C:\supabase\nuvem-service-role.key — uma linha só,
       a chave e nada mais (linhas começadas por # são ignoradas).

   Por que num arquivo, e não no comando: a SERVICE_ROLE_KEY da nuvem passa por
   cima de todas as permissões. Na linha de comando ela fica visível no
   Agendador de Tarefas, no histórico do PowerShell e em qualquer print da tela.
   Num arquivo só do servidor, não. Nunca no navegador, nunca no repositório.

   Tudo continua podendo ser passado à mão:
     --local <url> --local-key <chave> --nuvem <url> --nuvem-key <chave>
     --env <caminho do .env> --nuvem-key-arquivo <caminho>
*/
const fs = require('fs');

const ENV_PADRAO = 'C:\\supabase\\docker\\.env';
const CHAVE_NUVEM_PADRAO = 'C:\\supabase\\nuvem-service-role.key';

// A primeira linha que não é comentário nem vazia. Assim o arquivo pode ter uma
// explicação em cima da chave sem quebrar nada.
function chaveDoArquivo(caminho) {
  try {
    return (fs.readFileSync(caminho, 'utf8').split(/\r?\n/)
      .map(l => l.trim()).filter(l => l && !l.startsWith('#'))[0]) || '';
  } catch (e) { return ''; }
}

function chaveDoEnv(caminho) {
  try {
    const m = fs.readFileSync(caminho, 'utf8').match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m);
    return m ? m[1].trim() : '';
  } catch (e) { return ''; }
}
const { espelhar } = require('./espelho');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const cfg = {
  local: (arg('local') || 'http://localhost:8000').replace(/\/+$/, ''),
  localKey: arg('local-key') || chaveDoEnv(arg('env') || ENV_PADRAO),
  nuvem: (arg('nuvem') || 'https://ckkqrjkhorvaahyazqsr.supabase.co').replace(/\/+$/, ''),
  nuvemKey: arg('nuvem-key') || chaveDoArquivo(arg('nuvem-key-arquivo') || CHAVE_NUVEM_PADRAO),
  forcar: process.argv.includes('--forcar'),
  semImagens: process.argv.includes('--sem-imagens'),
  semMensagens: process.argv.includes('--sem-mensagens'),
  log: m => console.log(`[${new Date().toISOString().slice(11, 19)}] ${m}`)
};
if (!cfg.localKey) {
  console.error('Não achei a chave do servidor da fábrica.\n'
    + `  Esperava SERVICE_ROLE_KEY em ${arg('env') || ENV_PADRAO}\n`
    + '  (ou passe --local-key <chave>).');
  process.exit(1);
}
if (!cfg.nuvemKey) {
  console.error('Não achei a chave da NUVEM.\n'
    + `  Ponha a service_role key da nuvem em ${arg('nuvem-key-arquivo') || CHAVE_NUVEM_PADRAO}\n`
    + '  — uma linha só, a chave e nada mais. Ela está no painel do Supabase,\n'
    + '  em Project Settings → API → service_role. (Ou passe --nuvem-key <chave>.)');
  process.exit(1);
}

espelhar(cfg).then(rel => {
  if (rel.dados === 'bloqueado') {
    console.error(`\n⛔ ESPELHO BLOQUEADO — nada foi enviado.\n   ${rel.motivo}\n`
      + '   Isto costuma significar que o servidor da fábrica está sem dados por\n'
      + '   algum problema (restauração pela metade, migração que não rodou).\n'
      + '   A nuvem foi PRESERVADA. Resolva a fábrica antes de espelhar de novo.\n');
    process.exit(2);
  }
  const parte = rel.dados === 'espelhado' ? 'dados enviados' : `dados: ${rel.motivo || 'sem mudança'}`;
  cfg.log(`${parte}; imagens novas: ${rel.imagens}`);
}).catch(e => {
  // Internet fora é o caso mais comum e não é motivo de alarme: o espelho é
  // best-effort e a próxima execução recupera o atraso sozinha.
  console.error('Espelho não concluiu: ' + e.message);
  process.exit(1);
});
