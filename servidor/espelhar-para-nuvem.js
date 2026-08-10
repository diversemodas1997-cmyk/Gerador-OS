/* Copia o estado do servidor da fábrica para a nuvem. SEMPRE de mão única.
   Roda NO SERVIDOR, de tempos em tempos (ver Agendador de Tarefas no README).

   Para que serve: a nuvem passa a ter uma cópia viva, que serve para consultar
   de fora da fábrica e como segunda cópia dos dados. Ninguém edita por lá — o
   app abre a nuvem em modo consulta —, então não há dois lados para conciliar.

   Uso:
     node servidor\espelhar-para-nuvem.js ^
       --local     http://localhost:8000 ^
       --local-key <SERVICE_ROLE_KEY do servidor da fábrica> ^
       --nuvem     https://ckkqrjkhorvaahyazqsr.supabase.co ^
       --nuvem-key <SERVICE_ROLE_KEY da nuvem> ^
       [--forcar] [--sem-imagens]

   A SERVICE_ROLE_KEY da nuvem passa por cima de todas as permissões. Ela fica
   só neste comando, na máquina do servidor — nunca no navegador e nunca dentro
   do repositório.
*/
const { espelhar } = require('./espelho');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const cfg = {
  local: (arg('local') || '').replace(/\/+$/, ''),
  localKey: arg('local-key'),
  nuvem: (arg('nuvem') || '').replace(/\/+$/, ''),
  nuvemKey: arg('nuvem-key'),
  forcar: process.argv.includes('--forcar'),
  semImagens: process.argv.includes('--sem-imagens'),
  log: m => console.log(`[${new Date().toISOString().slice(11, 19)}] ${m}`)
};
if (!cfg.local || !cfg.localKey || !cfg.nuvem || !cfg.nuvemKey) {
  console.error('Faltou --local, --local-key, --nuvem ou --nuvem-key. Veja o cabeçalho deste arquivo.');
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
