/* Rode com:  node testes/gravacao-parcial.js

   GRAVAR SÓ A PARTE QUE MUDOU.

   Todos os dados vivem numa linha só (shared_data), numa coluna jsonb. Para
   trocar uma chave — marcar uma etapa mexe em `ordens` —, o app devolvia a
   linha INTEIRA: 2,5 MB em 26/08/2026, crescendo ~5 KB por OS nova. Era isso
   que deixava o ícone em "Salvando" por segundos e punha as gravações em fila.

   A leitura já era parcial (mapa de versões por chave). Esta é a escrita: sobe
   só o que este dispositivo sujou, e o BANCO costura (`data || pares`, na
   função gravar_chaves).

   O que este teste protege:

     · sobe o que está sujo, e SÓ isso;
     · o `_device` vai sempre junto — é por ele que a outra aba reconhece a
       gravação como sua ou de terceiro;
     · restaurar/importar (que sujam TODAS as chaves) continuam reescrevendo
       tudo, como antes;
     · servidor sem a função volta ao caminho antigo em vez de quebrar;
     · e o pacote parcial é MESMO pequeno perto do blob.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}

const api = new Function(`
  ${recorte('function _paresParaGravar', 'o montador do pacote')}
  ${recorte('function _ehFuncaoAusente', 'o reconhecedor da funcao ausente')}
  return { _paresParaGravar, _ehFuncaoAusente };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

// Um blob parecido com o de verdade: ordens grande, o resto pequeno.
const CACHE = {
  ordens: JSON.stringify(Array.from({ length: 251 }, (_, i) => ({ id: 'o' + i, os: String(1000 + i), lixo: 'x'.repeat(5000) }))),
  operacoes: 'x'.repeat(500 * 1024),
  grades: 'x'.repeat(133 * 1024),
  desenhos: 'x'.repeat(70 * 1024),
  meta: '{"acessos":{}}',
  _device: 'antigo'
};

console.log('-- o que sobe --');
let pares = api._paresParaGravar(CACHE, new Set(['ordens']), 'dev-1');
ok('1. sobe a chave suja', pares.ordens === CACHE.ordens);
ok('2. e nao sobe as outras', !('operacoes' in pares) && !('grades' in pares) && !('desenhos' in pares),
   Object.keys(pares).join(','));
ok('3. o _device vai sempre junto, com o desta aba', pares._device === 'dev-1', pares._device);
ok('4. e o _device velho do cache nao vaza', Object.keys(pares).join(',') === '_device,ordens',
   Object.keys(pares).join(','));

pares = api._paresParaGravar(CACHE, new Set(['ordens', 'meta']), 'dev-1');
ok('5. duas chaves sujas sobem as duas',
   pares.ordens && pares.meta && Object.keys(pares).length === 3, Object.keys(pares).join(','));

pares = api._paresParaGravar(CACHE, new Set(), 'dev-1');
ok('6. nada sujo: sobe so o carimbo do dispositivo', Object.keys(pares).join(',') === '_device',
   Object.keys(pares).join(','));

// Restaurar backup e importar dados sujam TODAS as chaves — e ai o pacote e o
// blob inteiro mesmo, como antes.
pares = api._paresParaGravar(CACHE, new Set(Object.keys(CACHE)), 'dev-1');
ok('7. restaurar/importar (tudo sujo) reescreve tudo',
   Object.keys(pares).sort().join(',') === Object.keys(CACHE).sort().join(','),
   Object.keys(pares).join(','));

// Chave suja que nao existe no cache (apagada da memoria) nao vira "undefined"
// dentro do jsonb.
pares = api._paresParaGravar(CACHE, new Set(['ordens', 'inexistente']), 'dev-1');
ok('8. chave suja sem valor no cache fica de fora', !('inexistente' in pares), Object.keys(pares).join(','));

console.log('');
console.log('-- o tamanho, que e o ponto --');
const blob = JSON.stringify(CACHE).length;
const parcial = JSON.stringify(api._paresParaGravar(CACHE, new Set(['meta']), 'dev-1')).length;
ok('9. trocar uma chave pequena manda kilobytes, e nao megabytes',
   parcial < 1024 && blob > 1024 * 1024,
   'pacote ' + parcial + ' bytes vs blob ' + Math.round(blob / 1024) + ' KB');
// O PISO, dito em voz alta: salvar uma OS ainda manda a chave `ordens` INTEIRA,
// porque a lista das 251 OS e um valor so dentro do jsonb. O ganho aqui e tirar
// da frente tudo o que nao mudou (operacoes, grades, desenhos); descer desse
// piso exige quebrar `ordens` em um registro por linha, que e outra obra.
const soOrdens = JSON.stringify(api._paresParaGravar(CACHE, new Set(['ordens']), 'dev-1')).length;
ok('10. salvar uma OS manda a chave `ordens`, e nada alem dela',
   soOrdens > CACHE.ordens.length * 0.99 && soOrdens < blob * 0.7,
   Math.round(soOrdens / 1024) + ' KB de ' + Math.round(blob / 1024) + ' KB');
ok('11. e o que sobrou de fora e o grosso do resto',
   (blob - soOrdens) / 1024 > 600, Math.round((blob - soOrdens) / 1024) + ' KB poupados');

console.log('');
console.log('-- servidor sem a funcao --');
ok('12. reconhece o erro do PostgREST', api._ehFuncaoAusente({ code: 'PGRST202' }) === true);
ok('13. reconhece o erro do Postgres', api._ehFuncaoAusente({ code: '42883' }) === true);
ok('14. reconhece pela mensagem',
   api._ehFuncaoAusente({ message: 'Could not find the function public.gravar_chaves' }) === true);
ok('15. e NAO confunde com erro de verdade',
   api._ehFuncaoAusente({ code: '23505', message: 'duplicate key value' }) === false);
ok('16. nem com ausencia de erro', api._ehFuncaoAusente(null) === false);

console.log('');
console.log('-- o caminho no app --');
ok('17. a gravacao normal chama a funcao do banco',
   /supa\.rpc\('gravar_chaves'/.test(src), 'nao chama gravar_chaves');
ok('18. e o caminho antigo continua existindo como rede de seguranca',
   /_semGravacaoParcial = true/.test(src) && /from\('shared_data'\)\.upsert\(/.test(src),
   'sem plano B');
const sql = fs.readFileSync(path.join(__dirname, '..', 'sql', 'supabase-gravar-chaves.sql'), 'utf8');
ok('19. a funcao costura sem apagar o resto (data || pares)',
   /data\s*\|\|\s*pares/.test(sql), 'a costura nao e por concatenacao de jsonb');
ok('20. e roda com a permissao de quem chamou (o RLS continua valendo)',
   /security invoker/.test(sql), 'a funcao driblaria o RLS');

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
