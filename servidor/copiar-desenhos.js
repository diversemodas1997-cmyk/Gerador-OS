/* Copia as imagens dos desenhos técnicos da NUVEM para o servidor da fábrica.

   POR QUE ISTO EXISTE
   As imagens não ficam dentro dos dados: os dados guardam só o NOME do arquivo,
   e o app monta o endereço contra o servidor em uso (ver urlDesenho no app.js).
   Ou seja: o mesmo desenho é buscado na fábrica quando ela está de pé, e na
   nuvem quando não. Para a parte da fábrica funcionar, os arquivos precisam
   existir lá — é o que este script faz.

   Como o nome do arquivo é o mesmo dos dois lados, NADA nos dados precisa ser
   alterado. Este script só copia arquivos, e por isso é seguro rodar de novo
   quantas vezes quiser.

   QUANDO RODAR
   Com a NUVEM ACESSÍVEL. É o único passo da migração que depende dela — os
   dados vêm do backup local. Faça antes de virar a chave para a rede local.

   Uso:
     node servidor\copiar-desenhos.js ^
       --local http://localhost:8000 --key <SERVICE_ROLE_KEY> ^
       [--so-listar]
*/
const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const LOCAL = (arg('local') || '').replace(/\/+$/, '');
const KEY = arg('key');
const SO_LISTAR = process.argv.includes('--so-listar');
if (!LOCAL || !KEY) { console.error('Faltou --local ou --key.'); process.exit(1); }

const NUVEM = 'https://ckkqrjkhorvaahyazqsr.supabase.co';
const ANON_NUVEM = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNra3Fyamtob3J2YWFoeWF6cXNyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4MTY2MjMsImV4cCI6MjA5MjM5MjYyM30.yT3Tb6KKx4sDNJXetwIoA77WudWUqQ2gCgT7JLi0iT8';
const cabLocal = { apikey: KEY, Authorization: 'Bearer ' + KEY };

async function listarNaNuvem() {
  // Lista o bucket direto, em vez de garimpar nomes dentro dos dados: assim
  // nenhuma imagem fica para trás por estar referenciada de um jeito inesperado.
  const nomes = [];
  for (let pagina = 0; ; pagina++) {
    const r = await fetch(`${NUVEM}/storage/v1/object/list/desenhos`, {
      method: 'POST',
      headers: { apikey: ANON_NUVEM, Authorization: 'Bearer ' + ANON_NUVEM,
                 'Content-Type': 'application/json' },
      body: JSON.stringify({ prefix: '', limit: 100, offset: pagina * 100 })
    });
    if (!r.ok) throw new Error(`a nuvem respondeu ${r.status} ao listar: ${await r.text()}`);
    const lote = await r.json();
    if (!Array.isArray(lote) || !lote.length) break;
    lote.forEach(o => { if (o && o.name) nomes.push(o.name); });
    if (lote.length < 100) break;
  }
  return nomes;
}

async function principal() {
  const nomes = await listarNaNuvem();
  console.log(`Imagens no bucket da nuvem: ${nomes.length}`);
  if (!nomes.length) { console.log('Nada a copiar.'); return; }
  if (SO_LISTAR) { nomes.forEach(n => console.log('  ' + n)); return; }

  let copiadas = 0, falhas = 0;
  for (const nome of nomes) {
    try {
      const img = await fetch(`${NUVEM}/storage/v1/object/public/desenhos/${encodeURIComponent(nome)}`);
      if (!img.ok) throw new Error('nuvem respondeu ' + img.status);
      const corpo = Buffer.from(await img.arrayBuffer());
      if (!corpo.length) throw new Error('veio vazia');
      const env = await fetch(`${LOCAL}/storage/v1/object/desenhos/${encodeURIComponent(nome)}`, {
        method: 'POST',
        headers: Object.assign({}, cabLocal, {
          'Content-Type': img.headers.get('content-type') || 'image/png',
          'x-upsert': 'true'
        }),
        body: corpo
      });
      if (!env.ok) throw new Error('servidor local respondeu ' + env.status + ' ' + await env.text());
      copiadas++;
      process.stdout.write(`\r  copiadas ${copiadas}/${nomes.length}`);
    } catch (e) {
      falhas++;
      console.log(`\n  FALHOU ${nome}: ${e.message}`);
    }
  }
  console.log('');

  if (falhas) {
    console.error(`\n⚠️  ${falhas} imagem(ns) não copiaram. Rode de novo — copiar por cima`
      + ` não faz mal. Enquanto faltarem, esses desenhos aparecem em branco quando`
      + ` o app estiver falando com a fábrica.\n`);
    process.exit(1);
  }
  console.log(`\n✅ ${copiadas} imagens no servidor da fábrica.`);
  console.log(`   Os dados não precisaram ser alterados: o app monta o endereço`);
  console.log(`   conforme o servidor em uso.\n`);
}

principal().catch(e => { console.error(e.message); process.exit(1); });
