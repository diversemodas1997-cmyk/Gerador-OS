/* Copia as imagens dos desenhos técnicos da NUVEM para o servidor da fábrica,
   e reescreve os endereços guardados nos dados.

   POR QUE ISTO EXISTE
   As imagens não ficam dentro do blob: o blob guarda o ENDEREÇO delas, hoje
   apontando para o Storage da nuvem. Com a internet fora, a folha de OS abre
   sem desenho — que é justamente o centro dela. Sem esta cópia, a rede local
   não substitui a nuvem de verdade.

   QUANDO RODAR
   Com a NUVEM ACESSÍVEL (projeto fora da restrição). É o único passo da
   migração que depende dela; os dados vêm do backup local.

   Uso:
     node servidor/copiar-desenhos.js \
       --local http://localhost:8000 --key <SERVICE_ROLE_KEY> \
       [--so-listar]

   Com --so-listar ele apenas mostra o que faria, sem gravar nada. */

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const LOCAL = (arg('local') || '').replace(/\/+$/, '');
const KEY = arg('key');
const SO_LISTAR = process.argv.includes('--so-listar');
if (!LOCAL || !KEY) { console.error('Faltou --local ou --key.'); process.exit(1); }

const NUVEM = 'https://ckkqrjkhorvaahyazqsr.supabase.co';
const cab = { apikey: KEY, Authorization: 'Bearer ' + KEY };

async function principal() {
  // 1. Ler o blob do servidor LOCAL (já migrado) e achar os endereços da nuvem.
  const linha = await fetch(`${LOCAL}/rest/v1/shared_data?id=eq.main&select=data`, { headers: cab })
    .then(r => r.json());
  const blob = (linha && linha[0] && linha[0].data) || null;
  if (!blob) { console.error('shared_data ainda não foi migrado. Rode migrar-do-backup.js antes.'); process.exit(1); }

  const texto = blob.desenhos || '';
  const re = new RegExp(NUVEM.replace(/[.]/g, '\\.') + '/storage/v1/object/public/desenhos/([^"\\\\ ]+)', 'g');
  const objetos = [...new Set([...texto.matchAll(re)].map(m => m[1]))];
  console.log(`Imagens referenciadas: ${objetos.length}`);
  if (!objetos.length) { console.log('Nada a copiar.'); return; }
  if (SO_LISTAR) { objetos.forEach(o => console.log('  ' + o)); return; }

  // 2. Baixar da nuvem e enviar para o Storage local, MESMO nome de objeto.
  //    Manter o nome idêntico é o que permite, mais adiante, guardar só o
  //    caminho e montar o endereço conforme o servidor em uso.
  let copiadas = 0, falhas = 0;
  for (const obj of objetos) {
    try {
      const img = await fetch(`${NUVEM}/storage/v1/object/public/desenhos/${obj}`);
      if (!img.ok) throw new Error('nuvem respondeu ' + img.status);
      const corpo = Buffer.from(await img.arrayBuffer());
      const tipo = img.headers.get('content-type') || 'image/png';
      const env = await fetch(`${LOCAL}/storage/v1/object/desenhos/${obj}`, {
        method: 'POST',
        headers: Object.assign({}, cab, { 'Content-Type': tipo, 'x-upsert': 'true' }),
        body: corpo
      });
      if (!env.ok) throw new Error('servidor local respondeu ' + env.status + ' ' + await env.text());
      copiadas++;
      process.stdout.write(`\r  copiadas ${copiadas}/${objetos.length}`);
    } catch (e) {
      falhas++;
      console.log(`\n  FALHOU ${obj}: ${e.message}`);
    }
  }
  console.log('');

  if (falhas) {
    console.error(`\n${falhas} imagem(ns) não copiaram. Os endereços NÃO foram reescritos —`
      + ` corrija e rode de novo. Reescrever pela metade deixaria desenhos quebrados.`);
    process.exit(1);
  }

  // 3. Só agora reescrever os endereços, e só se TUDO veio.
  const novo = texto.split(NUVEM + '/storage/v1/object/public/desenhos/')
                    .join(LOCAL + '/storage/v1/object/public/desenhos/');
  const r = await fetch(`${LOCAL}/rest/v1/shared_data?id=eq.main`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json' }),
    body: JSON.stringify({ data: Object.assign({}, blob, { desenhos: novo }) })
  });
  if (!r.ok) { console.error('Falha ao reescrever os endereços:', r.status, await r.text()); process.exit(1); }

  console.log(`\n✅ ${copiadas} imagens copiadas e endereços apontados para ${LOCAL}.`);
  console.log(`\n   ATENÇÃO: os endereços agora são da rede local. Abrindo o app de FORA`);
  console.log(`   da fábrica (modo nuvem), os desenhos não vão aparecer. A solução`);
  console.log(`   definitiva é guardar só o nome do arquivo e montar o endereço conforme`);
  console.log(`   o servidor em uso — está anotado no README como próximo passo.\n`);
}

principal().catch(e => { console.error(e); process.exit(1); });
