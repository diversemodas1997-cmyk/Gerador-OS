/*
 * DESENHO 0025 DE VOLTA AO SEU CADASTRO, 0050 NUM CADASTRO PROPRIO, E A FOTO DO
 * DESENHO GRAVADA EM CADA OS.
 *
 * Junior, 15/10/2026: "Analise a OS547 e descubra porque o desenho tecnico
 * mudou, de preto/mostarda/off-white para Azul/branco/preto".
 *
 * Em 14/10, 11:37, o desenho 0025 (Blusa Moletom Tricolor | Preto/Mostarda/
 * Off-White) foi DUPLICADO. A copia nasceu como "0025 (copia)" e virou "0025 ".
 * Quem foi editado para 0050 Azul/Branco/Preto (imagem nova as 11:42) foi o
 * ORIGINAL -- o cadastro a que 14 OS antigas apontam (0211 ... 0548). Como a
 * folha lia o desenho ao vivo, as 14 passaram a mostrar o desenho Azul.
 *
 * O script troca o CONTEUDO dos dois cadastros (os ids ficam):
 *   id_1776955198418_953 -> volta a ser 0025 Preto/Mostarda/Off-White
 *   id_1791988631415_576 -> passa a ser 0050 Azul/Branco/Preto
 * e aponta para o _576 so as OS do Azul (0661 e 0662).
 *
 * Depois grava em TODA OS ligada a um desenho a `desenhoFoto` (ver
 * _fotoDoDesenho no app.js): a partir dela, editar o cadastro do desenho nao
 * muda mais nenhuma OS ja emitida.
 *
 *   node servidor/desenho-0025-0050-1510.js            (so mostra)
 *   node servidor/desenho-0025-0050-1510.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const ORIGINAL = 'id_1776955198418_953';
const COPIA = 'id_1791988631415_576';
const OS_DO_AZUL = ['0661', '0662'];

async function req(method, caminho, body, extra) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: Object.assign({ apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' }, extra || {}),
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t;
}
const arr = v => { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; };

// Mesma foto do app (_fotoDoDesenho): so o que identifica o desenho na folha.
function foto(d, em) {
  return {
    id: d.id, codigo: d.codigo || '', desc: d.desc || '', img: d.img || '',
    skuLinha: d.skuLinha || '',
    corPrincipalId: d.corPrincipalId || '', corSecundariaId: d.corSecundariaId || '',
    corTerciariaId: d.corTerciariaId || '', em
  };
}

(async () => {
  const linha = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = linha.data;
  const desenhos = arr(data.desenhos);
  const ordens = arr(data.ordens);
  const dO = desenhos.find(d => d.id === ORIGINAL);
  const dC = desenhos.find(d => d.id === COPIA);
  if (!dO || !dC) throw new Error('nao achei os dois cadastros');

  const jaFeito = String(dO.codigo).trim() === '0025' && String(dC.codigo).trim() === '0050';
  if (!jaFeito) {
    if (String(dO.codigo).trim() !== '0050' || String(dC.codigo).trim() !== '0025') {
      throw new Error(`estado inesperado: ${ORIGINAL}=${dO.codigo}, ${COPIA}=${dC.codigo} -- nada feito`);
    }
    const conteudoO = Object.assign({}, dO), conteudoC = Object.assign({}, dC);
    Object.keys(dO).forEach(k => delete dO[k]);
    Object.keys(dC).forEach(k => delete dC[k]);
    Object.assign(dO, conteudoC, { id: ORIGINAL, codigo: '0025' });
    Object.assign(dC, conteudoO, { id: COPIA });
    console.log(`  ${ORIGINAL}: ${dO.codigo} · ${dO.desc} · ${dO.img} · ${dO.skuLinha}`);
    console.log(`  ${COPIA}: ${dC.codigo} · ${dC.desc} · ${dC.img} · ${dC.skuLinha}`);
  } else console.log('  cadastros ja trocados');

  ordens.forEach(o => {
    if (OS_DO_AZUL.includes(o.os) && o.desenhoId === ORIGINAL) {
      o.desenhoId = COPIA;
      o.codigo = dC.codigo;
      delete o.desenhoFoto;
      console.log(`  OS ${o.os}: desenho -> ${dC.codigo}`);
    }
  });

  const em = new Date().toISOString();
  const porId = new Map(desenhos.map(d => [d.id, d]));
  let fotos = 0;
  ordens.forEach(o => {
    const d = o.desenhoId && porId.get(o.desenhoId);
    if (!d || (o.desenhoFoto && o.desenhoFoto.id === d.id)) return;
    o.desenhoFoto = foto(d, em);
    fotos++;
  });
  console.log(`  foto do desenho gravada em ${fotos} OS`);
  const doOriginal = ordens.filter(o => o.desenhoId === ORIGINAL).map(o => o.os);
  console.log(`  OS no 0025: ${doOriginal.join(', ')}`);
  console.log(`  OS no 0050: ${ordens.filter(o => o.desenhoId === COPIA).map(o => o.os).join(', ')}`);

  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-desenho-0025-' + em.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, { desenhos: JSON.stringify(desenhos), ordens: JSON.stringify(ordens) });
  const volta = JSON.parse(await req('PATCH',
    `/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(linha.updated_at)}`,
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU ENTRE A LEITURA E A ESCRITA -- nada alterado. Rode de novo.');
  const conf = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const o547 = arr(conf.ordens).find(o => o.os === '0547');
  console.log(`gravado. OS 0547 -> ${o547.desenhoFoto && o547.desenhoFoto.codigo} ${o547.desenhoFoto && o547.desenhoFoto.desc}`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
