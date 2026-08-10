/* Espelho do servidor da fábrica para a nuvem — SEMPRE de mão única.

   A nuvem é cópia, nunca origem. Nada aqui lê a nuvem para trazer de volta:
   como o app abre a nuvem em modo consulta, nada é editado lá, e por isso não
   existe reconciliação entre dois bancos — que é onde mora o risco de perder
   dados. Se um dia a nuvem passar a aceitar edição, este arquivo deixa de ser
   suficiente e o problema muda de tamanho.

   A lógica fica separada da linha de comando para poder ser testada contra um
   servidor de mentira (ver testes/espelho.js). */

const CAMINHO_BUCKET = 'desenhos';

function cab(chave, extra) {
  return Object.assign({ apikey: chave, Authorization: 'Bearer ' + chave }, extra || {});
}

/* Blob vazio = sem OS e sem desenhos. Mesmo critério do app. */
function blobVazio(data) {
  if (!data || typeof data !== 'object') return true;
  for (const k of ['ordens', 'desenhos']) {
    const v = data[k];
    if (typeof v !== 'string') continue;
    try { if (JSON.parse(v).length > 0) return false; } catch (e) { /* segue */ }
  }
  return true;
}

async function json(buscar, url, opcoes) {
  const r = await buscar(url, opcoes);
  if (!r.ok) throw new Error(`${url} respondeu ${r.status}`);
  const t = await r.text();
  return t ? JSON.parse(t) : null;
}

/* Lista os nomes de um bucket. Devolve Set. */
async function listarBucket(buscar, base, chave) {
  const nomes = new Set();
  for (let pagina = 0; ; pagina++) {
    const lote = await json(buscar, `${base}/storage/v1/object/list/${CAMINHO_BUCKET}`, {
      method: 'POST',
      headers: cab(chave, { 'Content-Type': 'application/json' }),
      body: JSON.stringify({ prefix: '', limit: 100, offset: pagina * 100 })
    });
    if (!Array.isArray(lote) || !lote.length) break;
    lote.forEach(o => { if (o && o.name) nomes.add(o.name); });
    if (lote.length < 100) break;
  }
  return nomes;
}

/**
 * @param {object} cfg  { local, localKey, nuvem, nuvemKey, forcar, semImagens, buscar, log }
 * @returns {object} relatório do que foi feito
 */
async function espelhar(cfg) {
  const buscar = cfg.buscar || fetch;
  const log = cfg.log || (() => {});
  const rel = { dados: 'nao-mexeu', imagens: 0, motivo: null };

  // 1. O que a fábrica tem agora.
  const localLinhas = await json(buscar,
    `${cfg.local}/rest/v1/shared_data?id=eq.main&select=data,updated_at`,
    { headers: cab(cfg.localKey) });
  const daFabrica = Array.isArray(localLinhas) ? localLinhas[0] : null;
  if (!daFabrica || !daFabrica.data) {
    rel.motivo = 'o servidor da fábrica não tem dados — nada a espelhar';
    return rel;
  }

  // 2. O que a nuvem já tem. Só o carimbo: é o bastante para saber se mudou.
  const nuvemLinhas = await json(buscar,
    `${cfg.nuvem}/rest/v1/shared_data?id=eq.main&select=data,updated_at`,
    { headers: cab(cfg.nuvemKey) });
  const naNuvem = Array.isArray(nuvemLinhas) ? nuvemLinhas[0] : null;

  // 3. TRAVA ANTI-APAGAMENTO. Mesma ideia que já protege o app: não deixar um
  //    lado vazio sobrescrever um lado cheio. Se a fábrica estiver sem dados
  //    (banco recém-criado, restauração pela metade, migração que não rodou), a
  //    nuvem é a única cópia boa que resta — e o espelho não pode ser justamente
  //    o que a destrói.
  if (blobVazio(daFabrica.data) && naNuvem && !blobVazio(naNuvem.data)) {
    rel.dados = 'bloqueado';
    rel.motivo = 'a fábrica está sem OS e sem desenhos, mas a nuvem tem dados — '
      + 'espelhar agora apagaria a única cópia boa';
    return rel;
  }

  // 4. Mudou? O carimbo da nuvem é escrito igual ao da fábrica, então iguais
  //    significa que já está espelhado.
  const precisa = cfg.forcar || !naNuvem || naNuvem.updated_at !== daFabrica.updated_at;
  if (precisa) {
    const r = await buscar(`${cfg.nuvem}/rest/v1/shared_data`, {
      method: 'POST',
      headers: cab(cfg.nuvemKey, {
        'Content-Type': 'application/json', 'Prefer': 'resolution=merge-duplicates'
      }),
      body: JSON.stringify({ id: 'main', data: daFabrica.data, updated_at: daFabrica.updated_at })
    });
    if (!r.ok) throw new Error('falha ao gravar os dados na nuvem: ' + r.status + ' ' + await r.text());
    rel.dados = 'espelhado';
    log('dados espelhados (' + daFabrica.updated_at + ')');

    // O sinal vai junto: com ele, quem abrir pela nuvem também baixa só a chave
    // que mudou, em vez do blob inteiro.
    try {
      const sinal = await json(buscar,
        `${cfg.local}/rest/v1/sync_signal?id=eq.main&select=updated_at,key_versions`,
        { headers: cab(cfg.localKey) });
      const s = Array.isArray(sinal) ? sinal[0] : null;
      if (s) {
        await buscar(`${cfg.nuvem}/rest/v1/sync_signal`, {
          method: 'POST',
          headers: cab(cfg.nuvemKey, {
            'Content-Type': 'application/json', 'Prefer': 'resolution=merge-duplicates'
          }),
          body: JSON.stringify({
            id: 'main', updated_at: s.updated_at,
            device_id: 'espelho', key_versions: s.key_versions
          })
        });
      }
    } catch (e) { log('aviso: o sinal não foi espelhado (' + e.message + ')'); }
  } else {
    rel.motivo = 'a nuvem já está com o mesmo carimbo da fábrica';
  }

  // 5. Imagens novas. Comparação por nome; nada é apagado da nuvem — imagem a
  //    mais lá não custa quase nada, imagem a menos quebra uma folha de OS.
  if (!cfg.semImagens) {
    const [naFabrica, jaNaNuvem] = await Promise.all([
      listarBucket(buscar, cfg.local, cfg.localKey),
      listarBucket(buscar, cfg.nuvem, cfg.nuvemKey)
    ]);
    for (const nome of naFabrica) {
      if (jaNaNuvem.has(nome)) continue;
      const img = await buscar(
        `${cfg.local}/storage/v1/object/public/${CAMINHO_BUCKET}/${encodeURIComponent(nome)}`);
      if (!img.ok) { log('aviso: não li a imagem ' + nome); continue; }
      const corpo = Buffer.from(await img.arrayBuffer());
      const env = await buscar(
        `${cfg.nuvem}/storage/v1/object/${CAMINHO_BUCKET}/${encodeURIComponent(nome)}`, {
          method: 'POST',
          headers: cab(cfg.nuvemKey, {
            'Content-Type': img.headers.get('content-type') || 'image/png', 'x-upsert': 'true'
          }),
          body: corpo
        });
      if (env.ok) { rel.imagens++; log('imagem enviada: ' + nome); }
      else log('aviso: falhou ao enviar ' + nome + ' (' + env.status + ')');
    }
  }

  return rel;
}

module.exports = { espelhar, blobVazio };
