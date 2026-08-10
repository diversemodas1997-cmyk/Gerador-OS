/* Descobre qual arquivo de imagem pertence a cada desenho técnico.

   O PAREAMENTO É PELO SKU, e não pela descrição. O SKU é a chave que o próprio
   cadastro já usa; a descrição é texto livre e leva a erro — "Camiseta Tricolor
   Preto/Caqui/Off-White" e "Camiseta Tricolor Preto/Grafite/Branco" começam
   pela mesma cor e disputariam o mesmo arquivo. Pelo SKU (CM.TRI.LISA-CAQUI e
   CM.TRI.LISA-PRE) não há dúvida.

   Formato esperado: <PREFIXO>-<COR>, com as pastas nomeadas pelo prefixo.
     CM.LISA-PRE        -> pasta CM.LISA, arquivo cuja primeira palavra
                           comece por "PRE"  (PRETO.png)
     CM.REC.LISA-VERM   -> pasta CM.REC   (a pasta é o maior prefixo que casa)
     BM.TRI-BEGE        -> pasta BM.TRI

   NA DÚVIDA, NÃO PAREIA. Um desenho sem imagem é um incômodo; o desenho ERRADO
   numa OS de corte é tecido perdido. Zero candidatos ou mais de um viram
   pendência para alguém decidir, nunca um palpite. */

function normalizar(s) {
  return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
}

/**
 * @param {Array}  desenhos       registros do cadastro (codigo, skuLinha, desc)
 * @param {Object} porPasta       { 'CM.LISA': ['PRETO.png', ...], ... }
 * @param {Object} mapaManual     { codigo: 'PASTA/ARQUIVO.png' } — tem prioridade
 * @returns {{pares: Map, pendencias: Array, sobrando: Array}}
 */
function parear(desenhos, porPasta, mapaManual) {
  const manual = mapaManual || {};
  const pastas = Object.keys(porPasta).sort((a, b) => b.length - a.length);
  const pares = new Map();
  const pendencias = [];

  for (const d of desenhos) {
    const cod = String(d.codigo || '').trim();
    const desc = d.desc || '';
    if (!cod) continue;

    if (manual[cod]) { pares.set(cod, { arquivo: manual[cod], origem: 'manual' }); continue; }

    const sku = normalizar(d.skuLinha);
    if (!sku) {
      pendencias.push({ cod, desc, motivo: 'sem SKU no cadastro — preencha o SKU do desenho no app, ou use --mapa' });
      continue;
    }
    const corte = sku.lastIndexOf('-');
    if (corte < 1) {
      pendencias.push({ cod, desc, motivo: `SKU "${sku}" não tem o formato PREFIXO-COR` });
      continue;
    }
    const prefixo = sku.slice(0, corte), cor = sku.slice(corte + 1);
    // A pasta é o MAIOR prefixo que casa: CM.REC.LISA-… mora em CM.REC, e
    // testar as maiores primeiro evita que CM.TRI caia numa pasta CM.
    const pasta = pastas.find(p => prefixo.startsWith(normalizar(p)));
    if (!pasta) {
      pendencias.push({ cod, desc, motivo: `nenhuma pasta corresponde ao prefixo "${prefixo}" do SKU` });
      continue;
    }
    // A cor casa quando alguma PALAVRA do nome do arquivo começa por ela:
    // "MARINHO" acha "AZUL MARINHO.png"; "CAQUI" acha "TRICOLOR CAQUI.png".
    const achados = (porPasta[pasta] || []).filter(f => {
      const base = normalizar(f.replace(/\.[a-z0-9]+$/i, ''));
      return base.split(/\s+/).some(p => p.startsWith(cor));
    });
    if (achados.length === 1) {
      pares.set(cod, { arquivo: pasta + '/' + achados[0], origem: 'sku' });
    } else if (achados.length === 0) {
      pendencias.push({ cod, desc, motivo: `nenhum arquivo em ${pasta} para a cor "${cor}"` });
    } else {
      pendencias.push({ cod, desc, motivo: `${achados.length} arquivos em ${pasta} para a cor "${cor}": ${achados.join(', ')}` });
    }
  }

  const usados = new Set([...pares.values()].map(v => v.arquivo));
  const sobrando = [];
  for (const p of Object.keys(porPasta)) {
    for (const f of porPasta[p]) if (!usados.has(p + '/' + f)) sobrando.push(p + '/' + f);
  }
  return { pares, pendencias, sobrando: sobrando.sort() };
}

module.exports = { parear, normalizar };
