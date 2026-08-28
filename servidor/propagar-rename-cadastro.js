/*
 * Leva um RENOME do cadastro (tecido ou cor) para os registros que guardaram o
 * nome antigo.
 *
 * POR QUE ISTO EXISTE
 *
 * A OS e o razao do estoque guardam o nome do pano e da cor como RETRATO, ao
 * lado do id — de proposito: a folha impressa de uma OS de marco tem de
 * continuar dizendo o que foi cortado em marco, mesmo que o cadastro mude
 * depois. O preco e que um renome no cadastro nao alcanca o que ja foi gravado.
 *
 * Em 28/08/2026 o Junior renomeou `Texturizado | Jaguar` para
 * `Texturizado | Jacguar`, para bater com as cores ("Preto Jacguar",
 * "Off-white Jacguar") e com os desenhos 0032/0033. O id ficou o mesmo, entao
 * nada quebrou — mas 18 retratos seguiram com a grafia velha, dois deles no
 * estoque, onde a linha e o que a pessoa le para decidir compra.
 *
 * COMO ELE SABE O QUE MUDOU
 *
 * Nao ha nome escrito aqui dentro. Ele COMPARA: onde um registro carrega um id
 * e um nome, o id manda — se o cadastro daquele id diz outro nome, o retrato
 * esta velho. Dessas divergencias sai o de-para (nome velho -> nome novo), e e
 * ele que conserta tambem os registros que so tem NOME e nenhum id: o razao do
 * estoque e o `nomeCor` dos blocos de enfesto.
 *
 * Por isso o script serve para qualquer renome futuro, nao so para este.
 *
 * O QUE ELE NAO FAZ
 *
 * De-para ambiguo (dois ids diferentes com o mesmo nome velho) e RELATADO e
 * pulado: escolher no escuro poe o pano de um no lugar do outro.
 *
 * E NAO GRAVA NADA SEM `--so`. O de-para inteiro pega tambem o desdobramento
 * antigo das cores ("Preto" -> "Preto Malha Algodao", 739 componentes), e
 * aquilo NAO deve ser reescrito: `movimentacoesEstoque` converte a cor pura em
 * tempo de LEITURA, de proposito — "nao reescreve nada no banco, e desfazer e
 * so reverter o codigo". Passar por cima de 739 registros historicos para
 * repetir o que a leitura ja faz e o oposto disso.
 *
 * COMO RODAR
 *
 *   node servidor/propagar-rename-cadastro.js               relata TODO o de-para
 *   node servidor/propagar-rename-cadastro.js --so Jaguar   so as linhas com esse texto
 *   node servidor/propagar-rename-cadastro.js --so Jaguar --gravar
 *
 * Antes de gravar salva o blob em backups/, e a escrita confere o carimbo.
 *
 * ATENCAO: recarregue a pagina (F5) depois de rodar. Aba aberta desde antes
 * regrava a chave inteira com o estado velho e desfaz isto.
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
// Recorte do de-para: so as linhas cujo nome VELHO contem este texto. Sem ele o
// script relata tudo e se recusa a gravar.
const SO = (() => { const i = process.argv.indexOf('--so'); return i > 0 ? process.argv[i + 1] : ''; })();

const vazio = v => v == null || String(v).trim() === '';

async function lerBlob() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  return { data: linhas[0].data, updatedAt: linhas[0].updated_at, cab };
}

async function gravarBlob(cab, data, updatedAt) {
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
}

(async () => {
  const { data, updatedAt, cab } = await lerBlob();
  const le = k => JSON.parse(data[k] || '[]');
  const ordens = le('ordens'), estoqueMov = le('estoqueMov');
  const nomeDoId = new Map();
  le('tecidos').forEach(t => nomeDoId.set('T:' + t.id, t.nome));
  le('cores').forEach(c => nomeDoId.set('C:' + c.id, c.nome));

  /* Passo 1: onde ha id E nome, o id manda. Cada divergencia vira uma linha do
     de-para e o retrato e corrigido na hora. */
  const dePara = new Map();     // 'T:nome velho' -> Set de nomes novos
  const anota = (tipo, velho, novo) => {
    const k = tipo + ':' + velho;
    (dePara.get(k) || dePara.set(k, new Set()).get(k)).add(novo);
  };
  const contas = { fases: 0, linhas: 0, componentes: 0, blocos: 0, estoque: 0 };

  // O passo 1 corrige no lugar para descobrir o de-para; o que ficar fora do
  // recorte e desfeito depois, pela lista `pendentes`.
  const pendentes = [];
  const corrige = (obj, tipo, idProp, nomeProp, prefixoId) => {
    const idBruto = obj[idProp];
    if (vazio(idBruto)) return false;
    const chave = prefixoId ? String(idBruto) : tipo + ':' + idBruto;
    const novo = nomeDoId.get(prefixoId ? String(idBruto) : chave);
    if (!novo || vazio(obj[nomeProp]) || obj[nomeProp] === novo) return false;
    anota(tipo, obj[nomeProp], novo);
    pendentes.push({ obj, nomeProp, tipo, velho: obj[nomeProp], novo });
    obj[nomeProp] = novo;
    return true;
  };

  ordens.forEach(o => {
    (o.fases || []).forEach(f => {
      if (corrige(f, 'T', 'tecidoId', 'tecidoNome')) pendentes[pendentes.length - 1].onde = 'fases';
      if (corrige(f, 'C', 'corId', 'corNome')) pendentes[pendentes.length - 1].onde = 'fases';
    });
    (o.tecidos || []).forEach(t => {
      if (corrige(t, 'T', 'tecidoId', 'tecidoNome')) pendentes[pendentes.length - 1].onde = 'linhas';
      if (corrige(t, 'C', 'corId', 'corNome')) pendentes[pendentes.length - 1].onde = 'linhas';
    });
    // O componente guarda o tecido como "T:<id>" no campo `material` — o
    // prefixo ja vem no valor, e por isso a chave nao e remontada aqui.
    (o.componentes || []).forEach(c => {
      if (corrige(c, 'T', 'material', 'materialNome', true)) pendentes[pendentes.length - 1].onde = 'componentes';
      if (corrige(c, 'C', 'cor', 'corNome')) pendentes[pendentes.length - 1].onde = 'componentes';
    });
  });

  // De-para ambiguo: dois ids com o mesmo nome velho. Nao da para saber qual
  // era qual no registro que so tem nome.
  const ambiguos = [...dePara.entries()].filter(([, s]) => s.size > 1);
  ambiguos.forEach(([k, s]) => console.log('  AMBIGUO, pulado: "' + k.slice(2)
    + '" virou ' + [...s].map(x => '"' + x + '"').join(' e ')));
  const todas = [...dePara.entries()].filter(([, s]) => s.size === 1)
    .map(([k, s]) => [k, [...s][0]]);
  const escolhidas = SO
    ? todas.filter(([k, v]) => (k.slice(2) + ' ' + v).toLowerCase().includes(SO.toLowerCase()))
    : [];
  const fora = todas.length - escolhidas.length;
  const mapa = new Map(escolhidas);

  // Devolve o nome velho a quem ficou fora do recorte: relatar nao e gravar.
  pendentes.forEach(p => { if (!mapa.has(p.tipo + ':' + p.velho)) p.obj[p.nomeProp] = p.velho; });
  contas.fases = contas.linhas = contas.componentes = 0;
  pendentes.forEach(p => { if (mapa.has(p.tipo + ':' + p.velho)) contas[p.onde] = (contas[p.onde] || 0) + 1; });

  /* Passo 2: os registros que so tem NOME — o razao do estoque e o retrato da
     cor nos blocos de enfesto. Sem o de-para do passo 1 nao havia como saber
     que "Texturizado | Jaguar" e o mesmo pano de hoje. */
  estoqueMov.forEach(m => {
    const nt = mapa.get('T:' + m.tecidoNome);
    const nc = mapa.get('C:' + m.corNome);
    if (nt) { m.tecidoNome = nt; contas.estoque++; }
    if (nc) { m.corNome = nc; contas.estoque++; }
  });
  ordens.forEach(o => {
    ((o.enfesto || {}).blocos || []).forEach(b => {
      const nc = mapa.get('C:' + b.nomeCor);
      if (nc) { b.nomeCor = nc; contas.blocos++; }
    });
  });

  console.log('');
  if (!todas.length && !ambiguos.length) { console.log('Nenhum retrato desatualizado.'); return; }
  console.log('DE-PARA encontrado (o cadastro mudou, o retrato ficou):');
  todas.sort().forEach(([k, v]) => {
    const dentro = mapa.has(k);
    console.log('  ' + (dentro ? '>>' : '  ') + ' ' + (k[0] === 'T' ? 'tecido' : 'cor  ')
      + '  "' + k.slice(2) + '"  ->  "' + v + '"' + (dentro ? '' : '   (fora do recorte)'));
  });
  if (!SO) {
    console.log('');
    console.log('SEM --so nada e gravado. Escolha o renome, ex.:  --so Jaguar');
    console.log('O desdobramento antigo das cores ("Preto" -> "Preto Malha Algodao") NAO deve');
    console.log('ser reescrito: movimentacoesEstoque ja converte isso em tempo de leitura.');
    return;
  }
  if (fora) {
    console.log('');
    console.log('  (' + fora + ' outro(s) renome(s) ficaram de fora do recorte "' + SO + '")');
  }
  console.log('');
  console.log('  fases de OS        : ' + contas.fases);
  console.log('  linhas de Tecidos  : ' + contas.linhas);
  console.log('  componentes        : ' + contas.componentes);
  console.log('  blocos de enfesto  : ' + contas.blocos);
  console.log('  movimentos estoque : ' + contas.estoque);

  const total = Object.values(contas).reduce((a, b) => a + b, 0);
  if (!total) { console.log('\nNada a gravar.'); return; }
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-rename-cadastro-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.ordens = JSON.stringify(ordens);
  data.estoqueMov = JSON.stringify(estoqueMov);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + total + ' retratos atualizados.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
