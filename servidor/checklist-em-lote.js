/* Acerta o CHECKLIST de todas as OS, e marca de uma vez o que já acabou.

   POR QUE ISTO EXISTE
   O checklist de uma OS é uma CÓPIA: quando a OS é emitida, a lista de etapas
   do desenho técnico é copiada para dentro dela (o.etapas), e fica congelada
   ali. Mexer no desenho depois só alcança as OS já emitidas por
   propagarEtapasDesenhoParaOS — que só roda quando ALGUÉM SALVA aquele desenho.
   Desenho que não foi salvo desde a última mudança deixou as OS dele para trás,
   com a lista velha.

   Em 15/09/2026 o STATUS da OS passou a nascer do checklist (Enfestando,
   Cortando, Costurando…), e aí a lista velha deixou de ser detalhe: OS sem a
   caixa "Enfesto" nunca mostra "Enfestando". Daí a primeira tarefa deste
   script — pôr todas as OS com o checklist que o desenho delas tem HOJE.

   E daí a segunda: as OS antigas estão todas com o checklist em branco, porque
   nasceram antes de alguém marcar caixa nenhuma. Elas já acabaram há meses, e
   em branco aparecem na tela como se ainda estivessem no corte — enchendo os
   campos do fluxo e o dashboard de trabalho que não existe. Marcar as caixas de
   uma faixa de uma vez é o que a tela não sabe fazer.

   Roda NO SERVIDOR da fábrica. Lê a chave de serviço do .env do Supabase local,
   lê o shared_data, mexe e grava de volta — a mesma conta read-modify-write da
   restauração credenciada e do status-os-em-lote.

   O QUE ELE NÃO FAZ
   Não inventa etapa: se o desenho da OS não tem lista de etapas (ou a OS não
   tem desenho), o checklist dela fica como está.

   E POR PADRÃO NÃO REMOVE etapa que já tem marca — ela vai para o fim da lista,
   exatamente como _etapasFinaisOS faz na tela, porque tirar da folha um trabalho
   que alguém registrou é como checklist "some". Só que essa mesma regra emperra
   a faxina de nome SUBSTITUÍDO: "Costura CM.LISA" virou "Costura CM.LISA |
   Descalvado" e "| São Carlos", e a OS antiga fica carregando o nome velho
   marcado para sempre, mostrando as duas versões da mesma etapa na folha. Daí o
   --forcar, que põe a lista do desenho e só ela — levando a MARCA do nome velho
   para o novo (migrarMarcas), e nomeando no relatório o que não teve para onde
   ir. Rode a conferência antes: é ela que diz quais OS RECENTES perderiam
   apontamento.

   Antes de gravar, o blob inteiro é copiado para backups/ — é a volta atrás.

   USO (confere primeiro, grava depois):
     node servidor/checklist-em-lote.js --etapas
     node servidor/checklist-em-lote.js --marcar-antes 2026-09-01
     node servidor/checklist-em-lote.js --etapas --forcar --marcar-antes 2026-09-01 --fim "Estoque"
     node servidor/checklist-em-lote.js --etapas --forcar --marcar-antes 2026-09-01 --fim "Estoque" --gravar
*/
const fs = require('fs');
const path = require('path');
const https = require('https');

const args = process.argv.slice(2);
const opt = (n) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : null; };
const ETAPAS = args.includes('--etapas');
const ANTES = (opt('--marcar-antes') || '').trim();
// A etapa que vai para o FIM da lista das OS marcadas (criada se faltar). Sem
// ela, vale a ordem que o desenho deu — ver o bloco em que FIM e usado.
const FIM = (opt('--fim') || '').trim();
// --forcar: a lista da OS vira a do desenho E MAIS NADA — ver etapasFinais.
const FORCAR = args.includes('--forcar');
const GRAVAR = args.includes('--gravar');
const BASE = opt('--url') || 'https://193.168.0.200';
const ENV = opt('--env') || 'C:\\supabase\\docker\\.env';

if (!ETAPAS && !ANTES) {
  console.error('nada a fazer. Use --etapas e/ou --marcar-antes AAAA-MM-DD');
  process.exit(1);
}
if (ANTES && !/^\d{4}-\d{2}-\d{2}$/.test(ANTES)) {
  console.error('--marcar-antes precisa ser uma data AAAA-MM-DD (ex.: 2026-09-01)');
  process.exit(1);
}

const envTxt = fs.readFileSync(ENV, 'utf8');
const mKey = envTxt.match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m);
if (!mKey) { console.error('não achei SERVICE_ROLE_KEY em ' + ENV); process.exit(1); }
const KEY = mKey[1].trim();

// O certificado é o da própria fábrica; este script fala com o servidor local.
const agente = new https.Agent({ rejectUnauthorized: false });

function req(method, caminho, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const u = new URL(BASE + caminho);
    const r = https.request({
      hostname: u.hostname, port: u.port || 443, path: u.pathname + u.search, method, agent: agente,
      headers: {
        apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json',
        Prefer: method === 'PATCH' ? 'return=minimal' : '',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, (resp) => {
      let d = ''; resp.on('data', c => d += c);
      resp.on('end', () => resp.statusCode < 300 ? res({ status: resp.statusCode, body: d })
                                                 : rej(new Error('HTTP ' + resp.statusCode + ': ' + d.slice(0, 300))));
    });
    r.on('error', rej); if (data) r.write(data); r.end();
  });
}

function pk(v) { return typeof v === 'string' ? JSON.parse(v) : v; }

/* --------- as mesmas regras do app.js, recortadas para cá --------- */

// Etapas que NÃO podem sair da lista: as que já têm marca (a etapa ou qualquer
// tarefa dela). É _etapasComMarcaOS do app.js.
function etapasComMarca(o) {
  const prog = (o && o.progresso) || {};
  const marcadas = new Set();
  Object.entries(prog.etapasCheck || {}).forEach(([n, v]) => { if (v) marcadas.add(n); });
  Object.entries(prog.tarefasCheck || {}).forEach(([n, obj]) => {
    if (obj && typeof obj === 'object' && Object.values(obj).some(Boolean)) marcadas.add(n);
  });
  return marcadas;
}

/* Lista final ao receber as etapas do desenho: as do desenho, mais as que já têm
   marca e sumiram de lá — estas vão para o fim, nunca somem. É _etapasFinaisOS.

   --forcar: a lista passa a ser a do desenho E MAIS NADA, inclusive jogando fora
   etapa marcada que não está lá.

   Por que existe: a regra de preservar é certa para o dia a dia — tirar da folha
   uma etapa que alguém marcou apaga trabalho registrado, e é assim que checklist
   "some". Mas ela emperra a faxina de nomes SUBSTITUÍDOS. Quando "Costura
   CM.LISA" virou "Costura CM.LISA | Descalvado" e "| São Carlos", e "Expedição"
   virou as duas direcionais, as OS antigas ficaram carregando o nome velho
   marcado — e ele nunca mais sai sozinho, porque tem marca. A folha passa a
   mostrar as duas versões da mesma etapa, uma marcada e outra não.

   Aqui o trabalho não se perde: o nome novo está na lista do desenho e recebe a
   marca no passo 2. Só use --forcar sabendo disso, e olhando antes o relatório
   de quais OS perdem marca sem ganhar a equivalente. */
function etapasFinais(o, novas, forcar) {
  if (forcar) return [...novas];
  const marcadas = etapasComMarca(o);
  const preservar = (o.etapas || []).filter(e => marcadas.has(e) && !novas.includes(e));
  return [...novas, ...preservar];
}

/* A ETAPA NOVA QUE SUBSTITUI UMA VELHA, quando o nome foi desdobrado.

   Os nomes velhos não foram apagados: foram PARTIDOS. "Costura CM.LISA" virou
   "Costura CM.LISA | Descalvado" e "Costura CM.LISA | São Carlos"; "Expedição"
   virou "Expedição Desc X São Carlos" e "Expedição São Carlos X Desc.". A marca
   que está no nome velho é trabalho que aconteceu de verdade, e jogá-la fora
   junto com o nome seria apagar apontamento de OS que ainda está em produção.

   Candidato = etapa nova cujo nome COMEÇA com o nome velho. Quando sobram as
   duas metades (uma de cada unidade), quem desempata é a caixa "Recebido em São
   Carlos" — a mesma regra que separa as duas costuras no status e nos campos do
   fluxo. Sem candidato claro, devolve null e quem chama avisa. */
function equivalenteNova(velho, novas, o) {
  const cand = novas.filter(n => n !== velho && n.startsWith(velho));
  if (!cand.length) return null;
  if (cand.length === 1) return cand[0];
  const sc = cand.filter(n => /s[ãa]o carlos/i.test(n));
  const desc = cand.filter(n => !/s[ãa]o carlos/i.test(n));
  if (sc.length === 1 && desc.length === 1) {
    const recebida = ((o.etapas || []).some(n => /recebido em s[ãa]o carlos/i.test(n)
      && ((o.progresso || {}).etapasCheck || {})[n]));
    return recebida ? sc[0] : desc[0];
  }
  return null;
}

/* Leva a marca (e as tarefas) do nome velho para o nome novo. Devolve os nomes
   que não tiveram para onde ir — são os que a folha realmente perde, e quem
   chama os põe no relatório em vez de deixá-los sumir calados. */
function migrarMarcas(o, atual, novas) {
  const prog = o.progresso || (o.progresso = {});
  const ck = prog.etapasCheck || (prog.etapasCheck = {});
  const sq = prog.etapasSeq || (prog.etapasSeq = {});
  const tk = prog.tarefasCheck || (prog.tarefasCheck = {});
  const orfas = [];
  atual.forEach(velho => {
    if (novas.includes(velho)) return;
    const temMarca = !!ck[velho]
      || (tk[velho] && Object.values(tk[velho]).some(Boolean));
    if (!temMarca) { delete ck[velho]; delete sq[velho]; delete tk[velho]; return; }
    const nova = equivalenteNova(velho, novas, o);
    if (!nova) { orfas.push(velho); delete ck[velho]; delete sq[velho]; delete tk[velho]; return; }
    // Não pisa numa marca que já existe no nome novo: a mais ANTIGA é a que
    // conta como "quando aquela etapa aconteceu".
    if (ck[velho]) {
      const seqVelho = Number(sq[velho]) || 0;
      if (!ck[nova] || (Number(sq[nova]) || 0) > seqVelho) sq[nova] = seqVelho || sq[nova];
      ck[nova] = true;
    }
    if (tk[velho]) {
      const destino = tk[nova] = tk[nova] || {};
      Object.entries(tk[velho]).forEach(([t, v]) => { if (v) destino[t] = true; });
    }
    delete ck[velho]; delete sq[velho]; delete tk[velho];
  });
  return orfas;
}

// As tarefas cadastradas de uma etapa. É tarefasDaEtapa do app.js.
function tarefasDaEtapa(etapa, tarefas) {
  if (Array.isArray(etapa && etapa.tarefas) && etapa.tarefas.length) return etapa.tarefas;
  if (Array.isArray(etapa && etapa.tarefasIds) && etapa.tarefasIds.length) {
    return etapa.tarefasIds.map(tid => (tarefas || []).find(t => t.id === tid)).filter(Boolean);
  }
  return [];
}

// O viés não entra no checklist de Corte: não é enfesto com camadas, é tira
// cortada em diagonal. Mesma regra de _ehFaseVies.
function ehFaseVies(nome) {
  const n = String(nome || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  return !!n && n.split(/\s+/).some(p => p === 'vies' || p === 'vieses');
}

/* As caixas FILHAS de uma etapa, como a folha as desenha:
     - as tarefas do cadastro daquela etapa;
     - na etapa de CORTE, uma linha por fase do enfesto da OS ("Fase 1", "Fase
       2"…), que não vêm do cadastro e sim da grade;
     - as que já têm marca nesta OS, mesmo fora do cadastro (tarefa renomeada ou
       excluída depois) — a folha as resgata, e apagá-las aqui esconderia
       trabalho registrado. */
function tarefasDaFolha(o, nomeEtapa, etapasCad, tarefasCad) {
  const nomes = new Set();
  const cad = (etapasCad || []).find(e => e.nome === nomeEtapa);
  tarefasDaEtapa(cad, tarefasCad).forEach(t => { if (t && t.nome) nomes.add(t.nome); });
  if (/corte/i.test(nomeEtapa)) {
    (o.fases || []).forEach((f, i) => {
      if (ehFaseVies(f && f.nome)) return;
      const ordem = Number(f && f.ordem) || (i + 1);
      nomes.add('Fase ' + ordem);
    });
  }
  const jaMarcadas = ((o.progresso || {}).tarefasCheck || {})[nomeEtapa] || {};
  Object.entries(jaMarcadas).forEach(([n, v]) => { if (v) nomes.add(n); });
  return [...nomes];
}

/* --------- a rodada --------- */

(async () => {
  const rows = JSON.parse((await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at')).body);
  if (!rows.length) { console.error('shared_data id=main não encontrado'); process.exit(1); }
  const data = rows[0].data || {};
  const eraString = {};
  const ler = (k) => { eraString[k] = typeof data[k] === 'string'; return pk(data[k]) || []; };
  const ordens = ler('ordens');
  const desenhos = pk(data.desenhos) || [];
  const etapasCad = pk(data.etapas) || [];
  const tarefasCad = pk(data.tarefas) || [];

  console.log(`shared_data lido (updated_at ${rows[0].updated_at})`);
  console.log(`OS: ${ordens.length}  ·  desenhos: ${desenhos.length}  ·  etapas cadastradas: ${etapasCad.length}\n`);

  /* ---- 1. o checklist de cada OS, vindo do desenho dela ---- */
  let mudouEtapas = 0;
  let criadas = 0;
  const semDesenho = [];
  const semEtapasNoDesenho = new Set();
  const quaisEtapas = [];
  const perdeuRecente = [], perdeuAntiga = [];
  if (ETAPAS) {
    ordens.forEach(o => {
      const cod = String(o.codigo || '').trim();
      const d = (o.desenhoId && desenhos.find(x => x.id === o.desenhoId))
        || (cod && desenhos.find(x => String(x.codigo || '').trim() === cod))
        || null;
      if (!d) { semDesenho.push(o.os || '?'); return; }
      const novas = (Array.isArray(d.etapasNomes) ? d.etapasNomes : []).filter(Boolean);
      if (!novas.length) { semEtapasNoDesenho.add(d.codigo || d.id); return; }
      const final = etapasFinais(o, novas, FORCAR);
      const atual = o.etapas || [];
      if (atual.length === final.length && atual.every((e, i) => e === final[i])) return;
      /* A MARCA DO NOME VELHO MUDA DE CASA. Etapa marcada que sai da lista é
         trabalho registrado; migrarMarcas a leva para o nome que a substituiu
         ("Costura CM.LISA" → "Costura CM.LISA | Descalvado"). O que sobra sem
         equivalente sai nomeado no relatório — nas OS antigas não importa (o
         passo 2 remarca a lista nova inteira), nas RECENTES importa e muito. */
      let orfas = [];
      if (FORCAR) orfas = migrarMarcas(o, atual, final);
      if (orfas.length) {
        const recente = !ANTES || String(o.data || '') >= ANTES;
        (recente ? perdeuRecente : perdeuAntiga).push(`${o.os || '?'} (${o.data || 's/ data'}): ${orfas.join(', ')}`);
      }
      quaisEtapas.push(`${o.os || '?'}: ${atual.length} → ${final.length} etapas`);
      o.etapas = final;
      mudouEtapas++;
    });
    console.log('--- 1. checklist vindo do desenho ---');
    console.log(`OS com o checklist reescrito: ${mudouEtapas}`);
    if (quaisEtapas.length <= 12) quaisEtapas.forEach(l => console.log('   ' + l));
    if (perdeuAntiga.length) {
      console.log(`OS antigas com marca sem equivalente: ${perdeuAntiga.length}`
        + ' — sem problema, o passo 2 remarca a lista nova inteira.');
    }
    if (perdeuRecente.length) {
      console.log(`
  ⚠  OS RECENTES com marca SEM equivalente (a folha perde o apontamento): ${perdeuRecente.length}`);
      perdeuRecente.slice(0, 40).forEach(l => console.log('     ' + l));
    }
    if (semDesenho.length) {
      console.log(`OS sem desenho encontrado (ficam como estão): ${semDesenho.length}`
        + (semDesenho.length <= 20 ? ' — ' + semDesenho.join(', ') : ''));
    }
    if (semEtapasNoDesenho.size) {
      console.log(`desenhos SEM lista de etapas (as OS deles ficam como estão): ${semEtapasNoDesenho.size}`
        + ' — ' + [...semEtapasNoDesenho].slice(0, 20).join(', '));
    }
    console.log('');
  }

  /* ---- 2. marcar tudo nas OS anteriores à data ---- */
  let marcadas = 0, caixasEtapa = 0, caixasTarefa = 0;
  const semData = [];
  const porUltimaEtapa = new Map();
  if (ANTES) {
    const agora = Date.now();
    ordens.forEach(o => {
      const d = String(o.data || '');
      if (!/^\d{4}-\d{2}-\d{2}/.test(d)) { semData.push(o.os || '?'); return; }
      if (d >= ANTES) return;
      const etapas = (o.etapas || []).filter(Boolean);
      if (!etapas.length) return;
      /* --fim <etapa>: esta etapa é carimbada POR ÚLTIMO, onde quer que ela
         esteja na lista impressa.

         Ela é necessária porque a ordem da lista não é a ordem do relógio. A
         lista é a do desenho técnico, e em 152 das 257 OS antigas "Estoque"
         aparece no MEIO dela — a última é "Costura" ou "Expedição". Quem manda
         no status e no campo do fluxo é a etapa de maior `etapasSeq`, então
         carimbar na ordem da lista poria 152 OS acabadas há meses dentro de
         Costurando e Expedição, como se estivessem em produção agora.

         O carimbo muda; a LISTA não. Reordenar o checklist impresso de algumas
         OS e não de outras faria a folha de duas OS do mesmo desenho sair
         diferente, e a ordem impressa é escolha de quem cadastrou o desenho —
         não é assunto deste script. Só se a etapa não existir na lista ela é
         acrescentada, no fim. */
      const iFim = FIM ? etapas.findIndex(n => n === FIM) : -1;
      if (FIM && iFim < 0) { etapas.push(FIM); o.etapas = etapas; criadas++; }
      o.progresso = o.progresso || {};
      o.progresso.etapasCheck = o.progresso.etapasCheck || {};
      o.progresso.etapasSeq = o.progresso.etapasSeq || {};
      o.progresso.tarefasCheck = o.progresso.tarefasCheck || {};
      let mexeu = false;
      /* O CARIMBO DE ORDEM SOBE A CADA ETAPA, e isso não é detalhe: o status da
         OS e o campo do fluxo saem da etapa de MAIOR `etapasSeq`. Marcando tudo
         com o mesmo instante, o desempate cairia na primeira da lista e as OS
         antigas iriam todas parar no Estoque de corte — o oposto do que se quer.

         O RELÓGIO É O DIA DA PRÓPRIA OS, e não o da rodada. `etapasSeq` é o que
         a coluna Data lê para dizer quando a OS terminou: carimbando tudo com
         `Date.now()`, as 257 OS antigas passariam a dizer que terminaram nesta
         tarde, todas juntas — e o filtro "Finalizada em" perderia o histórico
         inteiro. Com o dia da OS, cada uma guarda a sua data. Não é a hora real
         em que ela terminou (essa ninguém anotou), mas é a única data que o
         script sabe que pertence àquela OS. */
      const base = Date.parse(d.slice(0, 10) + 'T12:00:00') || agora;
      etapas.forEach((nome, i) => {
        if (!o.progresso.etapasCheck[nome]) { caixasEtapa++; mexeu = true; }
        o.progresso.etapasCheck[nome] = true;
        // A etapa de --fim leva o maior carimbo do lote, onde quer que esteja.
        o.progresso.etapasSeq[nome] = (FIM && nome === FIM) ? base + etapas.length : base + i;
        const filhas = tarefasDaFolha(o, nome, etapasCad, tarefasCad);
        if (filhas.length) {
          const mapa = o.progresso.tarefasCheck[nome] = o.progresso.tarefasCheck[nome] || {};
          filhas.forEach(t => {
            if (!mapa[t]) { caixasTarefa++; mexeu = true; }
            mapa[t] = true;
          });
        }
      });
      if (mexeu) marcadas++;
      const ultima = FIM || etapas[etapas.length - 1];
      porUltimaEtapa.set(ultima, (porUltimaEtapa.get(ultima) || 0) + 1);
    });
    console.log(`--- 2. marcar tudo nas OS anteriores a ${ANTES} ---`);
    console.log(`OS marcadas: ${marcadas}  ·  caixas de etapa: ${caixasEtapa}  ·  caixas de tarefa: ${caixasTarefa}`);
    if (FIM) console.log(`"${FIM}" carimbada por ultimo em todas elas`
      + (criadas ? ` (e acrescentada a lista de ${criadas} OS que nao a tinham)` : ' (todas ja a tinham na lista)'));
    if (semData.length) {
      console.log(`OS sem data (ficam de fora — não dá para dizer se são anteriores): ${semData.length}`
        + (semData.length <= 20 ? ' — ' + semData.join(', ') : ''));
    }
    /* ONDE ELAS VÃO PARAR. A última etapa da folha manda no status e no campo do
       fluxo. Se ela não for "Estoque", a OS não sai do fluxo — fica no campo
       daquela etapa. É o número que se confere antes de gravar. */
    console.log('última etapa de cada uma (é ela que decide onde a OS fica):');
    [...porUltimaEtapa.entries()].sort((a, b) => b[1] - a[1])
      .forEach(([nome, n]) => console.log(`   ${String(n).padStart(4)} × ${nome}`));
    /* A CAIXA "ESTOQUE" é a que tira a OS do fluxo em processo. Quem não a tem
       na lista não sai — fica para sempre no campo da última etapa que tem. Este
       balanço diz de quantas se trata ANTES de gravar, que é quando ainda dá
       para decidir usar --fim. */
    let temEstoqueUlt = 0, temEstoqueNoMeio = 0, semEstoque = 0;
    const alvo = ordens.filter(o => {
      const d = String(o.data || '');
      return /^\d{4}-\d{2}-\d{2}/.test(d) && d < ANTES && (o.etapas || []).length;
    });
    alvo.forEach(o => {
      const e = o.etapas || [];
      const i = e.findIndex(n => /estoque/i.test(n));
      if (i < 0) semEstoque++;
      else if (i === e.length - 1) temEstoqueUlt++;
      else temEstoqueNoMeio++;
    });
    console.log(`\ncaixa "Estoque" na lista: ${temEstoqueUlt} a têm por último`
      + ` · ${temEstoqueNoMeio} a têm no meio · ${semEstoque} não a têm`);
    if (!FIM && temEstoqueNoMeio + semEstoque > 0) {
      console.log(`  → ATENÇÃO: ${temEstoqueNoMeio + semEstoque} OS NÃO sairiam do fluxo em processo:`
        + ' ficariam nos campos de Costura/Expedição como se ainda estivessem em produção.');
      console.log('  → --fim "Estoque" resolve: carimba essa caixa por último onde quer que ela'
        + ' esteja na lista, e é o carimbo que manda no status e no campo.');
    }
    console.log('');
  }

  const total = mudouEtapas + marcadas + criadas;
  if (!GRAVAR) {
    console.log('(conferência — NADA foi gravado. Rode de novo com --gravar para aplicar.)');
    return;
  }
  if (!total) { console.log('nada a gravar.'); return; }

  // Cópia do blob ANTES de mexer: é a volta atrás se a rodada sair errada.
  const quando = new Date().toISOString();
  const dir = path.join(__dirname, '..', 'backups');
  fs.mkdirSync(dir, { recursive: true });
  const arq = path.join(dir, 'shared_data-antes-checklist-' + quando.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(rows[0]), 'utf8');
  console.log('cópia de segurança: ' + arq);

  data.ordens = eraString.ordens ? JSON.stringify(ordens) : ordens;
  data._device = 'checklist-em-lote-' + Date.now();   // sentinela: todos os clientes recarregam
  await req('PATCH', '/rest/v1/shared_data?id=eq.main', { data, updated_at: new Date().toISOString() });
  console.log(`\n✅ gravado. Os clientes conectados recarregam sozinhos.`);
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
