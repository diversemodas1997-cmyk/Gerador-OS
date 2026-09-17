/* REFAZ AS OS QUE SUMIRAM DO PROGRAMA E FICARAM SÓ EM PDF NA PASTA.
 *
 * POR QUE ISTO EXISTE (17/09/2026)
 * A 0509 (01/09), a 0536 e a 0537 (10/09) foram criadas, gravadas e imprimiram
 * PDF na pasta do Setor Corte — e depois desapareceram da lista. A causa está
 * consertada no app (ver _mergeListaPorRegistro e o polling); isto aqui é o
 * resgate delas, que nenhum backup tem: os pacotes diários de 01/09 (16:48)
 * e de 10/09 (20:26) já não as continham, e a folha em PDF é a única cópia que
 * restou.
 *
 * Depois vieram mais três, da mesma família e bem mais antigas — a 0292
 * (03/06), a 0293 (05/05) e a 0296 (07/05) —, achadas pela conferência da
 * pasta contra a lista (servidor/conferir-pasta-os.js). Seis ao todo.
 *
 * DE ONDE VEM CADA CAMPO. Não há adivinhação silenciosa aqui:
 *   - o que a FOLHA mostra (número, data, coleção, desenho, grade, cores,
 *     medidas de cada fase, camadas, observação) vem lida da folha;
 *   - o que a folha NÃO mostra (componentes, aviamentos, etapas, ids dos
 *     cadastros) vem de uma OS-MOLDE: a OS existente com o MESMO desenho e a
 *     MESMA grade, onde esses campos são necessariamente os mesmos, porque o
 *     programa os deriva justamente do desenho e da grade.
 * Cada OS refeita declara o seu molde em `refeitaDe`, e todas levam
 * `refeitaEm`/`refeitaPor` — para que daqui a um ano ninguém precise adivinhar
 * de onde elas vieram.
 *
 * O QUE NASCE VAZIO, de propósito: o checklist (progresso), o status, a data de
 * finalização e o carimbo de etiqueta — QUANDO a folha nada diz sobre eles. A
 * folha em PDF costuma ser tirada no dia em que a OS nasce e não sabe o que
 * aconteceu depois; inventar progresso seria pior do que não ter.
 *
 * MAS QUANDO A FOLHA DIZ, ela manda. A da 0292 foi regravada em 06/08, com o
 * checklist inteiro marcado, as camadas de cada fase e a divisão por tonalidade
 * — e tudo isso entra, porque está escrito no papel.
 *
 * E HÁ UM TERCEIRO CASO, o das folhas antigas (0293 e 0296, de maio): elas
 * foram tiradas no dia do corte e mostram só o começo do checklist. Recriá-las
 * assim poria um lote de maio no painel de produção como se estivesse sendo
 * cortado hoje. Para essas, o status vai CARIMBADO como Estoque — que é onde um
 * lote de quatro meses atrás está — e a OS leva uma observação dizendo que o
 * status é do resgate e não do checklist. É inferência, e por isso ela é
 * escrita na própria OS em vez de ficar só aqui.
 *
 * Roda NO SERVIDOR da fábrica, pela mesma conta de serviço dos outros scripts
 * de reparo (read-modify-write do shared_data), e copia o blob inteiro para
 * backups/ antes de gravar.
 *
 * USO (confere primeiro, grava depois):
 *   node servidor/refazer-os-perdidas.js
 *   node servidor/refazer-os-perdidas.js --gravar
 */
const fs = require('fs');
const path = require('path');
const https = require('https');

const args = process.argv.slice(2);
const opt = (n) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : null; };
const GRAVAR = args.includes('--gravar');
const BASE = opt('--url') || 'https://193.168.0.200';
const ENV = opt('--env') || 'C:\\supabase\\docker\\.env';
const POR = (opt('--por') || 'admin@diverse.local').trim().toLowerCase();

/* AS TRÊS FOLHAS, lidas dos PDFs em
   J:\Meu Drive\Ordens de Serviço\Ordens de Serviço_Setor Corte.
   `molde` é o número da OS de onde saem os campos que a folha não mostra.
   `criadoEm` é a hora em que o PDF foi escrito no disco — o instante mais
   próximo do nascimento da OS que sobrou. */
const FOLHAS = [
  {
    os: '0509', data: '2026-09-01', criadoEm: '2026-09-01T12:11:00.000Z', molde: '0507',
    // A folha da 0509 é, campo a campo, a da 0507: mesmo desenho 0027, mesma
    // grade 2G | BM.LISA | 182cm, mesmas medidas (2,34×1,82 · 0,57×0,56 ·
    // 3,00×1,17), mesmas cores (Verde / Verde / viés Preto), 1 camada. Era uma
    // cópia dela — e a própria observação da folha diz isso.
    camadas: { 1: 1, 2: 1, 3: 1 },
    obsDoMolde: true
  },
  {
    os: '0536', data: '2026-09-10', criadoEm: '2026-09-10T20:15:00.000Z', molde: '0533',
    // Desenho 002 (Camiseta Básica), grade P ao G3 | CM.LISA | 117cm, tudo
    // Branco, medidas 8,20×1,17 · 1,80×0,57 · 8,00×1,17 — iguais às da 0533.
    // Corpo e gola SEM camadas (a folha mostra "–": ainda não enfestada);
    // o viés em 1, como sempre.
    camadas: { 1: '', 2: '', 3: 1 },
    nota: 'OS531 conjugada em enfesto único da OS536.',
    notaEm: '2026-09-10T19:56:00.000Z'
  },
  {
    os: '0537', data: '2026-09-10', criadoEm: '2026-09-10T20:17:00.000Z', molde: '0533',
    // Mesma grade e mesmas medidas da 0536 (o molde é o mesmo), mas com o
    // DESENHO TRICOLOR 0014: corpo Preto, gola Preto, viés Caqui, e a costura
    // é a CM.TRI. O desenho, as cores e as etapas vêm da 0538, que é a OS
    // tricolor com o mesmo desenho — os componentes, de lá também, porque a
    // grade tem o mesmo total (P ao G3, 7 peças).
    camadas: { 1: '', 2: '', 3: 1 },
    nota: 'OS531 conjugada em enfesto único da OS536.',
    notaEm: '2026-09-10T19:56:00.000Z',
    tricolorDe: '0538',
    coresDasFases: { 1: 'Preto Malha Algodão', 2: 'Preto Ribana Malha Algodão', 3: 'Caqui Malha Algodão' }
  },
  {
    /* A 0292 É GÊMEA DA 0410 — mesmo dia (03/06), mesmo desenho 001, mesma
       grade G-G3 | CM.LISA, as mesmas 59 camadas, os mesmos tons (26/16/17) e
       as mesmas fases marcadas (1 e 3, a gola não). O que a folha regravada em
       06/08 mostra bate campo a campo com o que está gravado na 0410; só as
       MEDIDAS das fases diferem, porque a folha saiu depois de a grade ser
       mexida. Valem as da folha. */
    os: '0292', data: '2026-06-03', criadoEm: '2026-06-03T12:00:00.000Z', molde: '0410',
    fases: { 1: { comp: '2.91', larg: '1.17' }, 2: { comp: '0.40', larg: '0.67' }, 3: { comp: '3', larg: '1.17' } },
    camadas: { 1: 59, 2: 11, 3: 1 },
    // O checklist inteiro marcado: o status sai dele, sem carimbo nenhum.
    progressoDoMolde: true
  },
  {
    /* A 0293 tem a estrutura da 0285 (mesma grade M-2G-GG, as mesmas duas fases
       e as mesmas medidas — 4,22×1,17 e 0,80×0,64) e as CORES da 0332, que é a
       OS do desenho 002, o branco. Os componentes vêm da 0285 porque a grade é
       a dela: as quantidades da folha (640/640/1.280/640/640, total 3.840) são
       exatamente as que estão lá. */
    os: '0293', data: '2026-05-05', criadoEm: '2026-05-05T12:00:00.000Z', molde: '0285',
    camadas: { 1: 80, 2: 16 },
    desenhoDe: '0332', corDe: '0332',
    carimbar: 'estoque',
    nota: 'OS refeita em 17/09/2026 pela folha em PDF da pasta, depois de ter sumido do programa. '
        + 'A folha é do dia do corte e mostra só o começo do checklist: o status foi carimbado como '
        + 'Estoque por ser um lote de maio, e não veio do checklist.'
  },
  {
    /* A 0296 é da 0332: mesmo desenho 002, mesma grade M-GG-G3, as mesmas duas
       fases com as mesmas medidas (3,85×1,17 e 0,50×0,67) e as mesmas 80
       camadas. Molde exato. */
    os: '0296', data: '2026-05-07', criadoEm: '2026-05-07T12:00:00.000Z', molde: '0332',
    camadas: { 1: 80, 2: 16 },
    carimbar: 'estoque',
    nota: 'OS refeita em 17/09/2026 pela folha em PDF da pasta, depois de ter sumido do programa. '
        + 'A folha é do dia em que a OS nasceu e não traz nada marcado: o status foi carimbado como '
        + 'Estoque por ser um lote de maio, e não veio do checklist.'
  }
];

const envTxt = fs.readFileSync(ENV, 'utf8');
const mKey = envTxt.match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m);
if (!mKey) { console.error('não achei SERVICE_ROLE_KEY em ' + ENV); process.exit(1); }
const KEY = mKey[1].trim();
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

const pk = v => typeof v === 'string' ? JSON.parse(v) : v;
const numeroOS = o => {
  const n = parseInt(String((o && o.os) || '').replace(/\D/g, ''), 10);
  return Number.isNaN(n) ? null : n;
};
const uid = (i) => 'id_' + Date.now() + '_' + (900 + i);

// Monta a OS refeita a partir do molde e do que a folha diz.
function refazer(folha, ordens, i) {
  const molde = ordens.find(o => numeroOS(o) === numeroOS(folha));
  if (molde) return { erro: `a OS ${folha.os} JÁ EXISTE — nada a refazer` };
  const base = ordens.find(o => String(o.os) === folha.molde);
  if (!base) return { erro: `não achei a OS-molde ${folha.molde}` };
  const nova = JSON.parse(JSON.stringify(base));

  nova.id = uid(i);
  nova.os = folha.os;
  nova.data = folha.data;
  nova.criadoEm = folha.criadoEm;
  nova.criadoPor = POR;

  /* A folha foi tirada quando a OS nasceu: nada do que veio depois é dela — a
     não ser quando ela MOSTRA o que veio depois (`progressoDoMolde`, o caso da
     0292, cuja folha foi regravada com o checklist inteiro marcado). */
  if (!folha.progressoDoMolde) delete nova.progresso;
  delete nova.statusOS; delete nova.statusOSPor; delete nova.statusOSEm;
  delete nova.finalizadaEm;
  delete nova.etiquetaEm; delete nova.etiquetaPor;
  delete nova.conjugadaStatusPaiId;

  // As medidas de cada fase, quando a folha mostra outras (a grade foi mexida
  // entre a OS e a folha).
  if (folha.fases) {
    (nova.fases || []).forEach(f => {
      const m = folha.fases[f.ordem];
      if (!m) return;
      f.comp = m.comp; f.larg = m.larg;
    });
    if (nova.enfesto && Array.isArray(nova.enfesto.blocos)) {
      nova.enfesto.blocos.forEach(b => {
        const m = folha.fases[b.ordem];
        if (!m) return;
        b.comp = Number(m.comp); b.larg = Number(m.larg);
      });
      const prim = folha.fases[1];
      if (prim) { nova.enfesto.comprimento = Number(prim.comp); nova.enfesto.largura = Number(prim.larg); }
    }
  }

  // As camadas que a folha mostra (vazio = fase ainda não enfestada).
  if (nova.enfesto && Array.isArray(nova.enfesto.blocos)) {
    nova.enfesto.blocos.forEach(b => {
      if (Object.prototype.hasOwnProperty.call(folha.camadas, b.ordem)) b.camadas = folha.camadas[b.ordem];
      b.bobinas = '';
    });
    const prim = nova.enfesto.blocos.find(b => b.ordem === 1);
    nova.enfesto.camadas = prim ? prim.camadas : '';
    nova.enfesto.totalPecas = (Number(nova.enfesto.camadas) || 0) * (nova.grade ? Number(nova.grade.total) || 0 : 0);
  }

  // O TRICOLOR DA 0537: o desenho, as cores e a costura vêm da OS tricolor.
  if (folha.tricolorDe) {
    const tri = ordens.find(o => String(o.os) === folha.tricolorDe);
    if (!tri) return { erro: `não achei a OS ${folha.tricolorDe}, de onde vem o desenho tricolor` };
    nova.desenhoId = tri.desenhoId;
    nova.codigo = tri.codigo;
    nova.modeloId = tri.modeloId; nova.modeloNome = tri.modeloNome;
    nova.variantes = JSON.parse(JSON.stringify(tri.variantes || []));
    nova.componentes = JSON.parse(JSON.stringify(tri.componentes || []));
    nova.etapas = JSON.parse(JSON.stringify(tri.etapas || []));
    // As cores de cada fase, como a folha as mostra.
    const porNome = {};
    [].concat(tri.tecidos || [], base.tecidos || []).forEach(t => { if (t.corNome) porNome[t.corNome] = t; });
    nova.tecidos = (nova.fases || []).map(f => {
      const alvo = folha.coresDasFases[f.ordem];
      const achado = porNome[alvo];
      return {
        tecidoId: f.tecidoId, tecidoNome: f.tecidoNome,
        corId: achado ? achado.corId : '', corNome: alvo,
        c1: '', lote: '', loteTexto: '—'
      };
    });
    if (nova.enfesto && Array.isArray(nova.enfesto.blocos)) {
      nova.enfesto.blocos.forEach(b => { if (folha.coresDasFases[b.ordem]) b.nomeCor = folha.coresDasFases[b.ordem]; });
    }
  }

  /* AS CORES VÊM DE OUTRA OS, casadas por TECIDO (o caso da 0293: a estrutura é
     da 0285, que é preta, e a folha é branca). Casar por tecido e não por
     posição é o que faz a ribana receber o branco DA RIBANA, e não o branco da
     malha — são dois cadastros de cor diferentes. */
  if (folha.corDe) {
    const fonte = ordens.find(o => String(o.os) === folha.corDe);
    if (!fonte) return { erro: `não achei a OS ${folha.corDe}, de onde vêm as cores` };
    const porTecido = new Map((fonte.tecidos || []).map(t => [String(t.tecidoId), t]));
    const porMaterial = new Map((fonte.componentes || []).map(c => [String(c.material || ''), c]));
    (nova.tecidos || []).forEach(t => {
      const f = porTecido.get(String(t.tecidoId));
      if (f) { t.corId = f.corId; t.corNome = f.corNome; }
    });
    (nova.componentes || []).forEach(c => {
      const f = porMaterial.get(String(c.material || ''))
        || porTecido.get(String(c.material || '').replace(/^T:/, ''));
      if (f) { c.cor = f.cor || f.corId || ''; c.corNome = f.corNome; }
    });
    nova.variantes = JSON.parse(JSON.stringify(fonte.variantes || nova.variantes || []));
    // O aviamento traz a cor no campo de APLICAÇÃO ("Preto"): é rótulo, e
    // acompanha a peça.
    const corPrincipal = (nova.tecidos && nova.tecidos[0] && nova.tecidos[0].corNome) || '';
    (nova.aviamentos || []).forEach(a => {
      const eraCor = (base.tecidos || []).some(t => t.corNome === a.app);
      if (eraCor && corPrincipal) a.app = corPrincipal;
    });
    if (nova.enfesto && Array.isArray(nova.enfesto.blocos)) {
      nova.enfesto.blocos.forEach(b => {
        const f = (nova.fases || []).find(x => x.ordem === b.ordem);
        const t = f && porTecido.get(String(f.tecidoId));
        if (t) b.nomeCor = t.corNome;
      });
    }
  }

  // O DESENHO vem de outra OS (a 0293 usa o 002, o branco, e a estrutura é de
  // uma OS do 001).
  if (folha.desenhoDe) {
    const fonte = ordens.find(o => String(o.os) === folha.desenhoDe);
    if (!fonte) return { erro: `não achei a OS ${folha.desenhoDe}, de onde vem o desenho` };
    nova.desenhoId = fonte.desenhoId;
    nova.codigo = fonte.codigo;
    nova.modeloId = fonte.modeloId; nova.modeloNome = fonte.modeloNome;
  }

  /* O CARIMBO DE STATUS, quando a folha não tem como dizer. Vai assinado pelo
     resgate e não por uma pessoa: quem olhar a dica na lista tem de ver que
     aquele status foi posto na recuperação. */
  if (folha.carimbar) {
    nova.statusOS = folha.carimbar;
    nova.statusOSPor = POR;
    nova.statusOSEm = new Date().toISOString();
  }

  // As observações: a da folha, ou as do molde quando a folha traz as mesmas.
  if (folha.nota) {
    nova.obs = '';
    nova.obsNotas = [{ login: POR, texto: folha.nota, em: folha.notaEm || folha.criadoEm }];
  } else if (!folha.obsDoMolde) {
    nova.obs = ''; delete nova.obsNotas;
  }

  // O recibo do resgate, dentro da própria OS.
  nova.refeitaDe = folha.molde + (folha.tricolorDe ? '+' + folha.tricolorDe : '');
  nova.refeitaEm = new Date().toISOString();
  nova.refeitaPor = POR;
  return { nova, base };
}

(async () => {
  const rows = JSON.parse((await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at')).body);
  if (!rows.length) { console.error('shared_data id=main não encontrado'); process.exit(1); }
  const data = rows[0].data || {};
  const eraString = typeof data.ordens === 'string';
  const ordens = pk(data.ordens) || [];
  console.log(`servidor: ${ordens.length} OS  ·  atualizado em ${rows[0].updated_at}`);

  const novas = [];
  FOLHAS.forEach((folha, i) => {
    const r = refazer(folha, ordens, i);
    if (r.erro) { console.log(`\nOS ${folha.os}: ${r.erro}`); return; }
    novas.push(r.nova);
    const cores = (r.nova.tecidos || []).map(t => t.corNome).join(' · ');
    const cam = (r.nova.enfesto && r.nova.enfesto.blocos || []).map(b => `${b.nomeTecido}:${b.camadas === '' ? '–' : b.camadas}`).join('  ');
    console.log(`\nOS ${folha.os}  (molde ${r.nova.refeitaDe})`);
    console.log(`  ${r.nova.modeloNome} · ${r.nova.colecaoNome} · ${r.nova.data} · grade ${r.nova.grade.descricao} (total ${r.nova.grade.total})`);
    console.log(`  cores: ${cores}`);
    console.log(`  camadas: ${cam}`);
    console.log(`  etapas: ${(r.nova.etapas || []).length}  ·  componentes: ${(r.nova.componentes || []).length}`
      + `  ·  aviamentos: ${(r.nova.aviamentos || []).length}`);
    console.log(`  observação: ${(r.nova.obsNotas || []).map(n => n.texto).join(' | ') || (r.nova.obs || '(nenhuma)')}`);
  });

  if (!novas.length) { console.log('\nnada a refazer.'); return; }
  if (!GRAVAR) {
    console.log(`\n(conferência — nada gravado. ${novas.length} OS prontas. Rode de novo com --gravar para aplicar.)`);
    return;
  }

  /* A FÁBRICA ESTÁ TRABALHANDO ENQUANTO ISTO RODA. Este script devolve o blob
     INTEIRO que leu lá em cima; se alguém gravou nesse meio-tempo, a gravação
     dele morreria aqui — que é exatamente o problema que este script veio
     reparar. Então confere o carimbo antes de escrever e desiste se mudou:
     rodar de novo custa segundos. */
  const agora = JSON.parse((await req('GET', '/rest/v1/shared_data?id=eq.main&select=updated_at')).body);
  if (!agora.length || agora[0].updated_at !== rows[0].updated_at) {
    console.error(`
ABORTADO: alguém gravou no servidor enquanto este script conferia`
      + ` (${rows[0].updated_at} -> ${agora[0] && agora[0].updated_at}).`
      + ` Nada foi escrito. Rode de novo.`);
    process.exit(1);
  }

  const quando = new Date().toISOString();
  const dir = path.join(__dirname, '..', 'backups');
  fs.mkdirSync(dir, { recursive: true });
  const arq = path.join(dir, 'shared_data-antes-refazer-os-' + quando.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(rows[0]), 'utf8');
  console.log('\ncópia de segurança: ' + arq);

  // Entram na ordem do número, junto das vizinhas — a lista da tela ordena por
  // número, mas o blob guardado na ordem certa é mais fácil de ler no futuro.
  const juntas = ordens.concat(novas).sort((a, b) => (numeroOS(a) || 0) - (numeroOS(b) || 0));
  data.ordens = eraString ? JSON.stringify(juntas) : juntas;
  data._device = 'refazer-os-perdidas-' + Date.now();
  await req('PATCH', '/rest/v1/shared_data?id=eq.main', { data, updated_at: quando });
  console.log(`\n✅ ${novas.length} OS refeitas (${novas.map(o => o.os).join(', ')}). `
    + `A lista passa de ${ordens.length} para ${juntas.length} OS. Os clientes conectados recarregam sozinhos.`);
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
