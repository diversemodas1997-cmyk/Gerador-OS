/*
 * Troca o TECIDO (e opcionalmente a cor) de uma OS ja emitida, e refaz a baixa
 * de estoque dela pela conta do proprio programa.
 *
 * POR QUE ISTO EXISTE
 *
 * A OS 0483 e a 0484 sairam da grade `P-M-G-GG-G1 | COT.JAC | 157cm`, que
 * estava cadastrada com `Texturizado | Rugao` em vez de `Texturizado | Jaguar`
 * — copia da grade COT.RUG, cujo encaixe da exatamente a mesma medida (os dois
 * riscos dao 5,1006 m x 1,57 m, porque as duas bobinas sao de 157 cm). So o
 * campo do tecido estava errado; a medida estava certa.
 *
 * A OS copia as fases da grade no momento em que e emitida, entao corrigir a
 * grade depois NAO alcanca a OS. Dai este script.
 *
 * A PROVA DE QUE ERA JAGUAR
 *
 * Os desenhos dizem: 0032 "Camiseta Oversized Texturizada | Jacguar Preto" e
 * 0033 "... | Jacguar Off-white". Os componentes da 0483 ja traziam
 * `Texturizado | Jaguar`; so o enfesto e que carregava Rugao. As seis OS da
 * familia formam 3 panos x 2 cores, e so a dupla JAC saiu torta.
 *
 * Na 0484 a cor tambem estava errada — "Off-white Prime" — porque a cor
 * principal do desenho 0033 aponta a Prime, e os componentes herdaram dela o
 * pano Prime. A cor certa, "Off-white Jacguar", ja existe no cadastro. O
 * DESENHO 0033 nao e corrigido aqui: e cadastro, e mexer nele muda o molde de
 * OS futuras — fica para o Junior, no proprio programa.
 *
 * O KG NAO E DIGITADO
 *
 * Jaguar tem 250 g/m2 e Rugao 270, entao o peso muda junto com o pano. O novo
 * valor sai de `consumoAgregadoPorTecidoCor` recortada do app.js — a mesma
 * funcao que a tela usa. Digitar o numero aqui criaria uma segunda conta.
 *
 * O movimento e reescrito NO LUGAR: mesmo id, mesma data, mesmo status. As duas
 * OS estao finalizadas, e inventar um movimento novo faria a baixa aparecer
 * duas vezes no historico.
 *
 * COMO RODAR
 *
 *   node servidor/corrigir-tecido-da-os.js             so relata, nao grava
 *   node servidor/corrigir-tecido-da-os.js --gravar    grava no servidor
 *
 * Antes de gravar salva o blob em backups/, e a escrita confere o carimbo.
 *
 * ATENCAO: recarregue a pagina do programa (F5) depois de rodar. Aba aberta
 * desde antes regrava a chave inteira com o estado velho e desfaz isto —
 * aconteceu tres vezes em 28/08 (ver project_restauracao_credenciada).
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

// O que corrigir. Por NOME, nao por id: o nome e o que da para conferir lendo.
const ALVOS = [
  { os: '0483', deTecido: 'Texturizado | Rugão', paraTecido: 'Texturizado | Jaguar' },
  { os: '0484', deTecido: 'Texturizado | Rugão', paraTecido: 'Texturizado | Jaguar',
    deCor: 'Off-white Prime', paraCor: 'Off-white Jacguar' }
];

/* ---- a conta do consumo vem do app.js, recortada ---- */
function contaDoApp(STATE) {
  const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
  const fn = n => {
    const i = src.indexOf('function ' + n + '(');
    if (i < 0) throw new Error('nao achei ' + n + ' no app.js');
    return src.slice(i, src.indexOf('\n}', i) + 2);
  };
  const cst = n => {
    const m = src.match(new RegExp('^const ' + n + ' = [^;]+;', 'm'));
    if (!m) throw new Error('nao achei a constante ' + n);
    return m[0];
  };
  const CONSTS = ['LIMITE_CAMADAS', 'MULTIPLICADOR_PECAS', 'LABEL_CATEGORIA',
                  'UNIDADES_PADRAO_FORRO', 'CEIL_BOBINA_EPS', '_EXC_LIGACAO', '_PAL_VIES'];
  const FNS = ['_normNome', '_sufixoTecidoNorm', 'corBaseNome', 'corCanonicaPorTecido',
    'categoriaEfetivaTecido', 'isTecidoRibana', 'calcularPapeisFases',
    '_tamanhoQueMandaNaGrade', 'camadasDaFaseRibana', '_ribanaEscalaComGrade',
    'camadasDaFaseForro', 'camadasPadraoDaFase', 'camadasCheiasDaFase',
    'multiplicadorPecaOS', '_faseNaoEnfestadaPorTom', 'unidadesPorCamadaTecido',
    'unidadesPorCamadaPrincipal', 'tecidosDaOS', 'gramaturaTecidoPorNome',
    'pesoBobinaPorNome', '_normFaseNome', '_faseSoDe', 'compraLimiteCamadasGrade',
    'consumoEnfestoOS', 'parseBobinas', 'bobinaInteira', 'ehFaseRibana',
    'bobinasEfetivasFase', 'consumoAgregadoPorTecidoCor'];
  return new Function('STATE', 'var comprasCache = [];\n'
    + CONSTS.map(cst).join('\n') + '\n' + FNS.map(fn).join('\n')
    + '\nreturn { consumoAgregadoPorTecidoCor };')(STATE);
}

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
  const tecidos = le('tecidos'), cores = le('cores');
  const STATE = { ordens, estoqueMov, tecidos, cores, grades: le('grades'), desenhos: le('desenhos'), componentes: [] };
  const { consumoAgregadoPorTecidoCor } = contaDoApp(STATE);
  const acha = (lista, nome) => lista.find(x => x.nome === nome);

  let mudou = 0;
  ALVOS.forEach(alvo => {
    const o = ordens.find(x => String(x.os) === alvo.os);
    if (!o) { console.log('  OS ' + alvo.os + ' nao encontrada'); return; }
    const tNovo = acha(tecidos, alvo.paraTecido);
    if (!tNovo) { console.log('  tecido "' + alvo.paraTecido + '" nao existe no cadastro'); return; }
    const cNovo = alvo.paraCor ? acha(cores, alvo.paraCor) : null;
    if (alvo.paraCor && !cNovo) { console.log('  cor "' + alvo.paraCor + '" nao existe no cadastro'); return; }

    const antes = estoqueMov.filter(m => m.origem === 'os' && m.osId === o.id)
      .map(m => m.tecidoNome + ' :: ' + m.corNome + '  ' + m.kg + ' kg');

    // 1. as FASES — a copia do enfesto que a OS carrega, e de onde sai o kg
    (o.fases || []).forEach(f => {
      if (f.tecidoNome === alvo.deTecido) { f.tecidoId = tNovo.id; f.tecidoNome = tNovo.nome; }
    });
    // 2. as LINHAS DE TECIDOS — de onde sai a cor de cada fase
    (o.tecidos || []).forEach(t => {
      if (t.tecidoNome === alvo.deTecido) { t.tecidoId = tNovo.id; t.tecidoNome = tNovo.nome; }
      if (cNovo && t.corNome === alvo.deCor) { t.corId = cNovo.id; t.corNome = cNovo.nome; }
    });
    // 3. os BLOCOS do enfesto guardam um retrato da cor
    if (cNovo) ((o.enfesto || {}).blocos || []).forEach(b => {
      if (b.nomeCor === alvo.deCor) b.nomeCor = cNovo.nome;
    });
    // 4. os COMPONENTES — sao eles que alimentam o Estoque de corte (pecas),
    //    que e outro campo e nao seria alcancado pelas fases.
    (o.componentes || []).forEach(c => {
      if (alvo.deCor && c.materialNome && c.materialNome !== alvo.paraTecido
          && c.corNome === alvo.deCor) {
        c.material = 'T:' + tNovo.id; c.materialNome = tNovo.nome;
      }
      if (cNovo && c.corNome === alvo.deCor) { c.cor = cNovo.id; c.corNome = cNovo.nome; }
    });

    // 5. a BAIXA DE ESTOQUE, refeita pela conta do programa — no lugar, com o
    //    mesmo id e o mesmo status: a OS ja esta finalizada.
    const itens = consumoAgregadoPorTecidoCor(o);
    const meus = estoqueMov.filter(m => m.origem === 'os' && m.osId === o.id);
    if (meus.length !== itens.length) {
      console.log('  OS ' + alvo.os + ': o consumo virou ' + itens.length
        + ' linha(s) e havia ' + meus.length + ' — nao reescrevo as cegas.');
      return;
    }
    meus.forEach((m, i) => {
      m.tecidoNome = itens[i].tecidoNome;
      m.corNome = itens[i].corNome;
      m.kg = Math.round(itens[i].kg * 1000) / 1000;
    });

    console.log('  OS ' + alvo.os + ':');
    console.log('     tecido : ' + alvo.deTecido + '  ->  ' + tNovo.nome);
    if (cNovo) console.log('     cor    : ' + alvo.deCor + '  ->  ' + cNovo.nome);
    antes.forEach((a, i) => console.log('     estoque: ' + a + '   ->   '
      + meus[i].tecidoNome + ' :: ' + meus[i].corNome + '  ' + meus[i].kg + ' kg'));
    mudou++;
  });

  console.log('');
  if (!mudou) { console.log('Nada corrigido.'); return; }
  if (!GRAVAR) { console.log('SIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-tecido-da-os-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify((await lerBlob()).data), 'utf8');
  console.log('copia de seguranca: ' + path.relative(RAIZ, arq));

  data.ordens = JSON.stringify(ordens);
  data.estoqueMov = JSON.stringify(estoqueMov);
  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + mudou + ' OS.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
