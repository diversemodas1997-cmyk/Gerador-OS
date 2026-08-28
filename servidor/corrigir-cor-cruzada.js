/*
 * Passa a cor gravada para a cor DO TECIDO DAQUELA LINHA.
 *
 * POR QUE ISTO EXISTE
 *
 * Junior, 28/08/2026: "o campo estoque por tecido mais cor mostra moletom:
 * bege malha algodao. Essa informacao e incongruente."
 *
 * Estava certo. A peca tem UMA cor e o cadastro de cores e POR TECIDO (o tecido
 * vive so no sufixo do nome: "Bege Malha Algodao", "Bege Moletom", "Bege Ribana
 * Malha Algodao"). A cor do componente do desenho descia igual para a linha de
 * qualquer tecido, e o estoque ganhava a prateleira fantasma
 * "Moletom · Bege Malha Algodao" ao lado da de verdade — 21 prateleiras fisicas
 * partidas em duas linhas.
 *
 * O CONSERTO DE VERDADE E NO CODIGO, NAO AQUI
 *
 * `corCanonicaPorTecido` passou a corrigir tambem a cor composta de outro
 * tecido, e `movimentacoesEstoque` ja fazia essa conversao NA LEITURA. Ou seja:
 * a tela e o saldo ja saem certos sem este script. Ele existe para o dado
 * GRAVADO tambem ficar certo — o que le o razao cru (a copia no ERP, o JSON do
 * backup, a contabilidade) nao passa por aquela funcao.
 *
 * O caso que so este script resolve: as 26 entradas manuais da ABERTURA DE
 * ESTOQUE de 27/08. Movimento de OS e reescrito toda vez que a OS e salva, e ja
 * nascera certo; entrada manual nao se regenera nunca.
 *
 * A REGRA E RECORTADA DO app.js
 *
 * Nao ha uma segunda versao dela aqui. Se o app mudar de ideia sobre o que e a
 * cor certa, este script muda junto — e uma copia consertada sem a outra e como
 * o erro sobrevive.
 *
 * O QUE ELE NAO FAZ
 *
 * Nao inventa cor: quando nao existe no cadastro a cor base sob aquele tecido,
 * deixa como esta e relata. E o caso dos Texturizados, cuja convencao de nome
 * usa o nome curto ("Preto Rugao" para "Texturizado | Rugao") — mexer neles e
 * decisao de cadastro, nao de script.
 *
 * Nao toca no cadastro dos DESENHOS. Os 76 componentes com cor de outro tecido
 * continuam la; o formulario passou a re-resolver a cor ao montar a linha, e e
 * de proposito que o cadastro em si so mude com o Junior olhando.
 *
 * COMO RODAR
 *
 *   node servidor/corrigir-cor-cruzada.js             so relata, nao grava
 *   node servidor/corrigir-cor-cruzada.js --gravar    grava no servidor
 *
 * Antes de gravar ele salva o blob inteiro em backups/, e a escrita e
 * read-modify-write com CONFERENCIA DE CARIMBO: se alguem gravou no servidor
 * entre a leitura e a escrita, aborta em vez de passar por cima (o blob e uma
 * linha so, compartilhada — ver project_restauracao_credenciada).
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

/* ---- a regra vem do app.js, recortada: uma so versao dela ---- */
function regraDoApp(STATE) {
  const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
  const solta = nome => {
    const i = src.indexOf('function ' + nome + '(');
    if (i < 0) throw new Error('nao achei ' + nome + ' no app.js');
    return src.slice(i, src.indexOf('\n}', i) + 2);
  };
  return new Function('STATE', [
    solta('_normNome'), solta('_sufixoTecidoNorm'),
    solta('corBaseNome'), solta('corCanonicaPorTecido'),
    'return { corCanonicaPorTecido, _normNome };'
  ].join('\n'))(STATE);
}

/* ---- o servidor da fabrica ---- */
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
  // Cada chave do blob e guardada como TEXTO JSON, nao como objeto — le e
  // devolve pela mesma porta, senao a chave vira um objeto e o app nao le mais.
  const le = k => JSON.parse(data[k] || '[]');
  const estoqueMov = le('estoqueMov');
  const ordens = le('ordens');
  const STATE = { cores: le('cores'), tecidos: le('tecidos') };
  const { corCanonicaPorTecido, _normNome } = regraDoApp(STATE);
  const corPorNome = nome => (STATE.cores || []).find(c => _normNome(c.nome) === _normNome(nome));

  // Um so lugar decide: devolve a cor certa, ou '' quando nao ha o que mudar.
  const certa = (corNome, tecidoNome) => {
    if (!corNome || !tecidoNome) return '';
    const r = corCanonicaPorTecido(corNome, tecidoNome);
    return _normNome(r) === _normNome(corNome) ? '' : r;
  };

  const contas = { estoqueMov: 0, linhasTecido: 0, blocos: 0 };
  const porPar = new Map();
  const semDestino = new Map();

  // 1. o razao do estoque — inclusive as entradas manuais da abertura, que sao
  //    as unicas que nunca se regeneram sozinhas.
  estoqueMov.forEach(m => {
    const nova = certa(m.corNome, m.tecidoNome);
    if (!nova) {
      if (m.corNome && m.tecidoNome) {
        const n = _normNome(m.corNome), t = _normNome(m.tecidoNome);
        // Sufixo que nao e tecido cadastrado: nada casa, e fica como esta.
        if (!n.endsWith(' ' + t) && n !== t) {
          const k = m.tecidoNome + ' :: ' + m.corNome;
          semDestino.set(k, (semDestino.get(k) || 0) + 1);
        }
      }
      return;
    }
    const k = m.tecidoNome + ' :: ' + m.corNome + '  ->  ' + nova;
    porPar.set(k, (porPar.get(k) || 0) + 1);
    m.corNome = nova;
    contas.estoqueMov++;
  });

  // 2. a linha de Tecidos da OS — a fonte de onde o movimento e recalculado.
  //    Leva o corId junto: sem ele, reabrir a OS traria a cor velha de volta.
  ordens.forEach(o => {
    (o.tecidos || []).forEach(t => {
      const nova = certa(t.corNome, t.tecidoNome);
      if (!nova) return;
      t.corNome = nova;
      const c = corPorNome(nova);
      if (c) t.corId = c.id;
      contas.linhasTecido++;
    });
    // 3. o bloco do enfesto guarda um retrato da cor; o nomeTecido dele e o
    //    nome da FASE ("Gola"), nao o do tecido — por isso o tecido vem da
    //    linha de Tecidos de mesmo indice, que e como o app ja o re-deriva.
    const blocos = (o.enfesto && o.enfesto.blocos) || [];
    if (blocos.length && (o.tecidos || []).length === blocos.length) {
      blocos.forEach((b, i) => {
        const nova = certa(b.nomeCor, (o.tecidos[i] || {}).tecidoNome);
        if (!nova) return;
        b.nomeCor = nova;
        contas.blocos++;
      });
    }
  });

  console.log('');
  console.log('COR CRUZADA -> COR DO TECIDO DA LINHA');
  console.log('');
  [...porPar.entries()].sort().forEach(([k, n]) =>
    console.log('  ' + String(n).padStart(4) + 'x  ' + k));
  console.log('');
  console.log('  movimentos de estoque corrigidos : ' + contas.estoqueMov);
  console.log('  linhas de Tecidos de OS          : ' + contas.linhasTecido);
  console.log('  blocos de enfesto                : ' + contas.blocos);

  if (semDestino.size) {
    console.log('');
    console.log('DEIXADOS COMO ESTAO (nao existe a cor base sob esse tecido no cadastro):');
    [...semDestino.entries()].sort().forEach(([k, n]) =>
      console.log('  ' + String(n).padStart(4) + 'x  ' + k));
  }

  const total = contas.estoqueMov + contas.linhasTecido + contas.blocos;
  if (!total) { console.log('\nNada a corrigir.'); return; }

  if (!GRAVAR) {
    console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.');
    return;
  }

  data.estoqueMov = JSON.stringify(estoqueMov);
  data.ordens = JSON.stringify(ordens);

  const arq = path.join(RAIZ, 'backups',
    'shared_data-antes-cor-cruzada-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(await (await lerBlob()).data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  await gravarBlob(cab, data, updatedAt);
  console.log('gravado no servidor: ' + total + ' correcoes.');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
