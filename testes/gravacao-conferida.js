/* Rode com:  node testes/gravacao-conferida.js

   GRAVAR SÓ SE NINGUÉM GRAVOU NO MEIO (30/10/2026: o número 700 pulado).

   Em 30/10 o Admin e a Nathaly gravaram a lista de OS no mesmo segundo. Cada
   um tinha relido o servidor e juntado as suas OS, mas o outro gravou entre a
   leitura e a gravação — e a lista de quem chegou por último apagou a OS do
   outro. Agora a gravação leva a versão lida (gravar_chaves_se) e o banco
   recusa se ela mudou; o cloudFlush relê, junta de novo e tenta outra vez.

   O que este teste protege, com o cloudFlush DE VERDADE (recortado do app.js),
   dois computadores e um banco de mentira com a mesma regra da função SQL:
     · duas gravações ao mesmo tempo, cada uma com uma OS nova: as DUAS ficam;
     · a segunda tentativa junta a partir das edições originais — a OS que o
       outro mudou no meio não volta à versão velha;
     · servidor sem a gravar_chaves_se: grava como antes (gravar_chaves);
     · outro gravando sem parar: desiste depois das tentativas e avisa. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  return src.slice(i, j + 2);
}
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^\\n]*;$', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

/* O BANCO: uma linha só (shared_data), com updated_at. Cada operação devolve
   o controle (await) antes de responder, para os dois computadores se
   intercalarem como na rede. */
const novoBanco = (opcoes) => {
  const o = opcoes || {};
  const banco = {
    linha: { updated_at: '2026-10-30T11:00:00.000Z', data: o.data || {} },
    gravacoes: [], recusadas: 0,
    async pausa() { for (let i = 0; i < 3; i++) await Promise.resolve(); },
    supa() {
      const b = banco;
      return {
        from(tabela) {
          return {
            select(cols) {
              return { eq() { return { async maybeSingle() {
                await b.pausa();
                if (tabela !== 'shared_data') return { data: null };
                if (b.antesDeLer) b.antesDeLer();
                return cols.includes('data')
                  ? { data: { updated_at: b.linha.updated_at, data: Object.assign({}, b.linha.data) } }
                  : { data: { updated_at: b.linha.updated_at } };
              } }; } };
            },
            async upsert() { await b.pausa(); return { error: null }; }
          };
        },
        async rpc(nome, args) {
          await b.pausa();
          if (b.antesDeGravar) b.antesDeGravar();
          if (nome === 'gravar_chaves_se') {
            if (o.semConferida) return { data: null, error: { code: 'PGRST202', message: 'Could not find the function' } };
            if (b.linha.updated_at && args.esperado !== b.linha.updated_at) { b.recusadas++; return { data: false, error: null }; }
          } else if (nome !== 'gravar_chaves') {
            return { data: null, error: { code: 'PGRST202', message: 'Could not find the function' } };
          }
          Object.assign(b.linha.data, args.pares);
          b.linha.updated_at = args.carimbo;
          b.gravacoes.push({ nome, quem: args.pares._device });
          return { data: true, error: null };
        }
      };
    }
  };
  return banco;
};

// Um computador: o cloudFlush do app.js com o estado dele.
let relogio = Date.parse('2026-10-30T11:56:31.000Z');
const computador = (banco, device) => new Function('ctx', `
  const supa = ctx.supa;
  const currentUser = { id: ctx.device };
  const DEVICE_ID = ctx.device;
  let cloudCache = {}, _baseline = {}, _apagadosAqui = {};
  const _dirtyKeys = new Set();
  let _flushing = false, _flushDesde = 0, _flushGeracao = 0, _retryTimer = null;
  let _cloudLoadErro = false, _permitirFlushVazio = true, _appJaTeveExpedicao = false, _appJaTeveDados = false;
  let _ultimoUpdatedAtServidor = null, _ultimoUpdatedAtEnviado = null, _versoesConhecidas = null;
  let _semGravacaoParcial = false;
  let printOsAtual = null;
  const Date_ = Date;
  const podeGravar = () => true;
  const setSyncStatus = (s) => { ctx.status = s; };
  const _flushPodeComecar = (f) => !f;
  const scheduleCloudSave = () => {};
  const mostrarAlertaSalvamento = (tipo, msg) => { ctx.alertas.push(tipo + ': ' + msg); };
  const _blobEstaVazio = () => false;
  const _contarItens = () => 1;
  const _lerVersoesServidor = async () => null;
  // O STATE do computador: a tela e todo saveState leem dele.
  const STATE = { ordens: [] };
  const loadState = async () => { await Promise.resolve(); STATE.ordens = JSON.parse(cloudCache.ordens || '[]'); };
  const saveState = (k) => { cloudCache[k] = JSON.stringify(STATE[k]); _dirtyKeys.add(k); };
  const salvarSnapshotContingencia = () => {};
  const snapshotDiario = async () => {};
  const toast = () => {};
  const setTimeout = () => 0, clearTimeout = () => {};
  const console = { warn: () => {}, error: () => {}, log: () => {} };
  ${constante('GRAVACAO_TENTATIVAS')}
  let _semGravacaoConferida = false;
  ${recorte('function _versaoConfere', 'a conferencia da versao')}
  ${recorte('function _ehFuncaoAusente', 'a funcao ausente')}
  ${recorte('function _paresParaGravar', 'o montador do pacote')}
  ${recorte('function _juntarCampos', 'o juntar campo a campo')}
  ${recorte('function _mergeObjetoPorCampo', 'o merge de objeto')}
  ${recorte('function _mergeListaPorRegistro', 'o merge por registro')}
  ${recorte('async function cloudFlush', 'o cloudFlush')}
  return {
    cloudFlush,
    // Começa sincronizado com o banco, como depois de uma leitura.
    sincronizar(linha) {
      cloudCache = Object.assign({}, linha.data); _baseline = Object.assign({}, linha.data);
      _ultimoUpdatedAtServidor = linha.updated_at;
      STATE.ordens = JSON.parse(cloudCache.ordens || '[]');
    },
    // Um clique na tela: muda o STATE e grava a chave, como o app faz.
    clicar(chave, mudar) { mudar(STATE); saveState(chave); },
    editar(chave, valor) { cloudCache[chave] = valor; _dirtyKeys.add(chave); },
    cache: () => cloudCache, sujas: () => _dirtyKeys
  };
`)(Object.assign({ supa: banco.supa(), device, alertas: [] }));

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const lista = (...v) => JSON.stringify(v);
const numeros = (b) => JSON.parse(b.linha.data.ordens).map(o => o.os).sort().join(',');

(async () => {
  console.log('-- o caso de 30/10: duas OS novas gravadas no mesmo segundo --');
  {
    const banco = novoBanco({ data: { ordens: lista({ id: 'a', os: '0699' }) } });
    const nathaly = computador(banco, 'nathaly'), admin = computador(banco, 'admin');
    nathaly.sincronizar(banco.linha); admin.sincronizar(banco.linha);
    nathaly.editar('ordens', lista({ id: 'a', os: '0699' }, { id: 'n', os: '0700' }));
    admin.editar('ordens', lista({ id: 'a', os: '0699' }, { id: 'm', os: '0701' }));
    await Promise.all([nathaly.cloudFlush(), admin.cloudFlush()]);
    ok('1. as DUAS OS ficam no banco', numeros(banco) === '0699,0700,0701', numeros(banco));
    ok('2. uma das gravacoes foi recusada e refeita', banco.recusadas === 1 && banco.gravacoes.length === 2, [banco.recusadas, banco.gravacoes]);
    ok('3. foi pela gravar_chaves_se', banco.gravacoes.every(g => g.nome === 'gravar_chaves_se'), banco.gravacoes);
    ok('4. as chaves sujas baixam nos dois', !nathaly.sujas().size && !admin.sujas().size);
  }

  console.log('\n-- a segunda tentativa parte das edicoes originais --');
  {
    // A Nathaly muda o status da 0699. O admin troca o tom da mesma OS duas
    // vezes: o tom 2 antes de ela ler (entra na junção), o tom 3 entre a junção
    // e a gravação dela (a gravação é recusada). Na segunda tentativa vale o
    // tom 3: se a junção partisse da anterior, o tom 2 dela pareceria "editado
    // aqui" e venceria o 3.
    const banco = novoBanco({ data: { ordens: lista({ id: 'a', os: '0699', st: 'corte', tom: 1 }) } });
    const nathaly = computador(banco, 'nathaly');
    nathaly.sincronizar(banco.linha);
    nathaly.editar('ordens', lista({ id: 'a', os: '0699', st: 'costura', tom: 1 }));
    let leu = false, gravou = false;
    banco.antesDeLer = () => {
      if (leu) return; leu = true;
      banco.linha.data.ordens = lista({ id: 'a', os: '0699', st: 'corte', tom: 2 }); banco.linha.updated_at = '2026-10-30T11:56:32.000Z';
    };
    banco.antesDeGravar = () => {
      if (gravou) return; gravou = true;
      banco.linha.data.ordens = lista({ id: 'a', os: '0699', st: 'corte', tom: 3 }); banco.linha.updated_at = '2026-10-30T11:56:33.000Z';
    };
    await nathaly.cloudFlush();
    ok('5b. a primeira gravacao foi recusada', banco.recusadas === 1, banco.recusadas);
    const o = JSON.parse(banco.linha.data.ordens)[0];
    ok('5. o status da Nathaly subiu', o.st === 'costura', o);
    ok('6. e o tom MAIS NOVO do admin ficou (3, nao o 2 da primeira juncao)', o.tom === 3, o);
  }

  console.log('\n-- servidor sem a gravar_chaves_se --');
  {
    const banco = novoBanco({ semConferida: true, data: { ordens: lista({ id: 'a', os: '0699' }) } });
    const t = computador(banco, 'admin');
    t.sincronizar(banco.linha);
    t.editar('ordens', lista({ id: 'a', os: '0699' }, { id: 'm', os: '0700' }));
    await t.cloudFlush();
    ok('7. grava pela gravar_chaves, como antes', banco.gravacoes.length === 1 && banco.gravacoes[0].nome === 'gravar_chaves'
       && numeros(banco) === '0699,0700', [banco.gravacoes, numeros(banco)]);
  }

  console.log('\n-- outro computador gravando sem parar --');
  {
    // Entre cada leitura e cada gravação, outro computador grava.
    const banco = novoBanco({ data: { ordens: lista({ id: 'a', os: '0699' }) } });
    let n = 0;
    banco.antesDeGravar = () => { banco.linha.updated_at = '2026-10-30T12:00:' + String(++n).padStart(2, '0') + '.000Z'; };
    const furado = computador(banco, 'admin');
    furado.sincronizar(banco.linha);
    furado.editar('ordens', lista({ id: 'a', os: '0699' }, { id: 'm', os: '0700' }));
    await furado.cloudFlush();
    ok('8. desiste depois das tentativas, sem gravar por cima', banco.recusadas === 4 && !numeros(banco).includes('0700'), [banco.recusadas, numeros(banco)]);
    ok('9. e a chave continua suja, para a proxima tentativa', furado.sujas().has('ordens'));
  }

  console.log('\n-- o clique durante a gravacao nao desfaz a mudanca do outro (30/10: renumeracao desfeita) --');
  {
    // O servidor renumera a OS (0701 -> 0700) por fora. Este computador grava
    // outra coisa (meta); no meio da gravacao, alguem marca uma etapa na OS.
    // Depois, a gravacao da etapa nao pode levar o 0701 de volta.
    const banco = novoBanco({ data: { ordens: lista({ id: 'a', os: '0701', st: 'corte' }), meta: '{}' } });
    const t = computador(banco, 'admin');
    t.sincronizar(banco.linha);
    banco.linha.data = Object.assign({}, banco.linha.data, { ordens: lista({ id: 'a', os: '0700', st: 'corte' }) });
    banco.linha.updated_at = '2026-10-30T12:53:00.000Z';
    t.editar('meta', '{"fila":[1]}');
    let clicou = false;
    banco.antesDeGravar = () => {
      if (clicou) return; clicou = true;
      t.clicar('ordens', st => { st.ordens[0].st = 'costura'; });
    };
    await t.cloudFlush();   // grava o meta; o clique acontece no meio
    await t.cloudFlush();   // grava a etapa
    const o = JSON.parse(banco.linha.data.ordens)[0];
    ok('10. a etapa marcada subiu', o.st === 'costura', o);
    ok('11. e o numero novo (0700) NAO voltou para o antigo', o.os === '0700', o);
  }

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
