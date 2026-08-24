/* Rode com:  node testes/pasta-oe-sumida.js

   PASTA DAS OE QUE NÃO EXISTE MAIS.

   A pasta é escolhida uma vez e o handle fica guardado no IndexedDB. Só que a
   pasta do disco pode ser renomeada, movida, apagada — ou estar num drive que
   não está montado hoje (o Google Drive em J:\ é o caso comum aqui). O handle
   sobrevive a tudo isso: o navegador responde 'granted' na permissão e só
   estoura na PRIMEIRA operação de disco, com um "NotFoundError" cru — depois
   de o app ter gerado o PDF inteiro.

   Era isso que o usuário via ao clicar em "Salvar OE na pasta": esperava a
   captura e recebia um aviso que não dizia nem o que aconteceu nem o que fazer.

   O que este teste guarda:
     • pasta sumida é descoberta ANTES da captura (nenhum PDF é gerado);
     • a mensagem diz o nome da pasta e onde reconectar;
     • o sumiço no meio do caminho (entre o probe e a gravação) cai na mesma
       mensagem, e não no erro do navegador;
     • pasta boa continua gravando;
     • o clique sobre uma gravação em curso FALA (antes o botão emudecia);
     • o gatilho do Drive: pede que ele seja aberto, vigia a pasta e retoma a
       gravação sozinho quando ela reaparece;
     • o porteiro que TODAS as pastas usam (PDF da OS, etiquetas, OE, backup,
       snapshots, exportação): pasta muda, o aviso e a retomada não.

   Recorta pastaAcessivel, _ehErroPastaSumiu e salvarPdfOeNaPasta do app.js de
   verdade; o resto entra dublado — o que se mede aqui é a decisão, não como o
   PDF é desenhado. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
// Delimitador '\n}' (e nao '\n}\n'): o arquivo e gravado com CRLF.
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';

const motor = [
  corta('async function pastaAcessivel'),
  corta('function _ehErroPastaSumiu'),
  corta('function _modalDriveTexto'),
  corta('function pedirGoogleDrive'),
  corta('function _pararVigiaDrive'),
  corta('async function _driveVoltou'),
  corta('async function tentarDriveAgora'),
  corta('function fecharPedidoDrive'),
  corta('function _avisoPastaSumiu'),
  corta('async function pastaProntaOuPedeDrive'),
  corta('function tratarErroPastaSumiu'),
  corta('async function salvarPdfOeNaPasta')
].join('\n');

// Pasta dublê. `sumida` = o disco não tem mais essa pasta (values() estoura).
// `sumidaAoGravar` = existe na listagem, mas some na hora de criar o arquivo.
function pasta({ nome = 'OE', sumida = false, sumidaAoGravar = false } = {}) {
  const err = () => { const e = new Error('A requested file or directory could not be found'); e.name = 'NotFoundError'; return e; };
  const escritos = [];
  return {
    name: nome,
    escritos,
    // O Drive subindo = a pasta voltando a existir no meio da sessão.
    driveAbriu() { sumida = false; sumidaAoGravar = false; },
    async queryPermission() { return 'granted'; },
    async requestPermission() { return 'granted'; },
    async *values() { if (sumida) throw err(); },
    async getFileHandle(fn) {
      if (sumidaAoGravar) throw err();
      return { async createWritable() { return { async write(b) { escritos.push({ fn, b }); }, async close() {} }; } };
    }
  };
}

// Roda um clique em "Salvar OE na pasta" (silent=false) sobre a pasta dada.
// Devolve { ok, toasts, capturas } — capturas conta quantas vezes o PDF foi
// gerado, que é o custo que o probe existe para evitar.
async function clicarSalvar(handle, { travada = false, roteiro = null } = {}) {
  const toasts = [];
  const conta = { capturas: 0 };
  const fn = new Function('handle', 'toasts', 'conta', 'travada', 'roteiro', `
    return (async () => {
      let _oeSalvando = travada, _oeSalvandoDesde = travada ? Date.now() : 0;
      let _oeAvisoDado = '', oeFolderHandle = handle;
      let expPlanoModo = 'dia', currentRole = 'admin';
      const modal = { aberto: false, texto: '' };
      const document = {
        querySelector: () => null,
        getElementById: () => ({ set innerHTML(v) { modal.texto = v; }, get innerHTML() { return modal.texto; } })
      };
      const openModal = () => { modal.aberto = true; };
      const closeModal = () => { modal.aberto = false; };
      const esc = (v) => String(v == null ? '' : v);
      let _pedidoDrive = null, _pedidoDriveAutoDado = false;
      const toast = (m, t) => toasts.push({ m, t: t || '' });
      const _avisarOeUmaVez = (chave, m) => toasts.push({ m, t: 'err', silencioso: true });
      const ensureFolderPermission = async () => true;
      const loadOeFolderHandle = async () => handle;
      const oeTemConteudo = () => true;                       // há OE no período
      const renderPrintPlanoExpedicao = () => {};
      const _comFolhaOeRenderizavel = async (f) => await f();
      const gerarPdfDaSheetExp = async () => { conta.capturas++; return 'BLOB'; };
      const oeFilenameForPlano = () => 'OE-24-08-2026.pdf';
      ${motor}
      const r = await salvarPdfOeNaPasta();
      // A vigia é um setInterval de 4s: sem desligar, o teste ficaria pendurado.
      const pedido = { modalAberto: modal.aberto, texto: modal.texto };
      if (roteiro) return await roteiro({ salvar: salvarPdfOeNaPasta, tentarDriveAgora, pedido, modal });
      _pararVigiaDrive();
      return { r, pedido };
    })();
  `);
  const { r, pedido } = await fn(handle, toasts, conta, travada, roteiro);
  return { ok: r, toasts, capturas: conta.capturas, pedido };
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

(async () => {
  /* ---------- 1. pasta sumida: descoberta antes da captura ---------- */
  let p = pasta({ nome: 'OE 2026', sumida: true });
  let r = await clicarSalvar(p);
  const msg1 = (r.toasts[0] || {}).m || '';
  ok('pasta sumida: não diz que salvou', r.ok === false, r.ok);
  ok('pasta sumida: nenhum PDF é gerado à toa', r.capturas === 0, r.capturas);
  ok('pasta sumida: o aviso nomeia a pasta', /OE 2026/.test(msg1), msg1);
  ok('pasta sumida: o aviso diz "não encontrada"', /não encontrada/i.test(msg1), msg1);
  ok('pasta sumida: o aviso diz onde reconectar', /Configurações/.test(msg1), msg1);
  ok('pasta sumida: o aviso levanta a causa nº 1 (Drive fechado)', /Google Drive/.test(msg1), msg1);
  ok('pasta sumida: o aviso sai em vermelho', (r.toasts[0] || {}).t === 'err', r.toasts);

  /* ---------- 2. some entre o probe e a gravação ---------- */
  p = pasta({ nome: 'OE 2026', sumidaAoGravar: true });
  r = await clicarSalvar(p);
  const msgs2 = r.toasts.map(t => t.m).join(' | ');
  ok('sumiço na gravação: não diz que salvou', r.ok === false, r.ok);
  ok('sumiço na gravação: mesma instrução, não o erro cru',
     /não encontrada/i.test(msgs2) && !/NotFoundError/.test(msgs2), msgs2);

  /* ---------- 3. pasta boa continua gravando ---------- */
  p = pasta({ nome: 'OE 2026' });
  r = await clicarSalvar(p);
  ok('pasta boa: grava', r.ok === true, r.toasts);
  ok('pasta boa: o arquivo sai com o nome do período',
     p.escritos.length === 1 && p.escritos[0].fn === 'OE-24-08-2026.pdf', p.escritos.map(e => e.fn));
  ok('pasta boa: o aviso final confirma o nome',
     /OE salva: OE-24-08-2026\.pdf/.test(r.toasts.map(t => t.m).join(' | ')), r.toasts);

  /* ---------- 4. clique sobre gravação em curso: fala ---------- */
  p = pasta({ nome: 'OE 2026' });
  r = await clicarSalvar(p, { travada: true });
  ok('gravação em curso: não grava de novo', r.ok === false && p.escritos.length === 0, p.escritos.length);
  ok('gravação em curso: o botão não emudece',
     r.toasts.length > 0 && /aguarde/i.test(r.toasts[0].m), r.toasts);

  /* ---------- 5. o gatilho: pede o Drive e retoma sozinho ---------- */
  p = pasta({ nome: 'OE 2026', sumida: true });
  const g = await clicarSalvar(p, {
    roteiro: async ({ tentarDriveAgora, pedido, modal }) => {
      // O usuário clica "Já abri" com o Drive AINDA fechado: nada retomado.
      await tentarDriveAgora();
      const cedo = { modalAberto: modal.aberto, gravou: p.escritos.length };
      // Agora o Drive sobe de verdade e ele clica de novo.
      p.driveAbriu();
      await tentarDriveAgora();
      return { r: { pedidoInicial: pedido, cedo, modalDepois: modal.aberto }, pedido };
    }
  });
  ok('gatilho: o pedido do Drive aparece na tela', g.ok.pedidoInicial.modalAberto === true, g.ok);
  ok('gatilho: o pedido diz como abrir o Google Drive',
     /Google Drive/.test(g.ok.pedidoInicial.texto) && /Menu Iniciar/.test(g.ok.pedidoInicial.texto),
     g.ok.pedidoInicial.texto);
  ok('gatilho: Drive ainda fechado, nada é retomado',
     g.ok.cedo.modalAberto === true && g.ok.cedo.gravou === 0, g.ok.cedo);
  ok('gatilho: Drive aberto, o pedido sai da tela', g.ok.modalDepois === false, g.ok);
  ok('gatilho: a OE que tinha ficado pelo caminho é gravada',
     p.escritos.length === 1 && p.escritos[0].fn === 'OE-24-08-2026.pdf', p.escritos.map(e => e.fn));

  /* ---------- 6. o porteiro genérico, usado por todas as pastas ---------- */
  // A OE foi só a primeira: PDF da OS, etiquetas, backup, snapshots e exportação
  // passam pelo mesmo pastaProntaOuPedeDrive. Aqui ele é chamado direto.
  const porteiro = new Function('pastaBoa', 'pastaMorta', 'toasts', `
    return (async () => {
      const modal = { aberto: false, texto: '' };
      const document = {
        querySelector: () => null,
        getElementById: () => ({ set innerHTML(v) { modal.texto = v; }, get innerHTML() { return modal.texto; } })
      };
      const openModal = () => { modal.aberto = true; };
      const closeModal = () => { modal.aberto = false; };
      const esc = (v) => String(v == null ? '' : v);
      const toast = (m, t) => toasts.push({ m, t: t || '' });
      let _pedidoDrive = null, _pedidoDriveAutoDado = false;
      ${motor}
      let retomou = 0;
      const boa = await pastaProntaOuPedeDrive(pastaBoa, { oQue: 'x', retomar: () => { retomou++; } });
      const morta = await pastaProntaOuPedeDrive(pastaMorta, { oQue: 'salvar o PDF da OS 0501', retomar: () => { retomou++; } });
      const modalDepoisDaMorta = modal.aberto;
      const textoDaMorta = modal.texto;
      _pararVigiaDrive();
      // Erro no meio da gravação: só o sumiço é tratado como sumiço.
      const naoSumiu = new Error('quota'); naoSumiu.name = 'QuotaExceededError';
      const sumiu = new Error('sumiu'); sumiu.name = 'NotFoundError';
      const reconheceOutro = tratarErroPastaSumiu(naoSumiu, pastaMorta, { oQue: 'x' });
      const reconheceSumico = tratarErroPastaSumiu(sumiu, pastaMorta, { oQue: 'x' });
      _pararVigiaDrive();
      return { boa, morta, modalDepoisDaMorta, textoDaMorta, reconheceOutro, reconheceSumico, retomou };
    })();
  `);
  const t6 = [];
  const q = await porteiro(pasta({ nome: 'PDFs' }), pasta({ nome: 'PDFs', sumida: true }), t6);
  ok('porteiro: pasta boa libera a gravação', q.boa === true, q);
  ok('porteiro: pasta sumida barra a gravação', q.morta === false, q);
  ok('porteiro: pasta sumida abre o pedido do Drive', q.modalDepoisDaMorta === true, q);
  ok('porteiro: o pedido nomeia o que ficou por fazer',
     /salvar o PDF da OS 0501/.test(q.textoDaMorta), q.textoDaMorta);
  ok('porteiro: nada é retomado enquanto a pasta não volta', q.retomou === 0, q.retomou);
  ok('porteiro: erro que não é sumiço segue o caminho normal', q.reconheceOutro === false, q);
  ok('porteiro: sumiço na gravação é reconhecido', q.reconheceSumico === true, q);

  console.log(falhas ? `\n${falhas} falha(s)` : '\nTudo certo.');
  process.exit(falhas ? 1 : 0);
})();
