/* Rode com:  node testes/celular-so-consulta.js

   NO CELULAR NINGUÉM GRAVA — NEM O ADMIN.

   Decidido em 04/09/2026. O motivo não é permissão, é gravação duplicada: os
   dados da fábrica inteira moram numa linha só (shared_data), e cada save
   reescreve o blob inteiro em ler-alterar-gravar. Duas telas com a MESMA conta
   editando ao mesmo tempo não brigam por um campo — a última a gravar leva a
   linha toda, e o trabalho da outra some sem erro nenhum. O admin é quem tem
   esse poder nos dois aparelhos ao mesmo tempo; isentá-lo seria isentar o
   único caso perigoso.

   O QUE ESTE TESTE PROTEGE são as bordas da detecção, que é onde ela estraga —
   e ela estraga para os dois lados:

     · deixar passar um celular = o que a trava existe para impedir;
     · travar um COMPUTADOR por engano = tirar o trabalho de quem trabalha, e
       cai justamente sobre quem registra compra e folha de OS todo dia.

   Por isso o caso do PAINEL DE TOQUE (tela grande, sem mouse) tem um teste
   próprio: foi o motivo de "não tem mouse" não valer sozinho como sinal.

   E o RECADO é a exceção, liberada no celular a pedido do Junior em 04/09/2026.
   Ele não é dado da fábrica, é conversa: cada recado é uma linha própria que só
   nasce, e não há blob para a segunda tela apagar. O que ele continua exigindo
   é o servidor — recado escrito na nuvem some na passada seguinte do espelho.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}

/* Monta o programa dentro de um aparelho de mentira. navigator, window e screen
   entram como parâmetros: é a única forma de um teste trocar de aparelho sem
   depender do que o Node por acaso define. */
const montar = (amb) => new Function('amb', `
  const navigator = {
    userAgent: amb.ua,
    maxTouchPoints: amb.toques || 0,
    userAgentData: amb.uaData
  };
  const screen = amb.semScreen ? undefined : { width: amb.tela[0], height: amb.tela[1] };
  const window = {
    screen,
    matchMedia: amb.semMatchMedia ? undefined : (q => ({
      matches: q.includes('any-pointer: fine') ? !!amb.temMouse
             : q.includes('pointer: coarse')   ? !!amb.dedo
             : false
    }))
  };
  const console = { log: () => {} };
  let _ehCelular = null;
  ${recorte('function ehCelular', 'a deteccao do aparelho de mao')}

  // O servidor entra como controle do teste: a trava do celular tem de valer
  // MESMO com a fabrica de pe, que e o caso em que ela importa.
  const MODO_LOCAL = 'local';
  let _modoServidor = amb.servidorNoAr ? 'local' : 'nuvem';
  const servidorLocalConfig = () => ({ url: 'https://193.168.0.200', key: 'k' });
  ${recorte('function servidorAceitaGravar', 'a pergunta do servidor')}
  ${recorte('function podeGravar', 'a trava de gravacao')}
  ${recorte('function podeMandarRecado', 'a excecao do recado')}

  const toast = (m) => amb.recados.push(m);
  ${recorte('function _recusarSomenteLeitura', 'a recusa de quem so le')}
  ${recorte('function _recusarRecado', 'a recusa do recado')}

  return { ehCelular, podeGravar, podeMandarRecado, _recusarSomenteLeitura, _recusarRecado };
`)(amb);

/* ------------------------------ os aparelhos ----------------------------- */
const UA = {
  androidChrome: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36',
  iphone: 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1',
  ipadModerno: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
  desktopLinux: 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
  desktopWindows: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
  macDesktop: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
};

const aparelho = (nome, extra) => Object.assign({
  ua: UA.desktopWindows, tela: [1920, 1080], dedo: false, temMouse: true,
  toques: 0, servidorNoAr: true, recados: [], nome
}, extra);

const APARELHOS = [
  // ---- é celular, e não pode gravar ----
  [aparelho('celular Android (Chrome)', {
    ua: UA.androidChrome, tela: [412, 915], dedo: true, temMouse: false, toques: 5,
    uaData: { mobile: true }
  }), true],

  [aparelho('iPhone (Safari)', {
    ua: UA.iphone, tela: [393, 852], dedo: true, temMouse: false, toques: 5
  }), true],

  // O iPad de 2019 para cá se apresenta como Mac. Só a tela de toque o entrega.
  [aparelho('iPad (se diz Mac)', {
    ua: UA.ipadModerno, tela: [820, 1180], dedo: true, temMouse: false, toques: 5
  }), true],

  // "Site para computador" no menu do Chrome do Android: o userAgent inteiro é
  // reescrito e os três primeiros sinais somem. Sobram o vidro e a falta de
  // mouse — que é justamente o par que o menu não reescreve.
  [aparelho('Android em "Site para computador"', {
    ua: UA.desktopLinux, tela: [412, 915], dedo: true, temMouse: false, toques: 5
  }), true],

  // ---- é computador, e TEM de continuar gravando ----
  [aparelho('PC da fábrica (mouse)', {}), false],

  [aparelho('Mac de mesa', { ua: UA.macDesktop, toques: 0 }), false],

  // ESTE É O CASO QUE MANDA NO DESENHO DA DETECÇÃO. Sem mouse e com a tela toda
  // sensível ao toque, mas é um computador — e é onde alguém registra compra e
  // folha de OS. Travá-lo seria pior do que o problema que a trava resolve.
  [aparelho('painel de toque 1920, sem mouse', {
    dedo: true, temMouse: false, toques: 10, tela: [1920, 1080]
  }), false],

  // Um notebook com tela de toque tem os dois: dedo E mouse.
  [aparelho('notebook com tela de toque', {
    dedo: false, temMouse: true, toques: 10, tela: [1536, 864]
  }), false],

  // Navegador velho, sem matchMedia: a falta de uma API não pode travar
  // ninguém. Quem responde aqui é o userAgent, e ele diz computador.
  [aparelho('navegador sem matchMedia', { semMatchMedia: true }), false],
  [aparelho('navegador sem screen', { semScreen: true, dedo: true, temMouse: false }), false],
];

/* -------------------------------- o teste -------------------------------- */
let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

console.log('-- quem e celular, e quem nao e --');
for (const [amb, esperado] of APARELHOS) {
  const app = montar(amb);
  ok(`${amb.nome} -> ${esperado ? 'celular' : 'computador'}`,
     app.ehCelular() === esperado, `deu ${app.ehCelular()}`);
}

console.log('\n-- a trava, com o servidor da fabrica DE PE --');
{
  const celular = montar(aparelho('cel', {
    ua: UA.androidChrome, tela: [412, 915], dedo: true, temMouse: false, servidorNoAr: true
  }));
  ok('celular NAO grava, mesmo com a fabrica no ar', celular.podeGravar() === false);

  const pc = montar(aparelho('pc', { servidorNoAr: true }));
  ok('computador grava, como sempre gravou', pc.podeGravar() === true);

  // A trava vale para o admin porque nao passa por papel nenhum: quem responde
  // e o aparelho. Se um dia alguem acrescentar um "if admin" em podeGravar,
  // este teste cai.
  ok('a trava do celular nao consulta papel nenhum',
     !recorte('function podeGravar', 'a trava').includes('currentRole'));
}

console.log('\n-- o recado tem de apontar para o lugar certo --');
{
  const ambCel = aparelho('cel', {
    ua: UA.iphone, tela: [393, 852], dedo: true, temMouse: false, servidorNoAr: true
  });
  montar(ambCel)._recusarSomenteLeitura('salvar a OS');
  const noCelular = ambCel.recados.join(' ');
  ok('no celular o recado fala do celular', /celular/i.test(noCelular), noCelular);
  ok('no celular o recado manda ir ao computador', /computador/i.test(noCelular), noCelular);
  ok('no celular o recado NAO manda conferir o servidor',
     !/recarregue|servidor está ligado/i.test(noCelular), noCelular);

  // O outro motivo de so-leitura continua com o texto dele: as duas respostas
  // sao opostas, e trocar uma pela outra manda a pessoa atras do problema
  // errado ate desistir.
  const ambPc = aparelho('pc', { servidorNoAr: false });
  montar(ambPc)._recusarSomenteLeitura('salvar a OS');
  const semServidor = ambPc.recados.join(' ');
  ok('sem servidor o recado fala do servidor', /servidor da fábrica/i.test(semServidor), semServidor);
  ok('sem servidor o recado NAO fala de celular', !/celular/i.test(semServidor), semServidor);
}

console.log('\n-- o recado e a excecao --');
{
  const noCelular = aparelho('cel', {
    ua: UA.androidChrome, tela: [412, 915], dedo: true, temMouse: false, servidorNoAr: true
  });
  const cel = montar(noCelular);
  ok('no celular o recado PASSA (e conversa, nao dado da fabrica)',
     cel.podeMandarRecado() === true);
  ok('e a trava dos dados continua de pe no mesmo aparelho',
     cel.podeGravar() === false);
  cel._recusarRecado('mandar recado');
  ok('e o portao do recado nao recusa nada', noCelular.recados.length === 0,
     noCelular.recados.join(' '));

  // O recado dispensa o APARELHO, nunca o SERVIDOR: escrito na copia da nuvem,
  // ele sumiria na passada seguinte do espelho — pior do que nao ter enviado.
  const semServidor = aparelho('cel', {
    ua: UA.androidChrome, tela: [412, 915], dedo: true, temMouse: false, servidorNoAr: false
  });
  const cel2 = montar(semServidor);
  ok('sem o servidor o recado e recusado, mesmo no celular',
     cel2.podeMandarRecado() === false);
  cel2._recusarRecado('mandar recado');
  ok('e a recusa fala do servidor, nao do aparelho',
     /servidor da fábrica/i.test(semServidor.recados.join(' ')),
     semServidor.recados.join(' '));
}

console.log('\n-- os dois funis --');
{
  // Toda recusa passa por um dos dois funis. Se alguem criar um caminho novo
  // que nao passe por nenhum, a trava do celular nao vale naquele caminho — e o
  // furo nao aparece em teste nenhum.
  const dados = (src.match(/_recusarSomenteLeitura\(/g) || []).length - 1;   // menos a definicao
  const recado = (src.match(/_recusarRecado\(/g) || []).length - 1;
  ok(`os portoes de DADOS chamam o funil de dados (${dados})`, dados >= 7, String(dados));
  ok(`os tres portoes de RECADO chamam o funil de recado (${recado})`, recado === 3, String(recado));
  ok('nao sobrou nenhum nome antigo', !src.includes('_recusarPorModoNuvem'));
  // O cloudFlush e a raiz: mesmo que um botao escape, a gravacao do blob para aqui.
  ok('o cloudFlush continua barrado por podeGravar',
     /cloudFlush[\s\S]{0,2000}?if \(!podeGravar\(\)\)/.test(src));
  // O funil do recado NAO pode consultar o aparelho: e exatamente isso que o
  // libera no celular. Se alguem trocar podeMandarRecado por podeGravar ali, o
  // recado volta a ser bloqueado e estas duas linhas caem.
  ok('o funil do recado nao pergunta pelo aparelho',
     !recorte('function _recusarRecado', 'o funil do recado').includes('ehCelular'));
  ok('podeMandarRecado nao pergunta pelo aparelho',
     !recorte('function podeMandarRecado', 'a excecao do recado').includes('ehCelular'));
}

console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
