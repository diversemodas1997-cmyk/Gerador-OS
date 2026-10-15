/* Rode com:  node testes/desenho-foto-os.js

   A OS GUARDA A FOTO DO DESENHO (15/10/2026).

   O 0025 foi duplicado e o ORIGINAL editado para 0050 Azul: as 14 OS dele
   passaram a mostrar o Azul. O que o teste guarda:
     · OS nova fotografa o desenho de agora;
     · editar o cadastro do desenho não muda a OS (imagem, código, SKU, cores);
     · reeditar a OS sem mexer no desenho mantém a foto;
     · escolher outro desenho na OS troca a foto;
     · o botão "Usar a versão atual" (trocar=true) troca a foto;
     · o que não é foto (statusSeq) segue vindo do cadastro;
     · OS sem foto cai no cadastro vivo, e o aviso de "mudou" só acende quando mudou;
     · duplicar desenho: código novo (próximo livre), código repetido é o mesmo.

   Recorta as funções reais do app.js. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const corta = (de) => {
  const a = src.indexOf(de), b = src.indexOf('\n}', a);
  if (a < 0 || b < 0) { console.error('nao achei ' + de + ' no app.js'); process.exit(1); }
  return src.slice(a, b + 2);
};
const STATE = {};
const f = new Function('STATE', [
  'function _fotoDoDesenho', 'function desenhoDaOS', 'function _congelarDesenhoNaOS',
  'function _desenhoMudouDesdeAOS', 'function _proximoCodigoDesenho', 'function _mesmoCodigoDesenho'
].map(corta).join('\n') + '\nreturn { _fotoDoDesenho, desenhoDaOS, _congelarDesenhoNaOS, _desenhoMudouDesdeAOS, _proximoCodigoDesenho, _mesmoCodigoDesenho };')(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

STATE.desenhos = [
  { id: 'D25', codigo: '0025', desc: 'Blusa Moletom Tricolor | Preto/Mostarda/Off-White', img: 'velha.png',
    skuLinha: 'BM.TRI-BEGE', corPrincipalId: 'preto', corSecundariaId: 'mostarda', corTerciariaId: 'off',
    statusSeq: [{ k: 'enfestando' }] },
  { id: 'D24', codigo: '0024', desc: 'Blusa Moletom Tricolor | Verde/Preto/Bege', img: 'verde.png' }
];

const os547 = f._congelarDesenhoNaOS({ id: 'o547', os: '0547', desenhoId: 'D25' }, null);
ok('OS nova fotografa o desenho', os547.desenhoFoto && os547.desenhoFoto.img === 'velha.png', os547.desenhoFoto);
ok('sem mudança, nada a avisar', !f._desenhoMudouDesdeAOS(os547));

// O que aconteceu em 14/10: o cadastro do 0025 vira 0050 Azul.
Object.assign(STATE.desenhos[0], { codigo: '0050', desc: 'Blusa Moletom Tricolor | Azul/Branco/Preto',
  img: 'azul.png', skuLinha: 'BM.TRI-AZUL', corPrincipalId: 'azul', statusSeq: [{ k: 'cortando' }] });
const d = f.desenhoDaOS(os547);
ok('editar o desenho não muda a imagem da OS', d.img === 'velha.png', d.img);
ok('nem o código, a descrição e o SKU', d.codigo === '0025' && /Mostarda/.test(d.desc) && d.skuLinha === 'BM.TRI-BEGE', d);
ok('nem as cores', d.corPrincipalId === 'preto', d.corPrincipalId);
ok('o que não é foto vem do cadastro (statusSeq)', d.statusSeq[0].k === 'cortando', d.statusSeq);
ok('o aviso de "mudou" acende', f._desenhoMudouDesdeAOS(os547));

// Reeditar a OS sem mexer no desenho (o merge do formulário: {...ant, ...data}).
const reed = f._congelarDesenhoNaOS({ ...os547, obs: 'x' }, os547);
ok('reeditar a OS mantém a foto', reed.desenhoFoto.img === 'velha.png', reed.desenhoFoto);

// Trocar à mão.
const botao = f._congelarDesenhoNaOS({ ...os547 }, os547, true);
ok('o botão "usar a versão atual" troca a foto', botao.desenhoFoto.img === 'azul.png', botao.desenhoFoto);
const outro = f._congelarDesenhoNaOS({ ...os547, desenhoId: 'D24' }, os547);
ok('escolher outro desenho troca a foto', outro.desenhoFoto.id === 'D24' && f.desenhoDaOS(outro).img === 'verde.png', outro.desenhoFoto);
const sem = f._congelarDesenhoNaOS({ ...os547, desenhoId: '' }, os547);
ok('tirar o desenho tira a foto', !('desenhoFoto' in sem), sem);

// OS gravada sem foto (aba com o programa antigo): cai no cadastro vivo.
const antiga = { id: 'oX', desenhoId: 'D24' };
ok('OS sem foto lê o cadastro vivo', f.desenhoDaOS(antiga).img === 'verde.png');
ok('OS sem foto não acende o aviso', !f._desenhoMudouDesdeAOS(antiga));
ok('OS muito antiga casa pelo código', f.desenhoDaOS({ codigo: '0024' }, true).id === 'D24');
ok('sem porCodigo não casa pelo código', f.desenhoDaOS({ codigo: '0024' }) === null);

// Desenho que saiu do cadastro: a OS segue com a foto.
STATE.desenhos = STATE.desenhos.filter(x => x.id !== 'D25');
ok('desenho excluído: a OS segue com a foto', f.desenhoDaOS(os547).img === 'velha.png');

// Duplicar desenho.
STATE.desenhos = [{ codigo: '0050' }, { codigo: '035' }, { codigo: '0025 ' }, { codigo: 'Dx7282' }];
ok('próximo código livre', f._proximoCodigoDesenho() === '0051', f._proximoCodigoDesenho());
ok('"0025 " é o mesmo código que "0025"', f._mesmoCodigoDesenho('0025 ', '0025'));
ok('"035" é o mesmo código que "0035"', f._mesmoCodigoDesenho('035', '0035'));
ok('"0025" não é "0050"', !f._mesmoCodigoDesenho('0025', '0050'));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
