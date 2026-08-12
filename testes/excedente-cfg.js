/* Rode com:  node testes/excedente-cfg.js

   A REGRA DO EXCEDENTE É CADASTRADA, não escrita no código.

   Em 12/08/2026 a faixa dos 15 cm precisou ir de 8 m para 9 m. Foi um commit,
   um deploy e um Ctrl+Shift+R em cada máquina — para trocar um número que só
   quem enfesta sabe. Agora ela mora em STATE.meta.excedenteCfg, editável em
   Configurações, e o código só guarda o PADRÃO DE FÁBRICA.

   O que este teste protege:

   1. Sem nada cadastrado, vale o padrão. É o caso de todas as instalações no
      dia em que isto subiu, e não pode mudar comportamento nenhum.
   2. Cadastrado, manda o cadastro — inclusive nas exceções (gola e viés).
   3. Config PELA METADE cai no padrão CAMPO A CAMPO. Uma configuração
      incompleta não pode deixar o programa sem regra: ele decide medida de
      pano com isso.
   4. As faixas são ORDENADAS antes de usar. A busca pega a primeira que couber,
      então uma tabela fora de ordem responderia errado sem reclamar de nada —
      e é um erro que não dá sintoma, só número torto no cadastro.

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

// Monta a regra com um cadastro qualquer em STATE.meta.excedenteCfg.
function com(cfg) {
  return new Function('STATE', `
    ${/const EXCEDENTE_ENFESTO_PADRAO_CM = \d+;/.exec(src)[0]}
    ${/const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/.exec(src)[0]}
    ${/const EXCEDENTE_GOLA_CM = \d+;/.exec(src)[0]}
    ${/const EXCEDENTE_VIES_CM = \d+;/.exec(src)[0]}
    ${/const EXCEDENTE_BARRA_CM = \d+;/.exec(src)[0]}
    ${/const _EXC_LIGACAO = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
    ${/const _PAL_VIES = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
    ${/const _PAL_GOLA = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
    ${/const _PAL_BARRA = new Set\(\[[^\]]*\]\);/.exec(src)[0]}
    ${recorte('function _normNome', 'a normalizacao de nome')}
    ${recorte('function _normFaseNome', 'a normalizacao de nome de fase')}
    ${recorte('function _faseSoDe', 'o reconhecedor por nome inteiro')}
    ${recorte('function excedenteCfg', 'a regra cadastrada')}
    ${recorte('function excedentePorComprimento', 'a regra das faixas')}
    ${recorte('function excedenteRegraDaFase', 'a regra inteira da fase')}
    return { excedenteCfg, excedentePorComprimento, excedenteRegraDaFase, EXCEDENTE_FAIXAS };
  `)({ meta: cfg === undefined ? {} : { excedenteCfg: cfg } });
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

console.log('-- sem nada cadastrado: o padrao de fabrica --');
{
  const a = com(undefined);
  const p = a.excedenteCfg();
  ok('1. as faixas sao as do codigo', JSON.stringify(p.faixas) === JSON.stringify(a.EXCEDENTE_FAIXAS),
     p.faixas);
  ok('1b. 0,80 m -> 10 cm', a.excedentePorComprimento('0.80') === 10, a.excedentePorComprimento('0.80'));
  ok('1c. 8,20 m -> 15 cm', a.excedentePorComprimento('8.20') === 15, a.excedentePorComprimento('8.20'));
  ok('1d. 10 m -> 20 cm', a.excedentePorComprimento('10') === 20, a.excedentePorComprimento('10'));
  ok('1e. gola 5 / vies 0', p.gola === 5 && p.vies === 0, p);
}

console.log('\n-- cadastrado, manda o cadastro --');
{
  const a = com({ faixas: [{ ate: 2, cm: 8 }, { ate: 5, cm: 12 }, { ate: 20, cm: 30 }], gola: 7, vies: 3 });
  ok('2. 1,90 m -> 8 cm', a.excedentePorComprimento('1.90') === 8, a.excedentePorComprimento('1.90'));
  ok('2b. 4 m -> 12 cm', a.excedentePorComprimento('4') === 12, a.excedentePorComprimento('4'));
  ok('2c. 15 m -> 30 cm (e nao "acima da tabela")', a.excedentePorComprimento('15') === 30,
     a.excedentePorComprimento('15'));
  ok('2d. 21 m -> null (passou do ultimo limite)', a.excedentePorComprimento('21') === null,
     a.excedentePorComprimento('21'));
  ok('2e. a gola cadastrada vale', a.excedenteRegraDaFase({ nome: 'Gola', excedente: '' }, 4) === 7,
     a.excedenteRegraDaFase({ nome: 'Gola', excedente: '' }, 4));
  ok('2f. o vies cadastrado vale', a.excedenteRegraDaFase({ nome: 'Viés', excedente: '' }, 4) === 3,
     a.excedenteRegraDaFase({ nome: 'Viés', excedente: '' }, 4));
  ok('2g. a borda continua na faixa de baixo (2 m exatos -> 8)',
     a.excedentePorComprimento('2') === 8, a.excedentePorComprimento('2'));
}

console.log('\n-- a mudanca de 8 para 9 m, agora sem commit --');
{
  const oito = com({ faixas: [{ ate: 1.5, cm: 10 }, { ate: 8, cm: 15 }, { ate: 12, cm: 20 }] });
  const nove = com({ faixas: [{ ate: 1.5, cm: 10 }, { ate: 9, cm: 15 }, { ate: 12, cm: 20 }] });
  ok('3. com limite 8, o corpo de 8,20 leva 20', oito.excedentePorComprimento('8.20') === 20,
     oito.excedentePorComprimento('8.20'));
  ok('3b. com limite 9, leva 15', nove.excedentePorComprimento('8.20') === 15,
     nove.excedentePorComprimento('8.20'));
  ok('3c. e as excecoes seguem no padrao, que nao foi cadastrado',
     nove.excedenteCfg().gola === 5 && nove.excedenteCfg().vies === 0, nove.excedenteCfg());
}

console.log('\n-- config PELA METADE cai no padrao, campo a campo --');
{
  ok('4. so gola cadastrada: vies fica no padrao',
     com({ gola: 9 }).excedenteCfg().vies === 0, com({ gola: 9 }).excedenteCfg());
  ok('4b. so gola cadastrada: as faixas ficam no padrao',
     com({ gola: 9 }).excedentePorComprimento('8.20') === 15);
  ok('4c. faixas vazias caem no padrao', com({ faixas: [] }).excedentePorComprimento('0.8') === 10,
     com({ faixas: [] }).excedentePorComprimento('0.8'));
  ok('4d. faixas com lixo dentro caem no padrao',
     com({ faixas: [{ ate: 'abc', cm: 'x' }] }).excedentePorComprimento('0.8') === 10);
  ok('4e. faixa com limite zero e descartada',
     com({ faixas: [{ ate: 0, cm: 99 }] }).excedentePorComprimento('0.8') === 10);
  ok('4f. objeto vazio cai no padrao inteiro',
     com({}).excedentePorComprimento('8.20') === 15 && com({}).excedenteCfg().gola === 5);
  ok('4g. null cai no padrao', com(null).excedentePorComprimento('8.20') === 15);
  ok('4h. gola invalida cai no padrao', com({ gola: 'abc' }).excedenteCfg().gola === 5);
  ok('4i. gola NEGATIVA cai no padrao (nao existe sobra negativa)',
     com({ gola: -5 }).excedenteCfg().gola === 5, com({ gola: -5 }).excedenteCfg().gola);
  ok('4j. gola ZERO cadastrada vale (zero e um valor)',
     com({ gola: 0 }).excedenteCfg().gola === 0, com({ gola: 0 }).excedenteCfg().gola);
}

console.log('\n-- as faixas sao ordenadas antes de usar --');
{
  // Cadastradas fora de ordem, a busca pegaria a primeira que coubesse e a
  // tabela responderia errado calada. Ordenar e o que impede isso.
  const a = com({ faixas: [{ ate: 12, cm: 20 }, { ate: 1.5, cm: 10 }, { ate: 9, cm: 15 }] });
  ok('5. 0,80 m -> 10 (e nao 20, da faixa que veio primeiro)',
     a.excedentePorComprimento('0.80') === 10, a.excedentePorComprimento('0.80'));
  ok('5b. 5 m -> 15', a.excedentePorComprimento('5') === 15, a.excedentePorComprimento('5'));
  ok('5c. 11 m -> 20', a.excedentePorComprimento('11') === 20, a.excedentePorComprimento('11'));
  ok('5d. a tabela sai ordenada', a.excedenteCfg().faixas.map(f => f.ate).join(',') === '1.5,9,12',
     a.excedenteCfg().faixas);
}

console.log('\n-- virgula decimal, como se digita na tela --');
{
  const a = com({ faixas: [{ ate: '1,80', cm: '12' }, { ate: '9', cm: '15' }], gola: '7' });
  ok('6. limite com virgula funciona', a.excedentePorComprimento('1.70') === 12,
     a.excedentePorComprimento('1.70'));
  ok('6b. e 1,90 ja cai na faixa seguinte', a.excedentePorComprimento('1.90') === 15,
     a.excedentePorComprimento('1.90'));
  ok('6c. gola em texto tambem', a.excedenteCfg().gola === 7, a.excedenteCfg().gola);
}

console.log('\n-- a terceira excecao: barra/punhos --');
{
  const p = com(undefined);
  ok('8. no padrao, barra/punhos leva 15', p.excedenteCfg().barra === 15, p.excedenteCfg().barra);
  const a = com({ barra: 22 });
  ok('8b. cadastrada, manda o cadastro',
     a.excedenteRegraDaFase({ nome: 'Barra/Punhos', excedente: '' }, 1.55) === 22,
     a.excedenteRegraDaFase({ nome: 'Barra/Punhos', excedente: '' }, 1.55));
  ok('8c. e nao depende do comprimento',
     a.excedenteRegraDaFase({ nome: 'Barra/Punhos', excedente: '' }, 11) === 22);
  ok('8d. so ela cadastrada: gola e vies ficam no padrao',
     a.excedenteCfg().gola === 5 && a.excedenteCfg().vies === 0, a.excedenteCfg());
  ok('8e. barra invalida cai no padrao', com({ barra: 'x' }).excedenteCfg().barra === 15);
  ok('8f. barra ZERO cadastrada vale', com({ barra: 0 }).excedenteCfg().barra === 0);
}

console.log('\n-- duas faixas so, que e um cadastro legitimo --');
{
  const a = com({ faixas: [{ ate: 2, cm: 10 }, { ate: 12, cm: 18 }] });
  ok('7. 1,50 -> 10', a.excedentePorComprimento('1.50') === 10);
  ok('7b. 7 m -> 18', a.excedentePorComprimento('7') === 18);
  ok('7c. 13 m -> null', a.excedentePorComprimento('13') === null);
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
