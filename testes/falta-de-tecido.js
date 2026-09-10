/* Rode com:  node testes/falta-de-tecido.js

   "NÃO HÁ PANO PARA ESTA OS" — o aviso ao salvar (10/09/2026, Junior: "insira
   no programa um aviso quando o usuario cadastrar uma OS, mas que a quantidade
   de tecido reservado ultrapassa a quantidade necessaria para produzir aquela
   OS... o programa deve avisar se nao tiver tecido suficiente, por conta de os
   tecidos estarem reservados").

   O que este teste protege:

     · o saldo que vale e o DISPONIVEL (entradas − reservado − saidas), e nao o
       que esta no chao do deposito: pano prometido a outra OS ja tem dono;
     · a OS NAO concorre consigo mesma — salvar de novo uma OS que ja reservou
       nao pode acusar falta que nao existe. Este e o caso que um "compara com o
       saldo" ingenuo erra, e o erro aparece como aviso em toda edicao de OS;
     · o STATE volta intacto: a conta mexe em STATE.estoqueMov para ignorar os
       movimentos da propria OS, e deixar isso mexido corromperia o estoque do
       programa inteiro;
     · o texto diz que o pano esta RESERVADO. Sem isso a pessoa vai conferir a
       prateleira, acha o rolo la e conclui que o programa errou.

   Recorta as funcoes do app.js de verdade. */
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

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  // O consumo da OS vem pronto do ctx: quem o calcula (consumoEnfestoOS) tem
  // teste proprio, e o que se prova aqui e a COMPARACAO com o estoque.
  const consumoAgregadoPorTecidoCor = (o) => ctx.consumo[o.id] || [];
  const corSemTecido = (cor, tec) => String(cor || '').replace(new RegExp(' ' + tec + '$'), '');
  ${recorte('function _normNome', 'a normalizacao de nome')}
  const movimentacoesEstoque = () => STATE.estoqueMov || [];
  ${recorte('function calcularSaldosEstoque', 'os saldos do estoque')}
  ${recorte('function faltaDeTecidoParaOS', 'a falta de tecido da OS')}
  ${recorte('function _textoFaltaDeTecido', 'o texto do aviso')}
  return { faltaDeTecidoParaOS, _textoFaltaDeTecido, calcularSaldosEstoque };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const entrada = (tec, cor, kg) => ({ id: 'e' + Math.random(), tipo: 'entrada', tecidoNome: tec, corNome: cor, kg, origem: 'manual' });
const reserva = (tec, cor, kg, osId) => ({ id: 'r' + Math.random(), tipo: 'saida', tecidoNome: tec, corNome: cor, kg, origem: 'os', osId, status: 'reservado' });
const baixa  = (tec, cor, kg, osId) => ({ id: 'b' + Math.random(), tipo: 'saida', tecidoNome: tec, corNome: cor, kg, origem: 'os', osId, status: 'consumido' });

const MALHA = 'Malha Algodão', PRETO = 'Preto Malha Algodão';
const cenario = (mov, consumo) => monta({ STATE: { estoqueMov: mov }, consumo });
const osNova = { id: 'os_nova' };

console.log('-- quando avisa, e quando fica quieto --');
{
  // 100 kg na prateleira, a OS precisa de 60: tem pano.
  let api = cenario([entrada(MALHA, PRETO, 100)], { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] });
  ok('1. com pano sobrando, nao avisa', api.faltaDeTecidoParaOS(osNova).length === 0);

  // Os mesmos 100 kg, mas 80 ja prometidos a OUTRA OS: sobram 20.
  api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 80, 'os_outra')],
                { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] });
  const f = api.faltaDeTecidoParaOS(osNova);
  ok('2. o pano existe, mas esta reservado em outra OS -> AVISA', f.length === 1, f);
  ok('3. e a conta e a do disponivel: 60 pedidos, 20 livres, faltam 40',
     f[0] && f[0].precisa === 60 && f[0].disponivel === 20 && f[0].falta === 40, f[0]);
  ok('4. e diz quanto esta reservado, que e a explicacao do numero',
     f[0] && f[0].reservado === 80, f[0]);

  // Baixa definitiva conta igual: o pano ja saiu.
  api = cenario([entrada(MALHA, PRETO, 100), baixa(MALHA, PRETO, 90, 'os_velha')],
                { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] });
  ok('5. pano ja baixado tambem falta (nao voltou para a prateleira)',
     api.faltaDeTecidoParaOS(osNova)[0].falta === 50);

  // Exatamente o que tem: nao e falta.
  api = cenario([entrada(MALHA, PRETO, 60)], { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] });
  ok('6. o pano exato nao e falta', api.faltaDeTecidoParaOS(osNova).length === 0);

  // Tecido que nunca teve entrada nenhuma: falta tudo.
  api = cenario([entrada(MALHA, PRETO, 100)], { os_nova: [{ tecidoNome: 'Moletom', corNome: 'Bege Moletom', kg: 30 }] });
  const g = api.faltaDeTecidoParaOS(osNova);
  ok('7. tecido sem nenhuma entrada: falta o total pedido',
     g.length === 1 && g[0].falta === 30 && g[0].disponivel === 0, g[0]);
}

console.log('');
console.log('-- a OS nao concorre consigo mesma --');
{
  /* O CASO QUE UM "COMPARA COM O SALDO" INGENUO ERRA. A OS ja reservou os 100
     kg dela; abrir e salvar de novo mostraria "faltam 100 kg", em toda edicao,
     para sempre — e a reserva antiga vai ser substituida pela nova logo depois,
     em aplicarBaixaEstoqueOS. */
  const api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 100, 'os_nova')],
                      { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 100 }] });
  ok('8. salvar de novo a mesma OS nao acusa falta', api.faltaDeTecidoParaOS(osNova).length === 0,
     api.faltaDeTecidoParaOS(osNova));

  // Mas se ela CRESCEU alem do que sobra, a diferenca falta mesmo.
  const api2 = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 100, 'os_nova')],
                       { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 130 }] });
  ok('9. mas a OS que cresceu alem do estoque continua avisando',
     api2.faltaDeTecidoParaOS(osNova)[0].falta === 30, api2.faltaDeTecidoParaOS(osNova)[0]);
}

console.log('');
console.log('-- a conta nao pode deixar o estoque mexido --');
{
  const mov = [entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 80, 'os_nova')];
  const ctx = { STATE: { estoqueMov: mov }, consumo: { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 10 }] } };
  const api = monta(ctx);
  const antes = ctx.STATE.estoqueMov.length;
  api.faltaDeTecidoParaOS(osNova);
  ok('10. STATE.estoqueMov volta inteiro depois da conta',
     ctx.STATE.estoqueMov.length === antes && ctx.STATE.estoqueMov === mov, ctx.STATE.estoqueMov.length);
  ok('11. e os saldos do programa continuam os mesmos',
     api.calcularSaldosEstoque().detalhe[0].reservado === 80);
}

console.log('');
console.log('-- quando o aviso NAO deve existir --');
{
  ok('12. estoque vazio nao gera aviso (ninguem alimenta -> barulho em toda OS)',
     cenario([], { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] })
       .faltaDeTecidoParaOS(osNova).length === 0);
  ok('13. OS que nao consome nada (grade sem medida) nao gera aviso',
     cenario([entrada(MALHA, PRETO, 1)], { os_nova: [] }).faltaDeTecidoParaOS(osNova).length === 0);
  ok('14. sem OS nenhuma, nao quebra',
     cenario([entrada(MALHA, PRETO, 1)], {}).faltaDeTecidoParaOS(null).length === 0);
}

console.log('');
console.log('-- o texto do aviso --');
{
  const api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 80, 'os_outra')],
                      { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 60 }] });
  const txt = api._textoFaltaDeTecido(api.faltaDeTecidoParaOS(osNova));
  ok('15. diz o tecido e a cor', /Malha Algodão · Preto/.test(txt), txt);
  ok('16. diz quanto precisa, quanto ha e quanto falta',
     /precisa 60,000 kg/.test(txt) && /disponível 20,000 kg/.test(txt) && /FALTAM 40,000 kg/.test(txt), txt);
  ok('17. e diz que o pano esta RESERVADO em outras OS — a explicacao do numero',
     /80,000 kg estão reservados em outras OS/.test(txt), txt);
  ok('18. explica que o pano pode estar na prateleira mesmo assim',
     /prateleira/.test(txt) && /prometido a outra OS/.test(txt), txt);
  ok('19. e pergunta, em vez de barrar — a OS ainda pode ser gerada',
     /Gerar a OS assim mesmo\?/.test(txt), txt);
}

console.log('');
console.log('-- ligado no portao de salvar --');
{
  const portao = recorte('function validarAntesDeSalvar', 'o portao de salvar');
  ok('20. validarAntesDeSalvar chama a conta e pergunta ao usuario',
     /const faltando = faltaDeTecidoParaOS\(data\);/.test(portao)
     && /if \(faltando\.length\) return confirm\(_textoFaltaDeTecido\(faltando\)\);/.test(portao), portao.slice(-400));
  ok('21. e vem DEPOIS das checagens que ja existiam (a medida e as camadas)',
     portao.indexOf('fasesSemProvaDeMedida') < portao.indexOf('faltaDeTecidoParaOS')
     && portao.indexOf('calcularLimiteCamadas') < portao.indexOf('faltaDeTecidoParaOS'));
}

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
