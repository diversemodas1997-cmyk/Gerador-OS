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
  // O aviso fala em bobina desde 14/09/2026; o peso da bobina vem do cadastro
  // do tecido ou das entradas anteriores (pesoBobinaEstimado).
  const _normFake = (x) => _normNome(x);
  ${recorte('function pesoBobinaPorNome', 'o peso de bobina cadastrado')}
  ${recorte('function pesoBobinaEstimado', 'o peso de bobina estimado')}
  ${recorte('function _bobinasDoKg', 'a conversao de kg em bobinas')}
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
const cenario = (mov, consumo, tecidos) => monta({
  STATE: { estoqueMov: mov, tecidos: tecidos || [] }, consumo });
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


/* ---------------------------------------------------------------------------
   O VERMELHO NA LISTA DE MATERIAL RESERVADO (14/09/2026, Junior).

   O aviso da gravacao passa: some do modal e ninguem mais sabe qual das OS da
   lista esta segurando pano que a prateleira nao tem. Ele ficou de pe na linha
   da OS, em vermelho, e a conta e ESTA MESMA — duas contas para a mesma
   pergunta dariam, mais cedo ou mais tarde, duas respostas.

   E ele tem de SAIR SOZINHO quando a entrada daquele tecido for lancada. Por
   isso nada e gravado na OS: a falta e perguntada a cada desenho da tela. Marca
   gravada teria de ser apagada a mao e ficaria vermelha para sempre numa OS que
   ja tem o pano dela.
   --------------------------------------------------------------------------- */
console.log('');
console.log('-- o vermelho na lista de material reservado --');
{
  // A OS 0526 ja esta SALVA: a reserva dela ja esta no razao. 100 kg entraram,
  // 130 estao reservados nela: faltam 30 que ninguem comprou.
  const consumo = { os_0526: [{ tecidoNome: MALHA, corNome: PRETO, kg: 130 }] };
  let api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 130, 'os_0526')], consumo);
  const os0526 = { id: 'os_0526' };
  const f = api.faltaDeTecidoParaOS(os0526);
  ok('22. a OS ja salva que nao cabe continua acusando falta (linha vermelha)',
     f.length === 1 && f[0].falta === 30, f);

  // Lancada a entrada que faltava, a linha volta ao preto — sem tocar na OS.
  api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 130, 'os_0526'),
                 entrada(MALHA, PRETO, 30)], consumo);
  ok('23. lancada a entrada do tecido, o vermelho sai sozinho',
     api.faltaDeTecidoParaOS(os0526).length === 0, api.faltaDeTecidoParaOS(os0526));

  // Entrada menor que a falta nao basta: 10 dos 30 nao apagam o aviso.
  api = cenario([entrada(MALHA, PRETO, 100), reserva(MALHA, PRETO, 130, 'os_0526'),
                 entrada(MALHA, PRETO, 10)], consumo);
  ok('24. entrada PARCIAL nao apaga o vermelho — ainda faltam 20',
     api.faltaDeTecidoParaOS(os0526).length === 1
     && api.faltaDeTecidoParaOS(os0526)[0].falta === 20, api.faltaDeTecidoParaOS(os0526));

  /* DUAS OS NA MESMA PRATELEIRA CURTA: AS DUAS ACENDEM.

     100 kg na prateleira, uma OS quer 40 e outra quer 130. Nao existe "a que
     cabe": o pano nao serve as duas, e quem chegar primeiro ao enfesto leva.
     Dizer que so a maior esta em falta seria escolher uma ordem — por data? por
     numero? — que a fabrica nunca combinou, e mandaria a outra OS para a mesa
     confiante num pano que pode nao estar la.

     E a mesma resposta que o aviso da gravacao ja dava, o que e o ponto: as duas
     telas contam a mesma coisa. */
  api = cenario([entrada(MALHA, PRETO, 100),
                 reserva(MALHA, PRETO, 40, 'os_boa'), reserva(MALHA, PRETO, 130, 'os_0526')],
                { os_boa: [{ tecidoNome: MALHA, corNome: PRETO, kg: 40 }], ...consumo });
  ok('25. prateleira que nao serve as duas: as DUAS linhas acendem',
     api.faltaDeTecidoParaOS({ id: 'os_boa' }).length === 1
     && api.faltaDeTecidoParaOS(os0526).length === 1,
     [api.faltaDeTecidoParaOS({ id: 'os_boa' }), api.faltaDeTecidoParaOS(os0526)]);

  // Mas uma OS em OUTRA prateleira, essa com pano, nao e contagiada.
  const VERDE = 'Verde Malha Algodão';
  api = cenario([entrada(MALHA, PRETO, 100), entrada(MALHA, VERDE, 500),
                 reserva(MALHA, PRETO, 130, 'os_0526'), reserva(MALHA, VERDE, 80, 'os_verde')],
                { os_verde: [{ tecidoNome: MALHA, corNome: VERDE, kg: 80 }], ...consumo });
  ok('25b. a OS de outra cor, com pano de sobra, continua preta',
     api.faltaDeTecidoParaOS({ id: 'os_verde' }).length === 0,
     api.faltaDeTecidoParaOS({ id: 'os_verde' }));
}

/* A TELA usa esta conta, e nao grava marca nenhuma. */
{
  const tela = recorte('function renderEstoque', 'a tela do estoque de tecidos');
  ok('26. a lista de material reservado pergunta a falta por OS',
     /faltaPorOS\.set\(r\.osId, f\)/.test(tela) && /faltaDeTecidoParaOS\(os\)/.test(tela), '');
  // Desde 24/09/2026 a linha NAO e pintada inteira: so a celula da fase cujo
  // pano falta (a ribana da gola das OS rosa parecia falta de malha).
  ok('27. a linha nao e mais pintada inteira de vermelho',
     !/falta \? ' style="color:#c0392b;"/.test(tela), '');
  ok('27b. a celula da fase cujo pano falta e que fica vermelha',
     /emFalta \? 'background:#fbe6e6;color:#c0392b;'/.test(tela), '');
  ok('27c. e o selo diz qual pano falta',
     /falta \$\{panos\}/.test(tela), '');
  ok('28. a falta NAO e gravada na OS — nada de marca a limpar depois',
     !/\.faltaPano\s*=/.test(tela) && !/o\.semPano\s*=/.test(tela), '');
}


/* O AVISO FALA EM BOBINA, ALEM DO QUILO (14/09/2026, Junior).
   "Faltam 69,094 kg" nao diz a ninguem se o problema e meia bobina ou um
   caminhao. A bobina que FALTA arredonda para cima e a DISPONIVEL para baixo:
   meia bobina que falta obriga a comprar uma inteira, e meia bobina na
   prateleira ninguem vai buscar. */
console.log('');
console.log('-- o aviso em bobinas --');
{
  const TEC = [{ id: 't1', nome: MALHA, pesoBobina: 19 }];
  // 100 kg na prateleira, a OS quer 200: faltam 100 = 5,26 bobinas.
  const api = cenario([entrada(MALHA, PRETO, 100)],
    { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 200 }] }, TEC);
  const txt = api._textoFaltaDeTecido(api.faltaDeTecidoParaOS(osNova));
  ok('29. o que FALTA vem em bobinas, arredondado para cima (5,26 -> 6)',
     /FALTAM 100,000 kg \(6 bobinas\)/.test(txt), txt);
  ok('30. o DISPONIVEL vem em bobinas, arredondado para baixo (5,26 -> 5)',
     /disponível 100,000 kg \(5 bobinas\)/.test(txt), txt);
  ok('31. e o que a OS precisa tambem sai em bobinas',
     /precisa 200,000 kg \(11 bobinas\)/.test(txt), txt);

  // Sem peso de bobina conhecido, o aviso volta a falar so em quilos.
  const semPeso = cenario([entrada(MALHA, PRETO, 100)],
    { os_nova: [{ tecidoNome: MALHA, corNome: PRETO, kg: 200 }] }, [{ id: 't1', nome: MALHA }]);
  const txt2 = semPeso._textoFaltaDeTecido(semPeso.faltaDeTecidoParaOS(osNova));
  ok('32. pano sem peso de bobina conhecido: so quilos, sem bobina inventada',
     !/bobina\(s\)|\(\d+ bobinas?\)/.test(txt2) && /FALTAM 100,000 kg/.test(txt2), txt2);
  ok('33. ... e o aviso diz como fazer a bobina aparecer',
     /peso médio no cadastro do tecido/.test(txt2), txt2);
}

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
