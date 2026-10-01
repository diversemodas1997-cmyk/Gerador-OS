/* Rode com:  node testes/fechados-baixa-os.js

   A BAIXA DA OS DESCONTA AS BOBINAS FECHADAS (01/10/2026).

   O Moletom Preto entrou com 198 kg em 11 bobinas (18 kg cada); a OS 0592
   baixou 91,661 kg — cinco bobinas. Protege:
     · o peso da bobina sai das entradas CONTADAS da cor (kg ÷ fechados);
     · a largura da grade da OS escolhe a contagem daquela largura;
     · sem contagem na cor, cai no peso estimado do tecido;
     · sem peso nenhum, não desconta nada (não inventa bobina).

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + de); process.exit(1); }
  return src.slice(i, src.indexOf('\n}', i) + 2);
}
const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const _normNome = s => String(s || '').trim().toLowerCase();
  const movimentacoesEstoque = () => STATE.estoqueMov;
  const _larguraDaOSNoTecido = (o) => (o && o.larg) || 0;
  const pesoBobinaEstimado = () => ctx.estimado || null;
  ${recorte('function _pesoBobinaDaBaixa')}
  ${recorte('function _fechadosDaBaixa')}
  return { _fechadosDaBaixa, _pesoBobinaDaBaixa };
`)(ctx);

let falhas = 0;
const ok = (nome, obtido, esperado) => {
  if (obtido === esperado) console.log('ok  ' + nome);
  else { falhas++; console.log('FALHA ' + nome + '\n      obtido: ' + obtido + '  esperado: ' + esperado); }
};
const ent = (kg, fechados, largura, cor = 'Preto Moletom') =>
  ({ tipo: 'entrada', origem: 'manual', tecidoNome: 'Moletom', corNome: cor, kg, fechados, largura });
const baixa = (kg, osId = 'os1') => ({ tipo: 'saida', origem: 'os', osId, tecidoNome: 'Moletom', corNome: 'Preto Moletom', kg });

{
  const api = monta({ STATE: { ordens: [{ id: 'os1' }], estoqueMov: [ent(1518, 0, 120), ent(198, 11, 120)] } });
  ok('OS 0592: 91,661 kg de bobinas de 18 kg = 5', api._fechadosDaBaixa(baixa(91.661)), 5);
  ok('0548: 30,267 kg = 2', api._fechadosDaBaixa(baixa(30.267)), 2);
  ok('0589: 2,363 kg = 0 (retalho nao e bobina)', api._fechadosDaBaixa(baixa(2.363)), 0);
}
{
  const api = monta({ STATE: { ordens: [{ id: 'os1', larg: 80 }], estoqueMov: [ent(180, 10, 120), ent(130, 10, 80)] } });
  ok('grade de 80 cm usa a bobina de 80 cm (13 kg): 26 kg = 2', api._fechadosDaBaixa(baixa(26)), 2);
}
{
  const api = monta({ STATE: { ordens: [{ id: 'os1' }], estoqueMov: [ent(100, 0, 120)] }, estimado: { kg: 20 } });
  ok('sem contagem na cor: peso estimado do tecido (20 kg): 60 kg = 3', api._fechadosDaBaixa(baixa(60)), 3);
}
{
  const api = monta({ STATE: { ordens: [{ id: 'os1' }], estoqueMov: [ent(100, 0, 120)] } });
  ok('sem peso nenhum: nao desconta', api._fechadosDaBaixa(baixa(60)), 0);
}
if (falhas) { console.log(falhas + ' falha(s)'); process.exit(1); }
console.log('tudo certo');
