/*
 * ETIQUETAS DE SÃO CARLOS: BAIXA DAS OS COSTURADAS SEM A COSTURA MARCADA.
 *
 * Junior, 14/10/2026: "Considerando que todas as OS foram costuradas em São
 * Carlos, determine quantas etiquetas deveriam haver em residual", "Considera
 * apenas OS que foram expedidas para São Carlos" e "Atualize o estoque de
 * etiquetas unidade são carlos com esses números".
 *
 * O programa baixa a etiqueta da OS quando vê a costura (a caixa, o status
 * Costurando ou, desde 654c95e, a passagem por Costurando no diário). Estas 23
 * OS foram expedidas para São Carlos, cortadas/ensacadas depois de 25/09 e já
 * estão adiante da costura, mas não têm rastro nenhum dela — então nenhuma
 * unidade baixava as etiquetas delas. Entra uma SAÍDA de São Carlos por OS e
 * tamanho, com a data do primeiro registro de status depois da costura (o dia
 * em que ela já estava costurada) e o osId, para dar para conferir uma a uma.
 *
 * Não repete: lançamento já gravado com o mesmo id fica como está.
 *
 *   node servidor/etiquetas-sc-1410.js            (so mostra)
 *   node servidor/etiquetas-sc-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();

const LISTA = [
  {"osId":"id_1789587705056_554","os":"0546","tam":"P","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"M","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"G","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"GG","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"G1","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"G2","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789587705056_554","os":"0546","tam":"G3","cor":"Preto","qtd":32,"data":"2026-10-05"},
  {"osId":"id_1789644375013_585","os":"0548","tam":"M","cor":"Preto","qtd":56,"data":"2026-10-05"},
  {"osId":"id_1789644375013_585","os":"0548","tam":"G","cor":"Preto","qtd":56,"data":"2026-10-05"},
  {"osId":"id_1789644375013_585","os":"0548","tam":"GG","cor":"Preto","qtd":56,"data":"2026-10-05"},
  {"osId":"id_1790079109820_54","os":"0564","tam":"G","cor":"Preto","qtd":128,"data":"2026-09-30"},
  {"osId":"id_1790079109820_54","os":"0564","tam":"GG","cor":"Preto","qtd":128,"data":"2026-09-30"},
  {"osId":"id_1790079109820_54","os":"0564","tam":"G1","cor":"Preto","qtd":128,"data":"2026-09-30"},
  {"osId":"id_1790079109820_54","os":"0564","tam":"G2","cor":"Preto","qtd":128,"data":"2026-09-30"},
  {"osId":"id_1790079109820_54","os":"0564","tam":"G3","cor":"Preto","qtd":128,"data":"2026-09-30"},
  {"osId":"id_1790079220004_722","os":"0565","tam":"M","cor":"Preto","qtd":160,"data":"2026-09-30"},
  {"osId":"id_1790079220004_722","os":"0565","tam":"G","cor":"Preto","qtd":160,"data":"2026-09-30"},
  {"osId":"id_1790079220004_722","os":"0565","tam":"GG","cor":"Preto","qtd":160,"data":"2026-09-30"},
  {"osId":"id_1790079220004_722","os":"0565","tam":"G1","cor":"Preto","qtd":160,"data":"2026-09-30"},
  {"osId":"id_1790174525426_492","os":"0568","tam":"P","cor":"Preto","qtd":162,"data":"2026-09-30"},
  {"osId":"id_1790174525426_492","os":"0568","tam":"G","cor":"Preto","qtd":162,"data":"2026-09-30"},
  {"osId":"id_1790174525426_492","os":"0568","tam":"G2","cor":"Preto","qtd":162,"data":"2026-09-30"},
  {"osId":"id_1790174525426_492","os":"0568","tam":"G3","cor":"Preto","qtd":162,"data":"2026-09-30"},
  {"osId":"id_1790174585611_990","os":"0569","tam":"M","cor":"Preto","qtd":164,"data":"2026-09-30"},
  {"osId":"id_1790174585611_990","os":"0569","tam":"G","cor":"Preto","qtd":164,"data":"2026-09-30"},
  {"osId":"id_1790174585611_990","os":"0569","tam":"GG","cor":"Preto","qtd":164,"data":"2026-09-30"},
  {"osId":"id_1790174585611_990","os":"0569","tam":"G1","cor":"Preto","qtd":164,"data":"2026-09-30"},
  {"osId":"id_1790174639513_584","os":"0570","tam":"G","cor":"Preto","qtd":328,"data":"2026-09-30"},
  {"osId":"id_1790174639513_584","os":"0570","tam":"G2","cor":"Preto","qtd":164,"data":"2026-09-30"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"P","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"M","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"G","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"GG","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"G1","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"G2","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257287115_623","os":"0571","tam":"G3","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"P","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"M","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"G","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"GG","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"G1","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"G2","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257295860_104","os":"0572","tam":"G3","cor":"Preto","qtd":52,"data":"2026-10-05"},
  {"osId":"id_1790257662955_167","os":"0574","tam":"M","cor":"Preto","qtd":68,"data":"2026-10-05"},
  {"osId":"id_1790257662955_167","os":"0574","tam":"G","cor":"Preto","qtd":136,"data":"2026-10-05"},
  {"osId":"id_1790257662955_167","os":"0574","tam":"GG","cor":"Preto","qtd":68,"data":"2026-10-05"},
  {"osId":"id_1790257671131_377","os":"0573","tam":"M","cor":"Preto","qtd":96,"data":"2026-10-05"},
  {"osId":"id_1790257671131_377","os":"0573","tam":"G","cor":"Preto","qtd":192,"data":"2026-10-05"},
  {"osId":"id_1790257671131_377","os":"0573","tam":"GG","cor":"Preto","qtd":96,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"P","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"M","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"G","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"GG","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"G1","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"G2","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365332493_666","os":"0577","tam":"G3","cor":"Preto","qtd":140,"data":"2026-10-05"},
  {"osId":"id_1790365541733_66","os":"0578","tam":"G","cor":"Preto","qtd":144,"data":"2026-10-05"},
  {"osId":"id_1790365541733_66","os":"0578","tam":"GG","cor":"Preto","qtd":144,"data":"2026-10-05"},
  {"osId":"id_1790365541733_66","os":"0578","tam":"G1","cor":"Preto","qtd":144,"data":"2026-10-05"},
  {"osId":"id_1790365541733_66","os":"0578","tam":"G2","cor":"Preto","qtd":144,"data":"2026-10-05"},
  {"osId":"id_1790365541733_66","os":"0578","tam":"G3","cor":"Preto","qtd":144,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"P","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"M","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"G","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"GG","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"G1","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"G2","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366240357_833","os":"0580","tam":"G3","cor":"Preto","qtd":8,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"P","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"M","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"G","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"GG","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"G1","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"G2","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366341260_570","os":"0581","tam":"G3","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"P","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"M","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"G","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"GG","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"G1","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"G2","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366540557_848","os":"0583","tam":"G3","cor":"Preto","qtd":24,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"P","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"M","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"G","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"GG","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"G1","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"G2","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790366672635_887","os":"0585","tam":"G3","cor":"Preto","qtd":14,"data":"2026-10-05"},
  {"osId":"id_1790597417119_96","os":"0586","tam":"M","cor":"Preto","qtd":158,"data":"2026-10-05"},
  {"osId":"id_1790597417119_96","os":"0586","tam":"G","cor":"Preto","qtd":158,"data":"2026-10-05"},
  {"osId":"id_1790597417119_96","os":"0586","tam":"G1","cor":"Preto","qtd":158,"data":"2026-10-05"},
  {"osId":"id_1790597417119_96","os":"0586","tam":"G3","cor":"Preto","qtd":158,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"P","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"M","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"G","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"GG","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"G1","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"G2","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790615147537_162","os":"0587","tam":"G3","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790624073001_500","os":"0588","tam":"M","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790624073001_500","os":"0588","tam":"G","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790624073001_500","os":"0588","tam":"G1","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790624073001_500","os":"0588","tam":"G3","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790686556796_978","os":"0590","tam":"G","cor":"Preto","qtd":116,"data":"2026-10-05"},
  {"osId":"id_1790686556796_978","os":"0590","tam":"GG","cor":"Preto","qtd":116,"data":"2026-10-05"},
  {"osId":"id_1790686556796_978","os":"0590","tam":"G1","cor":"Preto","qtd":116,"data":"2026-10-05"},
  {"osId":"id_1790686556796_978","os":"0590","tam":"G2","cor":"Preto","qtd":116,"data":"2026-10-05"},
  {"osId":"id_1790686556796_978","os":"0590","tam":"G3","cor":"Preto","qtd":116,"data":"2026-10-05"},
  {"osId":"id_1790691631602_574","os":"0593","tam":"G","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790691631602_574","os":"0593","tam":"GG","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790691631602_574","os":"0593","tam":"G1","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790691631602_574","os":"0593","tam":"G2","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790691631602_574","os":"0593","tam":"G3","cor":"Preto","qtd":4,"data":"2026-10-05"},
  {"osId":"id_1790708344394_923","os":"0594","tam":"G","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790708344394_923","os":"0594","tam":"GG","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790708344394_923","os":"0594","tam":"G1","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790708344394_923","os":"0594","tam":"G2","cor":"Preto","qtd":2,"data":"2026-10-05"},
  {"osId":"id_1790708344394_923","os":"0594","tam":"G3","cor":"Preto","qtd":2,"data":"2026-10-05"}
];

async function req(method, caminho, body) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: { apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' },
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t;
}
const arr = s => (typeof s === 'string' ? JSON.parse(s || '[]') : (s || []));

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const movs = arr(data.aviamentosMov);
  const ordens = arr(data.ordens);
  const agora = new Date().toISOString();
  const novos = LISTA.map(x => ({
    id: 'ajuste-sc-1410:' + x.osId + ':' + x.tam,
    tipo: 'saida', unidade: 'sc', item: 'Etiqueta', tam: x.tam, cor: x.cor, kg: 0, qtd: x.qtd, data: x.data,
    osId: x.osId, obs: 'Costura da OS ' + x.os + ' em São Carlos (sem a costura marcada; ajuste de 14/10/2026)',
    por: 'ajuste 14/10/2026', em: agora
  })).filter(m => !movs.some(y => y.id === m.id));
  const faltaOS = [...new Set(LISTA.map(x => x.osId))].filter(id => !ordens.some(o => o.id === id));
  if (faltaOS.length) throw new Error('OS não encontradas: ' + faltaOS.join(', '));
  const soma = {};
  novos.forEach(m => { soma[m.tam] = (soma[m.tam] || 0) + m.qtd; });
  console.log(`${novos.length} saídas novas de ${[...new Set(novos.map(m => m.osId))].length} OS. Por tamanho: ${JSON.stringify(soma)}`);
  if (!novos.length) { console.log('Nada a fazer.'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-etiquetas-sc-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('backup: ' + path.relative(RAIZ, arq));
  const saida = movs.concat(novos);
  const novo = Object.assign({}, data, {
    aviamentosMov: typeof data.aviamentosMov === 'string' ? JSON.stringify(saida) : saida,
    _device: 'etiquetas-sc-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const m2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.aviamentosMov);
  console.log(`gravado: ${m2.filter(m => String(m.id).startsWith('ajuste-sc-1410:')).length} saídas de ajuste no blob (${m2.length} lançamentos).`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
