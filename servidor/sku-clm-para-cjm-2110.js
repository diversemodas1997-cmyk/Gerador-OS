/* SKU DA CALÇA MOLETOM: CLM → CJM (21/10/2026, Junior: "Corrija o sku para CJM").

   A grade "P ao G3 | CJM | 177cm" (cadastrada hoje pelo risco do CAD, que chama
   a peça de CJM) não aparecia na Nova OS: o desenho 0037 e o modelo Calça
   Moletom diziam CLM, e o filtro de grades da OS esconde a grade de outro SKU.
   O SKU certo é o do CAD, CJM. Esta correção:
     · modelo Calça Moletom e desenho 0037: skuLinha CLM → CJM;
     · grades "P-M-G-GG | CLM | 177cm" e "M-2G-GG | CLM | 177cm" → "| CJM |";
     · grade "P ao G3 | CJM | 177cm": a fase Corpo estava sem tecido — recebe o
       Moletom, como as outras duas;
     · OS 0589 (a única da calça): a cópia do desenho e o nome da grade gravados
       nela.
   A etapa "Costura CLM | Descalvado" é nome de etapa (a facção), não SKU: fica.

   Rode:  node servidor/sku-clm-para-cjm-2110.js            (só mostra)
          node servidor/sku-clm-para-cjm-2110.js --gravar   (grava, com backup) */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const MODELO = 'id_1790677966501_105';
const DESENHO = 'id_1790681903269_800';
const GRADES_CLM = ['id_1790679404154_678', 'id_1792601516247_940'];
const GRADE_CJM = 'id_1792601226352_810';
const MOLETOM = 'id_1776868289700_662';

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
const volta = (orig, v) => (typeof orig === 'string' ? JSON.stringify(v) : v);
const trocaNome = s => String(s).replace('| CLM |', '| CJM |');

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const modelos = arr(data.modelos), desenhos = arr(data.desenhos), grades = arr(data.grades), ordens = arr(data.ordens);

  const m = modelos.find(x => x.id === MODELO);
  if (!m || m.skuLinha !== 'CLM') throw new Error('modelo Calça Moletom não está com CLM: ' + (m && m.skuLinha));
  const d = desenhos.find(x => x.id === DESENHO);
  if (!d || d.skuLinha !== 'CLM') throw new Error('desenho 0037 não está com CLM: ' + (d && d.skuLinha));
  const outros = desenhos.filter(x => x.id !== DESENHO && x.skuLinha && /^CLM\b/.test(x.skuLinha));
  if (outros.length) throw new Error('há outros desenhos com CLM: ' + outros.map(x => x.codigo).join(', '));

  console.log(`modelo ${m.nome}: CLM -> CJM`); m.skuLinha = 'CJM';
  console.log(`desenho ${d.codigo} ${d.desc}: CLM -> CJM`); d.skuLinha = 'CJM';
  GRADES_CLM.forEach(id => {
    const g = grades.find(x => x.id === id);
    if (!g || !/\| CLM \|/.test(g.nome)) throw new Error('grade ' + id + ' mudou: ' + (g && g.nome));
    const novo = trocaNome(g.nome);
    if (grades.some(x => x.id !== id && x.nome === novo)) throw new Error('já existe a grade ' + novo);
    console.log(`grade ${g.nome} -> ${novo}`); g.nome = novo;
  });
  const gc = grades.find(x => x.id === GRADE_CJM);
  if (!gc) throw new Error('grade P ao G3 | CJM não encontrada');
  const corpo = (gc.fases || []).find(f => f.nome === 'Corpo');
  if (corpo && !corpo.tecidoId) { console.log(`grade ${gc.nome}: fase Corpo sem tecido -> Moletom`); corpo.tecidoId = MOLETOM; }

  ordens.filter(o => o.modeloId === MODELO || o.desenhoId === DESENHO).forEach(o => {
    if (o.grade && /\| CLM \|/.test(o.grade.descricao || '')) { console.log(`OS ${o.os}: grade ${o.grade.descricao} -> ${trocaNome(o.grade.descricao)}`); o.grade.descricao = trocaNome(o.grade.descricao); }
    if (o.desenhoFoto && o.desenhoFoto.skuLinha === 'CLM') { console.log(`OS ${o.os}: SKU do desenho gravado CLM -> CJM`); o.desenhoFoto.skuLinha = 'CJM'; }
  });

  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-sku-cjm-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    modelos: volta(data.modelos, modelos), desenhos: volta(data.desenhos, desenhos),
    grades: volta(data.grades, grades), ordens: volta(data.ordens, ordens),
    _device: 'sku-cjm-' + Date.now()
  });
  const r = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!r.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const conf = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  console.log('gravado: modelo ' + arr(conf.modelos).find(x => x.id === MODELO).skuLinha
    + ' · grades ' + arr(conf.grades).filter(g => /CALÇA MOLETOM/.test(g.tipoPeca)).map(g => g.nome).join(' / '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
