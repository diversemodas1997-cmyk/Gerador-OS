/*
 * Renomeia uma GRADE no cadastro do servidor.
 *
 * POR QUE ISTO EXISTE
 *
 * A auditoria de 20/08/2026 achou 4 grades (de 133) cujo NOME nao bate com as
 * QUANTIDADES cadastradas nelas - "G | CM.TRI" esta com G=1 e G1=2, ou seja, o
 * nome esconde um tamanho inteiro. Quem confere o risco le o nome; quem calcula
 * pano, pecas e volumes usa as quantidades. Enquanto os dois discordam, a
 * pessoa confere contra outra coisa e nao tem como perceber.
 *
 * Renomear pela tela daria na mesma, mas sao varias e o script deixa registro
 * do que foi trocado, confere antes e nao depende de ninguem estar logado.
 *
 * O QUE ELE CONFERE ANTES DE GRAVAR
 *
 *  - a grade existe, e uma so (nome exato);
 *  - o nome novo nao esta em uso por outra grade - dois nomes iguais no
 *    cadastro deixam a escolha na OS ambigua;
 *  - diz se o nome novo passa a bater com as quantidades (que e o motivo de
 *    renomear) e AVISA se continuar divergindo, sem impedir: o certo pode ser
 *    corrigir a quantidade, e nao o nome, e quem decide e quem cadastra.
 *
 * O QUE ELE NAO MEXE
 *
 * As OS ja emitidas guardam o nome da grade como TEXTO, um retrato do dia em
 * que foram criadas (o.grade.descricao). Isso fica como esta, de proposito: a
 * folha ja mostra o nome VIVO da grade (renderPrintSheet reaponta pelo gradeId)
 * e o agrupamento por grade tambem e por id (_osGradeKey). Reescrever o retrato
 * seria mudar o que aquele papel dizia no dia.
 *
 * COMO RODAR
 *
 *   node servidor/renomear-grade.js "G | CM.TRI | 116.5cm" "G-2G1 | CM.TRI | 116.5cm"
 *   ... e de novo com --gravar no fim, para valer.
 */
const fs = require('fs');
const path = require('path');

const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';

const args = process.argv.slice(2);
const GRAVAR = args.includes('--gravar');
const [DE, PARA] = args.filter(a => !a.startsWith('--'));

if (!DE || !PARA) {
  console.error('Uso: node servidor/renomear-grade.js "<nome atual>" "<nome novo>" [--gravar]');
  process.exit(1);
}

// O nome que as QUANTIDADES da grade pedem — recortado do app.js para nao
// existir uma segunda versao da convencao de nomes aqui.
function nomePelosTamanhos(tamanhos) {
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const i = src.indexOf('function _riscoNomeTamanhos');
  const fn = src.slice(i, src.indexOf('\n}', i) + 2);
  return new Function(fn + 'return _riscoNomeTamanhos;')()(tamanhos || {});
}

(async () => {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  const cab = { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
  const linhas = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!linhas || !linhas[0]) throw new Error('nao achei a linha main');
  const data = linhas[0].data, updatedAt = linhas[0].updated_at;
  const grades = JSON.parse(data.grades || '[]');

  const alvos = grades.filter(g => String(g.nome || '').trim() === DE.trim());
  if (!alvos.length) { console.error(`Nao achei grade chamada "${DE}".`); process.exit(1); }
  if (alvos.length > 1) { console.error(`Ha ${alvos.length} grades com esse nome — resolva a duplicata antes.`); process.exit(1); }
  const g = alvos[0];
  const jaUsado = grades.find(x => x !== g && String(x.nome || '').trim() === PARA.trim());
  if (jaUsado) { console.error(`Ja existe outra grade chamada "${PARA}".`); process.exit(1); }

  const qtd = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3']
    .filter(k => (parseInt((g.tamanhos || {})[k], 10) || 0) > 0)
    .map(k => k.toUpperCase() + '=' + g.tamanhos[k]).join(' ');
  const pedido = nomePelosTamanhos(g.tamanhos);
  const partePara = (PARA.split('|')[0] || '').trim();

  console.log(`de   : ${g.nome}`);
  console.log(`para : ${PARA}`);
  console.log(`quantidades cadastradas: ${qtd}   (a convencao pede "${pedido}")`);
  const ordens = JSON.parse(data.ordens || '[]');
  const usam = ordens.filter(o => o.gradeId === g.id);
  console.log(`OS que usam esta grade: ${usam.length}${usam.length ? ' — ' + usam.map(o => o.os).join(', ') : ''}`);
  console.log('  (o nome guardado nelas e um retrato do dia da emissao e NAO e mexido;');
  console.log('   a folha ja mostra o nome vivo, pelo gradeId)');
  if (partePara !== pedido) {
    console.log('');
    console.log(`AVISO: o nome novo ("${partePara}") continua diferente do que as quantidades`);
    console.log(`       pedem ("${pedido}"). Talvez o que precise mudar seja a QUANTIDADE.`);
  }

  if (!GRAVAR) { console.log('\n(nada foi gravado — repita com --gravar para valer)'); return; }

  g.nome = PARA.trim();
  data.grades = JSON.stringify(grades);
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  }
  console.log('\nGRAVADO no servidor.');
})().catch(e => { console.error('ERRO: ' + (e && e.message || e)); process.exit(1); });
