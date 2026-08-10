/* Rode com:  node testes/parear-desenhos.js
   Pareia desenho tecnico com arquivo de imagem. O erro que importa aqui nao e
   falhar: e ACERTAR ERRADO — um desenho trocado numa OS vira corte errado e
   tecido perdido. Por isso a maior parte destes casos verifica que, na duvida,
   ele NAO pareia.

   Roda contra os dados reais quando eles estao a mao (a pasta de desenhos e o
   backup mais recente); senao, so os casos sinteticos. */
const fs = require('fs');
const path = require('path');
const { parear } = require('../servidor/parear-desenhos');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};

/* ------------------------- casos sinteticos ------------------------------- */
const PASTAS = {
  'CM.LISA': ['PRETO.png', 'BRANCO.png', 'VERMELHO.png', 'VERDE.png', 'AZUL MARINHO.png', 'GRAFITE.png'],
  'CM.REC':  ['PRETO.png', 'VERDE.png'],
  'CM.TRI':  ['TRICOLOR PRETA.png', 'TRICOLOR CAQUI.png'],
  'BM.TRI':  ['BEGE.png', 'VERDE.png']
};
const D = (codigo, skuLinha, desc) => ({ codigo, skuLinha, desc: desc || '' });

let r = parear([
  D('001', 'CM.LISA-PRE'),
  D('005', 'CM.LISA-MARINHO'),
  D('009', 'CM.LISA-GRAF'),
  D('004', 'CM.LISA-VERM'),
  D('003', 'CM.LISA-VERDE')
], PASTAS);
ok('1. abreviacao acha a cor pelo comeco da palavra',
   r.pares.get('001').arquivo === 'CM.LISA/PRETO.png'
   && r.pares.get('009').arquivo === 'CM.LISA/GRAFITE.png');
ok('1b. cor no MEIO do nome do arquivo tambem acha',
   r.pares.get('005').arquivo === 'CM.LISA/AZUL MARINHO.png');
ok('1c. VERM nao rouba VERDE (nem o contrario)',
   r.pares.get('004').arquivo === 'CM.LISA/VERMELHO.png'
   && r.pares.get('003').arquivo === 'CM.LISA/VERDE.png');

// A pasta e o MAIOR prefixo que casa: CM.REC.LISA-… nao pode cair em CM.
r = parear([D('0010', 'CM.REC.LISA-PRE')], Object.assign({ 'CM': ['PRETO.png'] }, PASTAS));
ok('2. prefixo mais longo vence (CM.REC.LISA -> CM.REC, nao CM)',
   r.pares.get('0010').arquivo === 'CM.REC/PRETO.png', JSON.stringify([...r.pares]));

// O caso que a descricao erraria: dois tricolores comecando por "Preto".
r = parear([
  D('0014', 'CM.TRI.LISA-CAQUI', 'Camiseta Tricolor | Preto/Caqui/Off-White'),
  D('0015', 'CM.TRI.LISA-PRE',   'Camiseta Tricolor | Preto/Grafite/Branco')
], PASTAS);
ok('3. SKU separa o que a descricao confundiria',
   r.pares.get('0014').arquivo === 'CM.TRI/TRICOLOR CAQUI.png'
   && r.pares.get('0015').arquivo === 'CM.TRI/TRICOLOR PRETA.png');
ok('3b. e nenhum arquivo foi usado duas vezes',
   new Set([...r.pares.values()].map(v => v.arquivo)).size === r.pares.size);

/* --------------- na duvida, NAO pareia (o que mais importa) --------------- */
r = parear([D('0013', '', 'Camiseta Bicolor | Verde/Branco')], PASTAS);
ok('4. sem SKU vira pendencia, nao palpite',
   r.pares.size === 0 && /sem SKU/.test(r.pendencias[0].motivo));

r = parear([D('X', 'CM.LISA-ROSA')], PASTAS);
ok('5. cor sem arquivo vira pendencia', r.pares.size === 0
   && /nenhum arquivo/.test(r.pendencias[0].motivo));

r = parear([D('Y', 'CM.LISA-VER')], PASTAS);   // VER casa VERMELHO e VERDE
ok('6. cor ambigua NAO pareia (2 candidatos)', r.pares.size === 0
   && /2 arquivos/.test(r.pendencias[0].motivo), JSON.stringify(r.pendencias));

r = parear([D('Z', 'XX.NADA-PRE')], PASTAS);
ok('7. prefixo sem pasta vira pendencia', r.pares.size === 0
   && /nenhuma pasta/.test(r.pendencias[0].motivo));

r = parear([D('W', 'SEMTRACO')], PASTAS);
ok('8. SKU fora do formato vira pendencia', r.pares.size === 0
   && /formato/.test(r.pendencias[0].motivo));

/* ------------------------------ mapa manual ------------------------------- */
r = parear([D('0013', '', 'Camiseta Bicolor | Verde/Branco')], PASTAS,
           { '0013': 'CM.REC/VERDE.png' });
ok('9. mapa manual resolve o que o SKU nao resolve',
   r.pares.get('0013').arquivo === 'CM.REC/VERDE.png'
   && r.pares.get('0013').origem === 'manual' && r.pendencias.length === 0);

r = parear([D('001', 'CM.LISA-PRE')], PASTAS, { '001': 'CM.LISA/BRANCO.png' });
ok('9b. mapa manual tem prioridade sobre o SKU',
   r.pares.get('001').arquivo === 'CM.LISA/BRANCO.png');

/* ------------------------- sobras sao reportadas -------------------------- */
r = parear([D('001', 'CM.LISA-PRE')], { 'CM.LISA': ['PRETO.png', 'ROXO.png'] });
ok('10. arquivo sem desenho aparece em sobrando',
   r.sobrando.length === 1 && r.sobrando[0] === 'CM.LISA/ROXO.png');

/* --------------------------- dados REAIS, se houver ----------------------- */
const pastaReal = path.join(__dirname, '..', 'Desenhos técnicos');
const backups = (() => {
  try {
    return fs.readdirSync(path.join(__dirname, '..', 'backups'))
      .filter(f => /^BACKUP-COMPLETO-.*\.json$/.test(f)).sort().pop();
  } catch (e) { return null; }
})();
if (fs.existsSync(pastaReal) && backups) {
  const porPasta = {};
  for (const p of fs.readdirSync(pastaReal)) {
    const dir = path.join(pastaReal, p);
    if (fs.statSync(dir).isDirectory()) {
      porPasta[p] = fs.readdirSync(dir).filter(f => /\.png$/i.test(f));
    }
  }
  const dados = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'backups', backups), 'utf8'));
  const reais = dados.desenhos || [];
  const rr = parear(reais, porPasta);
  console.log(`\n  (dados reais: ${reais.length} desenhos, ${Object.keys(porPasta).length} pastas)`);
  ok('11. real: pareia a grande maioria', rr.pares.size >= reais.length - 1,
     `${rr.pares.size}/${reais.length}`);
  ok('11b. real: NENHUM arquivo serve a dois desenhos',
     new Set([...rr.pares.values()].map(v => v.arquivo)).size === rr.pares.size);
  ok('11c. real: toda pendencia tem motivo legivel',
     rr.pendencias.every(p => p.motivo && p.motivo.length > 10),
     JSON.stringify(rr.pendencias));
  rr.pendencias.forEach(p => console.log(`       pendente: ${p.cod} — ${p.motivo}`));
} else {
  console.log('\n  (dados reais ausentes — so os casos sinteticos rodaram)');
}

console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
