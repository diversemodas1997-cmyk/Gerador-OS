/* Rode com:  node testes/riscos-da-grade.js

   O atalho da grade para o PDF que mediu o comprimento e a largura dela.

   A coluna "Riscos" do cadastro de grades acha os PDFs pelo CAMINHO em que eles
   estão guardados, que na pasta da casa é uma classificação completa:

       BM.LISA / 2M-4G-2GG / 177 cm - MAPAS IMPRESSOS / BM.LISA - CORPO ....pdf
       linha     tamanhos    largura                    arquivo

   e o nome da grade diz as mesmas três coisas: "2M-4G-2GG | BM.LISA | 177cm".

   Em 19/08/2026 as pastas e os arquivos perderam o "x" dos tamanhos ("2xM" virou
   "2M"), para ficarem escritos como o cadastro escreve. O casamento não dependia
   disso nem antes: _chaveTam já tirava o x dos dois lados — e é por isso que a
   renomeação de 120 nomes não mexeu em uma linha de regra.

   O QUE ESTE TESTE PROTEGE, e por que cada regra existe:

   1. A LINHA É OBRIGATÓRIA. Na primeira versão da regra, quando a linha da
      grade não tinha pasta, a busca caía para "qualquer linha" — e a
      "P ao G3 | CM.TRI" oferecia os PDFs da CM.LISA como se fossem dela. Abrir
      o PDF errado é PIOR do que não ter atalho: quem confere uma medida no
      arquivo de outro produto confirma um erro em vez de achá-lo.

   2. AS DUAS GRAFIAS DOS TAMANHOS. A pasta escreve "P-M-G-GG-G1-G2-G3" e a
      grade se chama "P ao G3". As duas vêm de _riscoFormasDoNome, a mesma que o
      assistente de pasta usa no sentido contrário.

   3. LARGURA É PREFERÊNCIA, NÃO FILTRO. Havendo pasta na largura exata, são só
      aquelas — a "P ao G3 | CM.LISA | 117cm" não pode trazer junto a pasta de
      116,5 cm, que é outro encaixe. Não havendo, mostra as outras larguras COM
      AVISO: a "2G3 | BM.LISA | 182cm" tem pasta só em "183 cm", e é ela mesma.

   4. O ✓ DO ARQUIVO REGISTRADO casa por SUFIXO, porque o caminho guardado
      depende de onde a pasta foi escolhida na hora de importar.

   O teste recorta as funções do app.js de verdade e roda contra o índice de
   verdade (dados/riscos-pdf.json). Copiar a regra para cá testaria a cópia. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(deOnde, ate, oQue) {
  const i = src.indexOf(deOnde);
  const j = i < 0 ? -1 : src.indexOf(ate, i + deOnde.length);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);
// Para as constantes que ocupam mais de uma linha: recortar até a primeira
// quebra deixaria a expressão pela metade, e o erro sairia como "Unexpected
// token )" dentro de um `new Function` — vinte minutos para achar o óbvio.
const cortaAte = (nome, fim) => recorte(nome, fim, nome) + fim;

const motor = [
  corta('function _normNome'),
  cortaLinha('const _chaveTam'),
  corta('function _riscoNomeTamanhos'),
  corta('function _riscoFormasDoNome'),
  cortaAte('const _riscoCmDoTexto = s =>', '\n};'),
  cortaAte('const _riscoUrl = rel =>', ".join('/');"),
  cortaLinha('const _riscoCaminhoDoPdf = L =>'),
  corta('function _gradeNomePartes'),
  corta('function _riscosDaGrade'),
  corta('function _riscosRegistrados')
].join('\n');

const indice = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'dados', 'riscos-pdf.json'), 'utf8'));

// Monta o _riscosIdx do mesmo jeito que _riscosIndice monta no app.
const itens = indice.arquivos.map(rel => {
  const p = rel.split('/');
  return {
    rel,
    linha: p[0] || '',
    tam: p.length > 2 ? p[1] : '',
    cm: p.length > 3 ? (String(p[2]).match(/(\d+[.,]?\d*)\s*cm/i) || ['', ''])[1].replace(',', '.') : '',
    pasta: p.slice(0, -1).join('/'),
    arq: p[p.length - 1]
  };
});

function rodar(codigo, grade) {
  const fn = new Function('IDX', 'GRADE', `
    let _riscosIdx = IDX;
    ${motor}
    return (${codigo});
  `);
  return fn({ pasta: indice.pasta, gerado: indice.gerado, itens }, grade);
}
const achar = (grade) => rodar('_riscosDaGrade(GRADE)', grade);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------- as grades são as REAIS, do cadastro de 18/08/2026 ---------- */
const G_BM_177 = { id: 'a', nome: '2M-4G-2GG | BM.LISA | 177cm', tamanhos: { m: 2, g: 4, gg: 2 }, fases: [] };
const G_CMTRI  = { id: 'b', nome: 'P ao G3 | CM.TRI | 116.5cm', tamanhos: { p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }, fases: [] };
const G_CMLISA = { id: 'c', nome: 'P ao G3 | CM.LISA | 117cm', tamanhos: { p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }, fases: [] };
const G_2G3    = { id: 'd', nome: '2G3 | BM.LISA | 182cm', tamanhos: { g3: 2 }, fases: [] };

/* ---------- 1. acha os PDFs da própria pasta ---------- */
const r1 = achar(G_BM_177);
ok('2M-4G-2GG | BM.LISA | 177cm acha PDF', r1.itens.length > 0, r1.itens.length);
ok('  ... e todos são da pasta de 177 cm da BM.LISA',
  r1.itens.length > 0 && r1.itens.every(p => p.linha === 'BM.LISA' && p.cm === '177'),
  r1.itens.map(p => p.rel));
ok('  ... sem aviso de largura', r1.aviso === '', r1.aviso);

/* ---------- 2. a linha é obrigatória (o defeito que motivou o teste) ----------
   Até 19/08/2026 esta grade não tinha pasta nenhuma, e o teste cobrava LISTA
   VAZIA — que era só o sintoma de "não emprestou a da CM.LISA". Em 20/08 a
   pasta CM.TRI/P-M-G-GG-G1-G2-G3 passou a existir e a grade achou os cinco
   riscos DELA, corretamente; a asserção antiga acusou o acerto como falha.
   O que este teste protege é a LINHA, então é a linha que ele confere: venha
   um PDF ou nenhum, nenhum deles pode ser de outra linha. */
const r2 = achar(G_CMTRI);
ok('P ao G3 | CM.TRI NÃO empresta a pasta da CM.LISA',
  r2.itens.every(p => p.linha === 'CM.TRI'),
  r2.itens.map(p => p.rel));

/* ---------- 3. faixa ("P ao G3") casa com pasta extensa ---------- */
const r3 = achar(G_CMLISA);
ok('P ao G3 casa com a pasta P-M-G-GG-G1-G2-G3',
  r3.itens.length > 0 && r3.itens.every(p => p.linha === 'CM.LISA'), r3.itens.map(p => p.rel));
ok('  ... só a largura do nome (117), não a de 116.5 ao lado',
  r3.itens.length > 0 && r3.itens.every(p => p.cm === '117' || p.cm === ''),
  r3.itens.map(p => p.pasta));

/* ---------- 4. largura é preferência, não filtro ---------- */
const r4 = achar(G_2G3);
ok('2G3 | BM.LISA | 182cm acha a pasta de 183 cm', r4.itens.length > 0, r4.itens.length);
ok('  ... e avisa que a largura é outra', /outra largura/i.test(r4.aviso), r4.aviso);

/* ---------- 5. grade sem linha no nome não chuta ---------- */
const r5 = achar({ id: 'e', nome: 'P ao G3', tamanhos: { p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }, fases: [] });
ok('nome sem linha não devolve nada', r5.itens.length === 0, r5.itens.length);

/* ---------- 6. o endereço do PDF ---------- */
const url = rodar('_riscoUrl("BM.LISA/2G/182 cm/BM.LISA - CORPO 2G.pdf")', null);
ok('a URL escapa espaço e acento e mantém as barras',
  url === 'Desenhos%20t%C3%A9cnicos%20-grades%20de%20corte/BM.LISA/2G/182%20cm/BM.LISA%20-%20CORPO%202G.pdf', url);

/* ---------- 7. o ✓ do arquivo registrado, por sufixo ---------- */
const gReg = {
  id: 'f', nome: '2M-4G-2GG | BM.LISA | 177cm', tamanhos: { m: 2, g: 4, gg: 2 },
  // Caminho como o seletor nativo entrega: começa DENTRO da pasta escolhida.
  fases: [{ risco: 'BM.LISA/2M-4G-2GG/177 cm - MAPAS IMPRESSOS/BM.LISA - CORPO 2M-4G-2GG.pdf' }]
};
const usado = rodar('_riscosRegistrados(GRADE)', gReg);
ok('marca o PDF que a grade registrou',
  usado('BM.LISA/2M-4G-2GG/177 cm - MAPAS IMPRESSOS/BM.LISA - CORPO 2M-4G-2GG.pdf') === true);
ok('  ... mesmo com a pasta raiz na frente (input webkitdirectory)',
  rodar('_riscosRegistrados(GRADE)', {
    ...gReg,
    fases: [{ risco: 'Desenhos técnicos -grades de corte/BM.LISA/2M-4G-2GG/177 cm - MAPAS IMPRESSOS/BM.LISA - CORPO 2M-4G-2GG.pdf' }]
  })('BM.LISA/2M-4G-2GG/177 cm - MAPAS IMPRESSOS/BM.LISA - CORPO 2M-4G-2GG.pdf') === true);
ok('  ... e não marca um arquivo de outra fase',
  usado('BM.LISA/2M-4G-2GG/177 cm - MAPAS IMPRESSOS/BM.LISA - RIBANA 2M-4G-2GG.pdf') === false);

/* ---------- 8. o índice e a pasta continuam de pé ---------- */
ok('o índice tem PDFs', itens.length > 100, itens.length);
ok('todo item do índice tem linha e arquivo',
  itens.every(p => p.linha && p.arq), itens.filter(p => !p.linha || !p.arq).slice(0, 3));

console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
