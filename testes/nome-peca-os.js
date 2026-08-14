/* Rode com:  node testes/nome-peca-os.js

   COMO A FOLHA CHAMA A PEÇA.

   O modelo do cadastro é a CATEGORIA ("Camiseta Bicolor" = duas cores) e é dele
   que dependem a regra da conjugada, o filtro de tecidos e a grade vinculada. O
   NOME da peça mora na descrição do desenho ("Camiseta Recortada | Preto/Branco").

   Até 14/08/2026 a folha de OS e a de OE imprimiam o modelo, e a OS 0468 chegou
   na doca anunciada como "Camiseta Bicolor" — uma peça que a fábrica inteira
   chama de recortada, com a grade dizendo CM.REC e o SKU dizendo CM.REC.LISA-PRE.
   Ninguém lia o único campo que trazia o nome certo.

   O teste existe porque o erro não dá erro: dá papel com o nome errado, e quem
   está no chão acredita no papel. E porque o fallback importa tanto quanto a
   regra — se ele falhar, OS antiga sai SEM NOME nenhum, que é pior que sair com
   o nome da categoria.

   Recorta nomePecaOS do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const i = src.indexOf('function nomePecaOS');
const j = src.indexOf('\n}', i);
if (i < 0 || j < 0) { console.error('nao achei nomePecaOS no app.js'); process.exit(1); }
const motor = src.slice(i, j) + '\n}';

const nome = (STATE, os) => new Function('STATE', 'os', `${motor}\nreturn nomePecaOS(os);`)(STATE, os);

let falhas = 0;
const ok = (t, got, esperado) => {
  const bate = got === esperado;
  console.log((bate ? '  ok  ' : 'FALHA ') + t + (bate ? '' : `   esperado: ${JSON.stringify(esperado)}  obtido: ${JSON.stringify(got)}`));
  if (!bate) falhas++;
};

// O caso real: desenho 0010, cadastrado sob o modelo "Camiseta Bicolor".
const STATE = {
  desenhos: [
    { id: 'd_rec', codigo: '0010', desc: 'Camiseta Recortada | Preto/Branco' },
    { id: 'd_semdesc', codigo: '0099', desc: '' },
    { id: 'd_semtubo', codigo: '0098', desc: 'Camiseta Polo' }
  ],
  modelos: [{ id: 'm_bi', nome: 'Camiseta Bicolor' }]
};
const os = (extra) => Object.assign({ id: 'os_1', os: '0468', modeloNome: 'Camiseta Bicolor' }, extra);

ok('o nome vem da descrição do desenho, não do modelo',
  nome(STATE, os({ desenhoId: 'd_rec' })), 'Camiseta Recortada');

ok('a cor depois do "|" não entra no nome',
  nome(STATE, os({ desenhoId: 'd_rec', modeloNome: '' })), 'Camiseta Recortada');

ok('descrição sem "|" vale inteira',
  nome(STATE, os({ desenhoId: 'd_semtubo' })), 'Camiseta Polo');

// Fallbacks: o que não pode é a folha sair em branco.
ok('desenho sem descrição: cai no modelo gravado na OS',
  nome(STATE, os({ desenhoId: 'd_semdesc' })), 'Camiseta Bicolor');

ok('desenho apagado do cadastro: cai no modelo gravado na OS',
  nome(STATE, os({ desenhoId: 'd_sumiu' })), 'Camiseta Bicolor');

ok('OS antiga, sem desenhoId: cai no modelo gravado na OS',
  nome(STATE, os({})), 'Camiseta Bicolor');

ok('sem desenho e sem modelo: string vazia, e a folha põe o travessão',
  nome(STATE, os({ modeloNome: '' })), '');

ok('sem OS nenhuma: string vazia, sem estourar',
  nome(STATE, null), '');

ok('cadastro de desenhos ainda não carregado: cai no modelo',
  nome({ modelos: [] }, os({ desenhoId: 'd_rec' })), 'Camiseta Bicolor');

// Espaços em volta do "|" são hábito de digitação, não parte do nome.
ok('espaços em volta do separador são aparados',
  nome({ desenhos: [{ id: 'd_x', desc: '  Blusa Moletom Tricolor   |  Verde/Preto/Bege ' }] },
    os({ desenhoId: 'd_x' })), 'Blusa Moletom Tricolor');

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
