/* Rode com:  node testes/cofre.js
   O pacote de recuperação é a última linha de defesa: se ele não abrir no dia
   em que o servidor morrer, não há segunda chance. Estes testes cobrem tanto o
   caminho feliz quanto as formas de o arquivo chegar quebrado. */
const { cifrar, decifrar } = require('../servidor/cofre');

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   ' + (extra || '')));
  if (!cond) falhas++;
};
const lanca = fn => { try { fn(); return null; } catch (e) { return e.message; } };

const SENHA = 'senha-do-backup-2026';
const PACOTE = {
  gerado_em: '2026-08-10T12:00:00.000Z',
  env: { JWT_SECRET: 'abc123', SERVICE_ROLE_KEY: 'eyJ...', POSTGRES_PASSWORD: 's3nh4' },
  banco: 'CREATE TABLE shared_data (...);\n-- acentuação: ç ã õ é\n',
  desenhos: [{ nome: 'desenhos/1777290341246_1bsz6o.png', b64: 'iVBORw0KGgo=' }]
};

// 1. Ida e volta preserva tudo, inclusive acentos e binário em base64.
const cifrado = cifrar(PACOTE, SENHA);
const volta = decifrar(cifrado, SENHA);
ok('1. ida e volta preserva o pacote', JSON.stringify(volta) === JSON.stringify(PACOTE));
ok('1b. acentuação intacta', volta.banco.includes('ç ã õ é'));
ok('1c. imagem intacta', volta.desenhos[0].b64 === 'iVBORw0KGgo=');

// 2. O conteúdo não pode aparecer em claro dentro do arquivo.
const cru = cifrado.toString('latin1');
ok('2. segredos não aparecem em claro', !cru.includes('s3nh4') && !cru.includes('JWT_SECRET'));

// 3. Senha errada não abre.
ok('3. senha errada é recusada', /Senha errada/.test(lanca(() => decifrar(cifrado, 'outra-senha'))));

// 4. Arquivo adulterado no meio é recusado (é o GCM detectando).
const mexido = Buffer.from(cifrado);
mexido[mexido.length - 5] ^= 0xff;
ok('4. adulteração é detectada', /corrompido|Senha errada/.test(lanca(() => decifrar(mexido, SENHA))));

// 5. Arquivo truncado (sincronização interrompida) é recusado.
ok('5. arquivo truncado é recusado',
   lanca(() => decifrar(cifrado.subarray(0, cifrado.length - 40), SENHA)) !== null);

// 6. Outro arquivo qualquer dá mensagem clara, não erro de criptografia.
ok('6. arquivo estranho dá mensagem clara',
   /não é um pacote de recuperação/.test(lanca(() => decifrar(Buffer.from('qualquer coisa aqui dentro'), SENHA))));

// 7. Senha fraca é barrada na hora de criar, não na hora de restaurar.
ok('7. senha curta é recusada ao criar', /pelo menos 8/.test(lanca(() => cifrar(PACOTE, 'abc'))));

// 8. Dois pacotes iguais geram arquivos diferentes (sal e IV aleatórios) —
//    senão daria para comparar backups e deduzir que nada mudou.
ok('8. cifragens repetidas não são idênticas',
   !cifrar(PACOTE, SENHA).equals(cifrar(PACOTE, SENHA)));

// 9. Pacote grande (imagens de verdade) continua funcionando.
const grande = { desenhos: Array.from({ length: 30 }, (_, i) =>
  ({ nome: 'img' + i, b64: Buffer.alloc(200 * 1024, i % 251).toString('base64') })) };
const g = decifrar(cifrar(grande, SENHA), SENHA);
ok('9. pacote de ~8 MB sobrevive', g.desenhos.length === 30
   && g.desenhos[29].b64 === grande.desenhos[29].b64);

console.log(falhas ? '\n>>> ' + falhas + ' FALHA(S)' : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
