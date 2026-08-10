/* Cifra e decifra o pacote de recuperação do servidor.

   O pacote carrega o JWT_SECRET, a SERVICE_ROLE_KEY, a senha do banco e os
   hashes de senha das contas. Isso vai parar numa pasta do Google Drive — ou
   seja, num lugar sincronizado, compartilhável por engano e fora do seu
   controle direto. Cifrado, um vazamento da pasta não entrega nada.

   AES-256-GCM: além de esconder, ele DETECTA adulteração. Um pacote truncado
   pela sincronização ou corrompido no disco falha ao abrir, em vez de restaurar
   um servidor pela metade — que seria pior do que não ter backup.

   A chave sai da senha por scrypt, que é lento de propósito: torna inviável
   tentar senhas em massa contra o arquivo. */
const crypto = require('crypto');

const MAGICO = 'GEROS-BKP1';       // identifica o formato e a versão
const SAL_BYTES = 16, IV_BYTES = 12, TAG_BYTES = 16;
const SCRYPT = { N: 1 << 15, r: 8, p: 1 };   // ~100ms por tentativa

function derivar(senha, sal) {
  return crypto.scryptSync(senha, sal, 32, Object.assign({ maxmem: 128 * 1024 * 1024 }, SCRYPT));
}

/** Objeto -> Buffer cifrado, pronto para gravar em arquivo. */
function cifrar(objeto, senha) {
  if (!senha || String(senha).length < 8) {
    throw new Error('A senha do backup precisa de pelo menos 8 caracteres.');
  }
  const sal = crypto.randomBytes(SAL_BYTES);
  const iv = crypto.randomBytes(IV_BYTES);
  const cifra = crypto.createCipheriv('aes-256-gcm', derivar(senha, sal), iv);
  const corpo = Buffer.concat([
    cifra.update(Buffer.from(JSON.stringify(objeto), 'utf8')),
    cifra.final()
  ]);
  return Buffer.concat([Buffer.from(MAGICO, 'utf8'), sal, iv, cifra.getAuthTag(), corpo]);
}

/** Buffer cifrado -> objeto. Lança se a senha estiver errada ou o arquivo tiver
 *  sido alterado/corrompido — os dois casos dão a mesma falha, de propósito. */
function decifrar(buffer, senha) {
  const cab = buffer.subarray(0, MAGICO.length).toString('utf8');
  if (cab !== MAGICO) {
    throw new Error('Este arquivo não é um pacote de recuperação do Gerador-OS.');
  }
  let p = MAGICO.length;
  const sal = buffer.subarray(p, p += SAL_BYTES);
  const iv = buffer.subarray(p, p += IV_BYTES);
  const tag = buffer.subarray(p, p += TAG_BYTES);
  const corpo = buffer.subarray(p);
  const decifra = crypto.createDecipheriv('aes-256-gcm', derivar(senha, sal), iv);
  decifra.setAuthTag(tag);
  let texto;
  try {
    texto = Buffer.concat([decifra.update(corpo), decifra.final()]).toString('utf8');
  } catch (e) {
    throw new Error('Senha errada, ou o arquivo está corrompido/incompleto.');
  }
  return JSON.parse(texto);
}

module.exports = { cifrar, decifrar, MAGICO };
