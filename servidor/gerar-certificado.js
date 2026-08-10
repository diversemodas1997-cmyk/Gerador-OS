/* Gera o certificado HTTPS do servidor da fábrica.

   POR QUE PRECISA DE HTTPS
   O navegador só libera a escolha de pastas (a gravação automática do PDF da
   OS, das etiquetas, do backup e da OE) em "contexto seguro": https, ou
   localhost. Um endereço como http://192.168.0.50 não é. Sem HTTPS, essas
   quatro funções ficam presas à máquina do servidor.

   COMO FUNCIONA
   Cria uma autoridade certificadora SUA (uma "CA da fábrica") e, assinado por
   ela, o certificado do servidor. A CA é instalada uma vez em cada computador;
   a partir daí eles confiam no servidor e o cadeado aparece normalmente.
   Não envolve internet, domínio nem pagar nada.

   Uso:
     node servidor\gerar-certificado.js --ip 192.168.0.50 [--nome gerador-os]

   O --nome é opcional: serve se você der um nome ao servidor no roteador, para
   poder acessar por https://gerador-os em vez do IP.
*/
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const IP = arg('ip');
const NOME = arg('nome');
const SAIDA = arg('saida') || path.join(__dirname, 'tls');
const DIAS_CA = 3650;    // a CA vale 10 anos: trocá-la obriga a passar em todas as máquinas
const DIAS_SRV = 825;    // o certificado do servidor, ~2 anos (limite aceito pelos navegadores)

if (!IP || !/^\d{1,3}(\.\d{1,3}){3}$/.test(IP)) {
  console.error('Faltou --ip, ou o valor não é um endereço IPv4. Ex.: --ip 192.168.0.50');
  process.exit(1);
}

// O openssl vem junto com o Git para Windows; no Linux costuma estar no PATH.
function acharOpenssl() {
  const candidatos = [
    'openssl',
    'C:\\Program Files\\Git\\usr\\bin\\openssl.exe',
    'C:\\Program Files\\Git\\mingw64\\bin\\openssl.exe',
    'C:\\Program Files (x86)\\Git\\usr\\bin\\openssl.exe'
  ];
  for (const c of candidatos) {
    try { execFileSync(c, ['version'], { stdio: 'ignore' }); return c; } catch (e) { /* tenta o próximo */ }
  }
  console.error('Não achei o openssl. Ele vem com o Git para Windows —\n'
    + 'instale o Git (https://git-scm.com/download/win) e rode este comando de novo.');
  process.exit(1);
}
const OPENSSL = acharOpenssl();
const ssl = args => execFileSync(OPENSSL, args, { stdio: ['ignore', 'pipe', 'pipe'] }).toString();

fs.mkdirSync(SAIDA, { recursive: true });
const p = n => path.join(SAIDA, n);

// Os nomes pelos quais o servidor pode ser chamado. O navegador exige que o
// endereço digitado esteja NESTA lista — um certificado sem o IP aqui dá erro
// mesmo estando tudo certo no resto.
const nomes = [`IP:${IP}`, 'IP:127.0.0.1', 'DNS:localhost'];
if (NOME) nomes.push(`DNS:${NOME}`, `DNS:${NOME}.local`);

const cfg = p('openssl.cnf');
fs.writeFileSync(cfg, `
[req]
distinguished_name = dn
[dn]
[servidor]
basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = ${nomes.join(', ')}
[ca]
basicConstraints = critical, CA:TRUE, pathlen:0
keyUsage = critical, keyCertSign, cRLSign
`.trim() + '\n');

console.log('Gerando a autoridade certificadora da fábrica…');
ssl(['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-sha256',
  '-days', String(DIAS_CA), '-keyout', p('ca.key'), '-out', p('ca.crt'),
  '-subj', '/CN=Gerador-OS - CA da fabrica/O=Gerador-OS',
  '-config', cfg, '-extensions', 'ca']);

console.log('Gerando o certificado do servidor…');
ssl(['req', '-newkey', 'rsa:2048', '-nodes', '-sha256',
  '-keyout', p('servidor.key'), '-out', p('servidor.csr'),
  '-subj', `/CN=${NOME || IP}/O=Gerador-OS`, '-config', cfg]);
ssl(['x509', '-req', '-in', p('servidor.csr'), '-CA', p('ca.crt'), '-CAkey', p('ca.key'),
  '-CAcreateserial', '-out', p('servidor.crt'), '-days', String(DIAS_SRV), '-sha256',
  '-extfile', cfg, '-extensions', 'servidor']);

// Confere o que saiu, em vez de supor. Um certificado sem o IP na lista de
// nomes, ou que não valide contra a própria CA, dá erro de segurança no
// navegador e parece problema de configuração do servidor.
console.log('Conferindo…');
const verif = ssl(['verify', '-CAfile', p('ca.crt'), p('servidor.crt')]);
if (!/OK\s*$/m.test(verif)) { console.error('O certificado não validou contra a CA:\n' + verif); process.exit(1); }
const texto = ssl(['x509', '-in', p('servidor.crt'), '-noout', '-text']);
const san = (texto.match(/X509v3 Subject Alternative Name:\s*\n\s*(.+)/) || [])[1] || '';
if (!san.includes(`IP Address:${IP}`)) {
  console.error('O certificado saiu SEM o IP na lista de nomes — o navegador recusaria.');
  process.exit(1);
}
try { fs.unlinkSync(p('servidor.csr')); } catch (e) {}

console.log(`
=====================================================================
 Certificado pronto em ${SAIDA}
=====================================================================
 Nomes aceitos: ${san.trim()}
 Servidor vale ate: ${(ssl(['x509','-in',p('servidor.crt'),'-noout','-enddate']).split('=')[1]||'').trim()}

 ARQUIVOS
   ca.crt        -> instalar em CADA computador (e o que cria a confianca)
   ca.key        -> SEGREDO. So no servidor. Quem tem isto emite certificado
                    em nome da sua CA — nao coloque em pasta compartilhada.
   servidor.crt  -> usado pelo nginx
   servidor.key  -> usado pelo nginx, tambem segredo

 EM CADA COMPUTADOR (PowerShell como administrador), com o ca.crt copiado:

   Import-Certificate -FilePath .\\ca.crt -CertStoreLocation Cert:\\LocalMachine\\Root

 Depois feche e reabra o navegador. Chrome e Edge usam a lista do Windows.
 (O Firefox tem lista propria — se alguem usar Firefox, importe por
 Configuracoes -> Certificados -> Ver certificados -> Autoridades.)
=====================================================================
`);
