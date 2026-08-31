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

   O --ip aceita MAIS DE UM endereco, separados por virgula. Serve quando o
   servidor e alcancavel por dois caminhos — foi o caso em 28/08/2026, com o
   cabo de rede fora e a fabrica dependendo do Wi-Fi:
     node servidor\gerar-certificado.js --ip 193.168.0.200,192.168.1.158
   Reemitir assim NAO toca na autoridade: quem ja instalou o ca.crt continua
   valendo, sem passar de maquina em maquina.

   O --nome tambem aceita varios, separados por virgula, e vale MAIS que o IP:
   um numero pertence a UMA rede, um nome nao. O Windows resolve o nome da
   maquina sozinho na rede em que o cliente estiver (LLMNR/NetBIOS), entao
   https://DESKTOP-SOV61AF abre pelo cabo quando os dois estao no cabo, e pelo
   Wi-Fi quando os dois estao no Wi-Fi -- sem ninguem reescrever atalho nenhum.
     node servidor\gerar-certificado.js --ip 193.168.0.200,192.168.1.158 --nome DESKTOP-SOV61AF
   Cada nome entra duas vezes no SAN: "nome" e "nome.local".
*/
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
// Aceita um ou varios: "--ip a,b". O primeiro e o principal (vai no CN).
const IPS = String(arg('ip') || '').split(',').map(x => x.trim()).filter(Boolean);
const IP = IPS[0];
// Aceita um ou varios, igual ao --ip. O nome importa mais que o numero: um
// NOME resolve sozinho na rede em que o cliente estiver (LLMNR/NetBIOS), e por
// isso sobrevive a troca de cabo para Wi-Fi sem ninguem reescrever atalho.
const NOMES = String(arg('nome') || '').split(',').map(x => x.trim()).filter(Boolean);
const NOME = NOMES[0];
const SAIDA = arg('saida') || path.join(__dirname, 'tls');
const DIAS_CA = 3650;    // a CA vale 10 anos: trocá-la obriga a passar em todas as máquinas
const DIAS_SRV = 825;    // o certificado do servidor, ~2 anos (limite aceito pelos navegadores)

const invalido = IPS.filter(x => !/^\d{1,3}(\.\d{1,3}){3}$/.test(x));
if (!IPS.length || invalido.length) {
  console.error(invalido.length
    ? 'Nao e um endereco IPv4: ' + invalido.join(', ')
    : 'Faltou --ip. Ex.: --ip 192.168.0.50   (ou --ip 192.168.0.50,192.168.1.158)');
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
const nomes = [...IPS.map(x => `IP:${x}`), 'IP:127.0.0.1', 'DNS:localhost'];
for (const n of NOMES) nomes.push(`DNS:${n}`, `DNS:${n}.local`);

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

// A CA é REAPROVEITADA quando já existe, e isso é o ponto mais importante deste
// arquivo. Ela é o que cada computador da fábrica instalou uma vez; refazê-la
// invalida silenciosamente todos eles de uma vez, e o sintoma aparece só depois,
// como "deu erro de segurança em todas as máquinas". Reemitir o certificado do
// servidor — porque o IP mudou, ou porque venceu — não pode custar isso.
// Para trocar a CA de propósito, use --refazer-ca, sabendo que aí é obrigatório
// reinstalar o ca.crt em cada computador.
const REFAZER_CA = process.argv.includes('--refazer-ca');
const temCA = fs.existsSync(p('ca.key')) && fs.existsSync(p('ca.crt'));

if (temCA && !REFAZER_CA) {
  const fim = (ssl(['x509', '-in', p('ca.crt'), '-noout', '-enddate']).split('=')[1] || '').trim();
  console.log('Aproveitando a autoridade certificadora que já existe (vale até ' + fim + ').');
  console.log('  As máquinas que já têm o ca.crt instalado continuam valendo.');
} else {
  if (temCA) console.log('REFAZENDO a autoridade certificadora — será preciso reinstalar o ca.crt em CADA computador.');
  else console.log('Gerando a autoridade certificadora da fábrica…');
  ssl(['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-sha256',
    '-days', String(DIAS_CA), '-keyout', p('ca.key'), '-out', p('ca.crt'),
    '-subj', '/CN=Gerador-OS - CA da fabrica/O=Gerador-OS',
    '-config', cfg, '-extensions', 'ca']);
}

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
const faltando = IPS.filter(x => !san.includes(`IP Address:${x}`));
if (faltando.length) {
  console.error('O certificado saiu SEM estes enderecos na lista de nomes: '
    + faltando.join(', ') + ' — o navegador recusaria.');
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
