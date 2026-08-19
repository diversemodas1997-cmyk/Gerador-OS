/* Faz a LISTA dos PDFs de risco que estão na pasta de grades de corte.

   POR QUE ISTO EXISTE
   O app precisa mostrar, em cada grade cadastrada, um atalho para o PDF que
   mediu aquele comprimento e aquela largura. Os PDFs já estão no repositório —
   e, como o repositório é a própria pasta que o nginx serve, cada um deles já
   tem endereço: basta um link. O que faltava era o app SABER quais existem.

   Um navegador não lista pasta. Ele só pede arquivo por nome, e o nginx aqui
   está sem autoindex (nem deveria ter: seria expor a pasta inteira a quem
   digitar o endereço). Então a lista tem de estar escrita em algum lugar, e
   esse lugar é dados/riscos-pdf.json, gerado aqui e versionado junto.

   POR QUE NÃO GUARDAR ISSO NOS DADOS (no blob do Supabase)
   Porque não é dado da fábrica, é o que existe no disco — e o blob já custa
   1,8 MB por download (ver a memória de egress). Arquivo estático o navegador
   busca uma vez e guarda em cache; no blob, ele viajaria em toda gravação.

   QUANDO RODAR
   Toda vez que PDFs forem acrescentados, renomeados ou movidos na pasta
   "Desenhos técnicos -grades de corte". Rodar de novo é sempre seguro: o
   arquivo é reescrito do zero a partir do que está no disco.

     node servidor/indexar-riscos.js

   O caminho de cada PDF É a informação: "<LINHA>/<TAMANHOS>/<LARGURA> cm/
   <arquivo>.pdf" diz a que grade ele pertence. Por isso a lista guarda o
   caminho e nada mais — quem interpreta é o app (ver _riscosDaGrade no app.js).
*/
const fs = require('fs');
const path = require('path');

const RAIZ_REPO = path.resolve(__dirname, '..');
const PASTA = 'Desenhos técnicos -grades de corte';
const SAIDA = path.join(RAIZ_REPO, 'dados', 'riscos-pdf.json');

function varrer(dir, prefixo, achados) {
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    const rel = prefixo ? prefixo + '/' + ent.name : ent.name;
    if (ent.isDirectory()) varrer(path.join(dir, ent.name), rel, achados);
    else if (/\.pdf$/i.test(ent.name)) achados.push(rel);
  }
  return achados;
}

const raiz = path.join(RAIZ_REPO, PASTA);
if (!fs.existsSync(raiz)) {
  console.error(`não achei a pasta "${PASTA}" em ${RAIZ_REPO}`);
  process.exit(1);
}

// Ordem alfabética estável: sem isso o JSON muda de ordem entre máquinas e o
// git mostra diferença onde nada mudou.
const arquivos = varrer(raiz, '', []).sort((a, b) => a.localeCompare(b, 'pt-BR'));

// A data é do ARQUIVO, não do relógio de quem roda: serve para o app dizer "a
// lista é de tal dia" quando um PDF novo ainda não estiver nela.
const hoje = new Date().toISOString().slice(0, 10);
fs.mkdirSync(path.dirname(SAIDA), { recursive: true });
fs.writeFileSync(SAIDA, JSON.stringify({ pasta: PASTA, gerado: hoje, arquivos }, null, 0) + '\n', 'utf8');

const porLinha = {};
arquivos.forEach(a => { const l = a.split('/')[0]; porLinha[l] = (porLinha[l] || 0) + 1; });
console.log(`${arquivos.length} PDFs em dados/riscos-pdf.json`);
Object.keys(porLinha).sort().forEach(l => console.log(`  ${l}: ${porLinha[l]}`));
