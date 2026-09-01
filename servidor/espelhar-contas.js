/* Espelha as contas de acesso do servidor da fábrica para a nuvem.

   POR QUE ISTO EXISTE

   As contas de nome (`admin`, `nathaly`, `enfesto.corte`…) só nascem no
   servidor da fábrica: quem as cria é a função `usuarios`, que só roda lá. A
   nuvem nunca soube delas — tinha só as três contas de e-mail antigas.

   O resultado, em 01/09/2026: o Admin digitava a senha CERTA e o programa
   respondia "Nome ou senha incorretos". A senha estava certa; o servidor é que
   era o outro. Basta a máquina abrir o programa pelo GitHub Pages, ou a
   sondagem do servidor local falhar, para o programa cair na nuvem — e ali
   nenhuma conta da fábrica existia. Sem conta na nuvem, a cópia de consulta
   (que existe justamente para o dia em que o servidor cair) não serve a
   ninguém: ninguém consegue entrar nela.

   O QUE ELE LEVA, E O QUE NÃO LEVA

   Leva o HASH da senha, nunca a senha. O bcrypt do `auth.users` é copiado como
   está, então as senhas continuam sendo exatamente as mesmas nos dois lugares —
   e nenhuma senha em claro passa por este script, pela rede ou pela tela. Se a
   senha mudar na fábrica, rodar de novo atualiza a da nuvem.

   Leva também o id da conta, o mesmo dos dois lados. É o que permite ao papel
   (admin/usuário) apontar para a mesma pessoa nos dois servidores.

   NÃO mexe nas contas de e-mail de verdade (@gmail): só cuida do que termina
   em @diverse.local, que é o que a fábrica cria.

   COMO RODAR (na máquina do servidor, com o Docker de pé)
     node servidor/espelhar-contas.js
     node servidor/espelhar-contas.js --so-listar    (mostra o que faria)

   Rode depois de criar uma conta nova, ou depois de trocar a senha de alguém.
   Não é destrutivo e pode ser repetido à vontade: conta que já está igual é
   deixada em paz. */

const fs = require('fs');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const SO_LISTAR = process.argv.includes('--so-listar');
const NUVEM = (arg('nuvem') || 'https://ckkqrjkhorvaahyazqsr.supabase.co').replace(/\/+$/, '');
const CHAVE_NUVEM_ARQUIVO = arg('nuvem-key-arquivo') || 'C:/supabase/nuvem-service-role.key';
const CONTAINER_DB = arg('db') || 'supabase-db';

function chaveDaNuvem() {
  // A primeira linha que não é comentário nem vazia — mesmo formato que o
  // espelhar-para-nuvem.js já usa.
  try {
    return fs.readFileSync(CHAVE_NUVEM_ARQUIVO, 'utf8').split(/\r?\n/)
      .map(l => l.trim()).filter(l => l && !l.startsWith('#'))[0] || '';
  } catch (e) { return ''; }
}

/* O hash da senha não sai pela API: `/auth/v1/admin/users` devolve tudo da
   conta MENOS `encrypted_password`. Então vem direto do banco. */
function contasDaFabrica() {
  const sql = "select u.id||'\t'||u.email||'\t'||u.encrypted_password||'\t'||coalesce(r.role,'usuario') "
    + 'from auth.users u left join public.user_roles r on r.user_id = u.id '
    + "where u.email like '%@diverse.local' and u.encrypted_password is not null "
    + 'order by u.created_at';
  const saida = execFileSync('docker',
    ['exec', CONTAINER_DB, 'psql', '-U', 'postgres', '-d', 'postgres', '-At', '-c', sql],
    { encoding: 'utf8' });
  return saida.split(/\r?\n/).filter(Boolean).map(l => {
    const [id, email, hash, papel] = l.split('\t');
    return { id, email, hash, papel };
  });
}

/* O QUE JÁ FOI COPIADO. Um arquivo com a impressão digital (sha256) do hash de
   cada conta na última cópia bem-sucedida.

   A primeira versão comparava o `updated_at` dos dois servidores, e não
   funcionava: são dois relógios diferentes, e o carimbo que o Postgres imprime
   não é o mesmo formato que o JavaScript lê — uma troca de senha na fábrica
   passava despercebida e a nuvem ficava com a senha velha, que é exatamente o
   problema que este script existe para não ter. A impressão digital responde
   "este hash já foi para a nuvem?" sem depender de relógio nenhum.

   O arquivo mora fora do repositório, junto das chaves: ele não é segredo (é
   um resumo de mão única de um hash), mas também não é assunto do git. */
const FICHA = arg('ficha') || 'C:/supabase/espelho-contas.json';

function fichaLer() {
  try { return JSON.parse(fs.readFileSync(FICHA, 'utf8')) || {}; } catch (e) { return {}; }
}
function fichaGravar(f) {
  try { fs.writeFileSync(FICHA, JSON.stringify(f, null, 2), 'utf8'); } catch (e) {
    console.error('Aviso: não consegui gravar ' + FICHA + ' — a próxima execução vai reenviar as senhas.');
  }
}
const digital = hash => crypto.createHash('sha256').update(String(hash)).digest('hex').slice(0, 16);

/* O trabalho em si. Devolve um resumo curto, para quem chama poder registrar
   uma linha só no log do espelho.

   O QUE DECIDE MEXER OU NÃO: conta que falta é criada; conta que já está lá só
   tem a senha reenviada quando o hash da fábrica não é o que consta na ficha.
   Sem isso, rodar de meia em meia hora reenviaria os cinco hashes para sempre,
   à toa. */
async function espelharContas(opcoes) {
  const o = opcoes || {};
  const nuvem = (o.nuvem || NUVEM).replace(/\/+$/, '');
  const chave = o.nuvemKey || chaveDaNuvem();
  const log = o.log || (m => console.log(m));
  const soListar = !!o.soListar;
  if (!chave) throw new Error(`Não achei a chave de serviço da nuvem em ${CHAVE_NUVEM_ARQUIVO}`);
  const H = { apikey: chave, Authorization: 'Bearer ' + chave, 'Content-Type': 'application/json' };

  const locais = contasDaFabrica();
  const r = await fetch(`${nuvem}/auth/v1/admin/users?page=1&per_page=200`, { headers: H });
  if (!r.ok) throw new Error(`não consegui ler as contas da nuvem (${r.status})`);
  const d = await r.json();
  const naNuvem = new Map((Array.isArray(d) ? d : (d.users || []))
    .map(u => [String(u.email).toLowerCase(), u]));

  const rel = { criadas: 0, senhas: 0, iguais: 0, falhas: 0 };
  const ficha = fichaLer();

  for (const c of locais) {
    const existente = naNuvem.get(c.email.toLowerCase());
    const marca = digital(c.hash);
    const desatualizada = existente && ficha[c.email] !== marca;

    if (existente && !desatualizada) { rel.iguais++; continue; }
    if (soListar) {
      log(`  ${existente ? '~ realinharia a senha de' : '+ criaria'} ${c.email} (${c.papel})`);
      existente ? rel.senhas++ : rel.criadas++;
      continue;
    }

    let idNaNuvem;
    if (existente) {
      /* A senha mudou na fábrica: a conta da nuvem é REFEITA com o hash novo.

         Não é preciosismo. O caminho óbvio — PUT em /admin/users/<id> com
         `password_hash` — responde 200 e NÃO troca a senha: o campo só é lido
         na criação. Testado em 01/09/2026: depois do PUT "bem-sucedido", a
         nuvem continuava aceitando a senha velha e recusando a nova. Um
         espelho que diz ter copiado e não copiou é pior do que não existir.

         Apagar e recriar com o MESMO id mantém tudo que aponta para a conta; o
         papel é reposto logo abaixo, porque ele cai junto com a linha antiga.
         De quebra, derruba as sessões abertas com a senha velha — que é o que
         se quer quando uma senha muda. */
      const del = await fetch(`${nuvem}/auth/v1/admin/users/${encodeURIComponent(existente.id)}`,
        { method: 'DELETE', headers: H });
      if (!del.ok) { log(`x ${c.email}: não consegui refazer a conta (${del.status})`); rel.falhas++; continue; }
      const cr = await fetch(`${nuvem}/auth/v1/admin/users`, {
        method: 'POST', headers: H,
        body: JSON.stringify({ id: existente.id, email: c.email, password_hash: c.hash, email_confirm: true })
      });
      if (!cr.ok) { log(`x ${c.email}: APAGADA da nuvem e não recriada (${cr.status}) — rode de novo`); rel.falhas++; continue; }
      idNaNuvem = (await cr.json()).id;
      ficha[c.email] = marca;
      rel.senhas++;
      log(`~ ${c.email}: senha realinhada com a da fábrica`);
    } else {
      // O mesmo id dos dois lados. Se a nuvem recusar o id (versão antiga do
      // GoTrue), cria com um id novo em vez de desistir da conta.
      let criado = null;
      for (const corpo of [{ id: c.id, email: c.email, password_hash: c.hash, email_confirm: true },
                           { email: c.email, password_hash: c.hash, email_confirm: true }]) {
        const cr = await fetch(`${nuvem}/auth/v1/admin/users`, {
          method: 'POST', headers: H, body: JSON.stringify(corpo)
        });
        if (cr.ok) { criado = await cr.json(); break; }
      }
      if (!criado) { log(`x ${c.email}: NÃO criado na nuvem`); rel.falhas++; continue; }
      idNaNuvem = criado.id;
      ficha[c.email] = marca;
      rel.criadas++;
      log(`+ ${c.email}: criado${criado.id === c.id ? ' (mesmo id da fábrica)' : ' (id novo)'}`);
    }

    if (c.papel === 'admin') {
      const pr = await fetch(`${nuvem}/rest/v1/user_roles`, {
        method: 'POST',
        headers: Object.assign({ Prefer: 'resolution=merge-duplicates' }, H),
        body: JSON.stringify({ user_id: idNaNuvem, role: 'admin' })
      });
      if (!pr.ok) log(`    papel admin: FALHOU (${pr.status})`);
    }
  }
  if (!soListar) fichaGravar(ficha);
  return rel;
}

module.exports = { espelharContas };

// Rodado direto na linha de comando (e não chamado pelo espelho de dados).
if (require.main === module) {
  espelharContas({ soListar: SO_LISTAR }).then(rel => {
    console.log(`Contas: ${rel.criadas} criada(s), ${rel.senhas} senha(s) realinhada(s), `
      + `${rel.iguais} já igual(is)${rel.falhas ? `, ${rel.falhas} FALHA(S)` : ''}.`);
    if (SO_LISTAR) console.log('Nada foi alterado (--so-listar).');
  }).catch(e => { console.error('Falhou:', e && e.message ? e.message : e); process.exit(1); });
}
