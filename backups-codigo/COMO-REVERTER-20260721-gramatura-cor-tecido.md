# Como reverter — "cores com tecido no nome + total de bobinas por tecido"

Alteração de **21/07/2026**. Commit: `ea8d984`. Estado anterior: `8b5646c`.

Arquivos alterados: `app.js`, `index.html`.

---

## Opção 1 — Reverter pelo git (recomendado)

Cria um commit novo que desfaz o anterior. Não apaga histórico e já sobe pro ar.

```
cd C:\Users\Pichau\Desktop\Gerador-OS
git revert --no-edit ea8d984
git push origin main
```

Para desfazer o revert (voltar a ter a alteração), repita com o hash do commit de revert.

## Opção 2 — Voltar ao ponto marcado

Existe a tag `pre-gramatura-cor-tecido` apontando para o estado anterior.
Restaura só os dois arquivos, sem mexer no resto do histórico:

```
cd C:\Users\Pichau\Desktop\Gerador-OS
git checkout pre-gramatura-cor-tecido -- app.js index.html
git commit -m "Volta app.js e index.html ao estado pre-gramatura-cor-tecido"
git push origin main
```

## Opção 3 — Cópia dos arquivos (sem git)

```
cd C:\Users\Pichau\Desktop\Gerador-OS
copy backups-codigo\app.js.20260721-pre-gramatura-cor-tecido.ANTES app.js
copy backups-codigo\index.html.20260721-pre-gramatura-cor-tecido.ANTES index.html
```

Depois commite e faça push, senão o site online continua com a versão nova.

---

## Atenção ao cache do navegador

O `index.html` carrega `app.js?v=2026-07-21a`. Ao reverter, o `?v=` volta para
`2026-07-20h`, que o navegador provavelmente ainda tem em cache — nesse caso a
reversão funciona. Se ficar em dúvida, force o recarregamento com **Ctrl+F5**.

## O que NÃO é revertido

Nada de dado é alterado por essas opções — elas só mexem em código. Se você já
tiver **renomeado as cores** no cadastro (de "Preto" para "Preto Malha Algodão")
ou preenchido gramaturas, isso fica no Supabase e continua lá após o revert.

Com o código antigo, cores renomeadas ainda funcionam para o cálculo em kg
(a busca de gramatura sempre foi por nome), mas voltam a perder:

- a sigla do SKU auto-preenchida em cores novas;
- a ordem canônica de cores pelo desc do desenho;
- o casamento das compras da Contabilidade com o saldo (o saldo do tecido pode
  aparecer rachado em duas linhas: uma "Preto" e outra "Preto Malha Algodão").

Para desfazer também os dados, restaure um snapshot do Supabase da data anterior.
