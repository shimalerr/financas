# Finanças — guia de instalação

App pessoal (iPhone + PC) ligada ao Excel `Financas_Pessoais.xlsx` no OneDrive.

- **iPhone e PC:** Resumo (com «O que me sobra» de qualquer mês), movimentos (com «Dividir com alguém»), despesas planeadas, dívidas como no Splitwise, dias de trabalho. Tudo se pode editar ou apagar no separador Registos.
- A app **nunca altera o Excel**. Guarda tudo no ficheiro `Financas_iPhone.csv`, na mesma pasta do OneDrive, e o Excel importa-o pela folha **Dados**. Uma vez por dia, a app guarda também uma cópia em `Financas_iPhone_backup.csv`.
- Do Excel, a app só lê a configuração: contas e saldos iniciais, categorias, fixos, dias marcados, ETFs e reserva do carro. Faz as contas sozinha, com as mesmas regras do Excel.

São 4 partes, feitas uma única vez (cerca de 30 minutos no total).

---

## Parte 1 — Preparar o Excel (no PC)

1. Fecha o Excel. Substitui o teu `Financas_Pessoais.xlsx` no OneDrive pela versão nova, com o mesmo nome e na mesma pasta.
2. Copia o ficheiro `Financas_iPhone.csv` (só tem a linha de títulos) para **a mesma pasta** do OneDrive.
3. Abre o Excel e vai a **Dados › Obter Dados › De Outras Origens › Consulta em Branco**.
4. No editor do Power Query, abre **Editor Avançado**, apaga tudo e cola o código abaixo. **Troca o caminho** pelo caminho real do CSV no teu PC: no Explorador de Ficheiros, clica no CSV com Shift + botão direito › *Copiar como caminho*.

```
let
    Origem = Csv.Document(File.Contents("C:\Users\O_TEU_UTILIZADOR\OneDrive\PASTA\Financas_iPhone.csv"),
        [Delimiter=",", Columns=11, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Cabecalhos = Table.PromoteHeaders(Origem, [PromoteAllScalars=true]),
    Tipos = Table.TransformColumnTypes(Cabecalhos, {
        {"ID", type text}, {"Registado", type text}, {"Registo", type text}, {"Data", type date},
        {"Descricao", type text}, {"Tipo", type text}, {"Categoria", type text}, {"Conta", type text},
        {"Valor", type number}, {"Divida", type text}, {"Nota", type text}}, "en-US")
in
    Tipos
```

5. Clica em **Concluído**. No painel da direita, muda o nome da consulta para `DoiPhone`.
6. Clica em **Fechar e Carregar Para…**, depois em **Tabela** › **Folha de cálculo existente**, e escreve `=Dados!$A$1` › OK.
7. Para atualizar sozinho: em **Dados › Consultas e Ligações**, clica com o botão direito em `DoiPhone` › **Propriedades** e ativa:
   - ☑ *Atualizar dados ao abrir o ficheiro*
   - ☑ *Atualizar a cada 5 minutos*
8. Guarda.

> Na folha **Dados** não se escreve nada: tudo o que lá está vem do CSV.
> Na folha **Dias de Trabalho**, os dias e os salários marcados no iPhone aparecem a **azul**.

---

## Parte 2 — Registar a app na Microsoft (para o login)

1. Entra em **https://portal.azure.com** com a tua conta Microsoft pessoal.
   - Se aparecer um erro de diretório ou de "tenant", cria primeiro uma conta Azure gratuita em **https://azure.microsoft.com/free**. É grátis, mas pode pedir um cartão só para verificação. Depois volta ao portal.
2. Procura **Microsoft Entra ID** › **Registos de aplicações** (App registrations) › **Novo registo**.
   - **Nome:** `Finanças`
   - **Tipos de conta suportados:** *Apenas contas Microsoft pessoais*
   - **URI de redirecionamento:** escolhe **Aplicação de página única (SPA)** e escreve
     `https://O_TEU_UTILIZADOR_GITHUB.github.io/financas/` (com a barra no fim)
   - Clica em **Registar**.
3. Copia o **ID da Aplicação (cliente)**. É um código do tipo `1a2b3c4d-…`.
4. *(Opcional)* Em **Permissões de API**, adiciona *Microsoft Graph › Delegadas*: `Files.ReadWrite` e `offline_access`. Se não o fizeres, a Microsoft pede-te autorização no primeiro login.

Não precisas de criar nenhum segredo nem palavra-passe da app.

---

## Parte 3 — Publicar a app (GitHub Pages, grátis)

1. Cria uma conta em **https://github.com** (o nome de utilizador vai aparecer no endereço da app).
2. Clica em **New repository** › nome `financas` › **Public** › **Create repository**.
3. Clica em **uploading an existing file** e arrasta **o conteúdo** da pasta da app (`index.html`, `app.js`, `xlsx-lite.js`, `styles.css`, `sw.js`, `manifest.webmanifest`, a pasta `icons` e este `README.md`) › **Commit changes**.
4. Vai a **Settings › Pages**. Em *Source*, escolhe **Deploy from a branch**, depois **main** e **/ (root)** › **Save**.
5. Passados 1 a 2 minutos, a app fica em `https://O_TEU_UTILIZADOR_GITHUB.github.io/financas/`.

> O repositório é público, mas só contém o código da app. Os teus dados nunca vão para o GitHub: ficam no teu OneDrive e no teu telemóvel.

---

## Parte 4 — Instalar e ligar

**iPhone**
1. Abre o endereço da app no **Safari** › botão **Partilhar** › **Adicionar ao ecrã principal**.
2. Abre a app a partir do ícone e cola o **Client ID** da Parte 2.
3. Toca em **Entrar com a conta Microsoft** e aceita as permissões.
4. Procura `Financas` e toca em **Usar**, ao lado do teu Excel.

**PC**
- Abre o mesmo endereço no Edge ou no Chrome e repete os passos 2 a 4.
- Se quiseres, instala-a como app: no menu do Edge, **Aplicações › Instalar este site como aplicação**.

---

## Como funciona no dia a dia

- **Registar** na app grava logo no `Financas_iPhone.csv`, no OneDrive. Sem rede, fica guardado no telemóvel e é enviado quando a rede voltar.
- **Movimentos, planeadas e dívidas** só se registam na app. O Excel recebe-os na folha **Dados** (ao abrir, de 5 em 5 minutos ou em Dados › Atualizar Tudo).
- **Contas, categorias, fixos, metas, ETFs e reserva do carro** continuam a mudar-se no Excel. Depois de guardares, a app lê-os sozinha (ou em ⚙ › Voltar a ler o Excel).
- **Dividir uma despesa:** no +, liga «Dividir com alguém», escreve com quem, quem pagou e a tua parte. A tua parte conta como despesa; a outra parte fica como dívida.
- **Dívidas:** podem ter prazo. Pagar uma parte baixa o que falta. O que falta conta no «O que me sobra» do mês do prazo.
- **O que me sobra** = receitas do mês − fixos − variáveis − planeadas do mês (pagas e por pagar) − dívidas do mês − reserva do carro. Os ETFs ficam em «A pôr de parte».
- Despesas pagas pela conta da reserva do carro (Conta Caixa) não descontam outra vez no que sobra, porque já saíram na reserva mensal.

## Problemas comuns

| Sintoma | Solução |
|---|---|
| Erro *redirect_uri* no login | O endereço registado na Parte 2 tem de ser **exatamente** igual ao da app, com a barra final. A app mostra o endereço certo no ecrã de configuração. |
| No Excel, a consulta dá erro "ficheiro não encontrado" | Confirma o caminho do CSV no código do Power Query (Parte 1, passo 4). |
| Valores com vírgulas estranhas no Excel | O código tem de manter o `"en-US"` no fim da linha `Tipos`. A app escreve os valores com ponto decimal de propósito. |
| A app não mostra alterações recentes do Excel | O Excel tem de estar guardado (o salvamento automático do OneDrive faz isso). Na app, ⚙ › Voltar a ler o Excel. |
| Depois de atualizar a app no GitHub, o iPhone mostra a versão antiga | Fecha a app (desliza para cima) e abre outra vez. |
