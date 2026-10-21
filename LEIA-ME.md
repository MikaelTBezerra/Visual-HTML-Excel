# Visual HTML GTF — add-in de conteúdo para Excel

Equivalente ao visual "HTML Content" do Power BI, só que dentro do Excel: uma caixa
embutida na planilha que roda o seu HTML/CSS/JS e lê os dados das tabelas do Excel.

## Como funciona

- O **código do add-in** (esta pasta) fica hospedado num endereço HTTPS. Ele não contém dados.
- O **HTML do seu dashboard** e o nome das tabelas ficam salvos **dentro do .xlsx**.
  Cada visual inserido tem o seu próprio HTML.
- Os **dados** são lidos das tabelas do Excel na hora, no computador de quem abre o arquivo.
- Quem pode ver é quem tem acesso ao arquivo no SharePoint/OneDrive. O Excel, ou seja,
  a permissão do arquivo, continua sendo o controle de acesso.

No seu HTML você tem:

| Variável | Conteúdo |
|---|---|
| `window.DADOS` | linhas da 1ª tabela informada: `[{Coluna: valor, ...}]` |
| `window.TABELAS["Nome"]` | linhas de qualquer tabela informada |
| `window.dataExcel(n)` | converte número serial de data do Excel em `Date` |

O visual se redesenha sozinho quando a tabela muda ou quando as fórmulas recalculam.

## Arquivos

```
index.html   página do add-in
app.js       lógica (lê tabelas, salva HTML na pasta, renderiza)
app.css      estilo do editor
assets/      ícones
manifest.xml "cadastro" do add-in que você entrega ao Excel ou à TI
```

## Passo 1 — Hospedar a pasta em HTTPS

O Excel só carrega add-ins de um endereço HTTPS. O SharePoint não serve: ele baixa o
HTML em vez de executar. Opções, da mais corporativa à mais rápida:

1. **Servidor/site interno da TI** ou **Azure Static Web Apps** da empresa (plano grátis).
2. **GitHub Pages** para piloto. O código fica público, mas ele não tem dado nenhum:
   os dados e o HTML dos dashboards ficam no .xlsx. Mesmo assim, combine com a TI.

Basta subir os arquivos como estão (site estático, sem build).

## Passo 2 — Ajustar o manifesto

No `manifest.xml`, troque todas as ocorrências de `https://SEU-ENDERECO` pelo endereço
real, sem barra no final. Exemplo: `https://gtf.github.io/visual-html-gtf`.

## Passo 3 — Testar sozinho (sem TI)

**Excel na web:** abra uma pasta no navegador > Página Inicial > Suplementos >
Mais Configurações > **Carregar Meu Suplemento** > escolha o `manifest.xml`.

Depois: **Inserir > Suplementos > Visual HTML GTF**. A caixa aparece sobre a planilha.
Clique em **Inserir exemplo**, informe o nome de uma tabela (ex.: `Tabela1`) e em
**Aplicar e salvar**.

## Passo 4 — Publicar só para a empresa

Peça a um administrador do Microsoft 365:

> Centro de administração do Microsoft 365 > Configurações > **Aplicativos integrados** >
> Carregar aplicativos personalizados > enviar o `manifest.xml` > atribuir aos usuários
> ou ao grupo desejado.

O add-in aparece para todos do grupo sem instalar nada no PC. Pode levar algumas horas
para propagar. Quem está fora da empresa não recebe o add-in.

## Usar no dia a dia

- Passe o mouse no visual e clique em **⚙** para editar o HTML ou as tabelas.
- Os nomes das tabelas devem ser exatamente os do Excel (Design da Tabela > Nome).
- Prefira bibliotecas por CDN (`cdn.jsdelivr.net`, `cdnjs.cloudflare.com`) em vez de colar
  o código inteiro da biblioteca: o HTML fica salvo dentro do .xlsx.

## Limitações conhecidas

- Precisa de internet para carregar o add-in e as bibliotecas por CDN.
- Quem abrir o arquivo **sem** o add-in liberado vê uma imagem estática do visual
  (snapshot) em vez do dashboard interativo. Essa imagem mostra os dados da última
  vez que alguém com o add-in salvou o arquivo.
- Quem tem acesso só de leitura vê o visual, mas não consegue salvar alterações nele.
- Datas chegam como número serial: use `dataExcel(valor)`.
