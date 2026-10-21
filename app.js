"use strict";
/*
 * Visual HTML GTF — content add-in para Excel.
 * Guarda o HTML e o nome das tabelas DENTRO da pasta de trabalho (settings do add-in),
 * lê as tabelas via Office.js e renderiza o HTML num iframe isolado (sandbox).
 * O HTML recebe:
 *   window.DADOS          -> linhas da 1ª tabela, como [{Coluna: valor, ...}]
 *   window.TABELAS        -> { "NomeTabela": [...], ... } com todas as tabelas
 *   window.dataExcel(n)   -> converte número serial de data do Excel em Date
 */

const CHAVE_HTML = "visualHtml";
const CHAVE_TABELAS = "visualTabelas";
const LT = "<"; // evita escrever tags <script> literais neste arquivo

const $ = (id) => document.getElementById(id);

let estado = { html: "", tabelas: [] };
let ultimoJson = null;
let timer = null;
const monitoradas = new Set();
let calculoMonitorado = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    mostrarMensagem("Este visual funciona apenas no Excel.");
    return;
  }
  const s = Office.context.document.settings;
  estado.html = s.get(CHAVE_HTML) || "";
  estado.tabelas = s.get(CHAVE_TABELAS) || [];
  ligarBotoes();

  if (!estado.html) abrirEditor();
  else renderizar(true).catch(mostrarErro);
});

/* ---------- editor ---------- */

function ligarBotoes() {
  $("btnEditar").onclick = abrirEditor;
  $("btnCancelar").onclick = fecharEditor;
  $("btnExemplo").onclick = () => { $("codigo").value = EXEMPLO; };
  $("btnSalvar").onclick = () => salvar().catch(mostrarErro);
}

function abrirEditor() {
  $("codigo").value = estado.html;
  $("tabelas").value = estado.tabelas.join(", ");
  status("");
  $("editor").hidden = false;
  $("btnEditar").hidden = true;
}

function fecharEditor() {
  if (!estado.html) return; // sem HTML não há o que mostrar atrás do editor
  $("editor").hidden = true;
  $("btnEditar").hidden = false;
}

async function salvar() {
  const html = $("codigo").value;
  const tabelas = $("tabelas").value.split(/[;,]/).map((t) => t.trim()).filter(Boolean);
  estado = { html, tabelas };

  const s = Office.context.document.settings;
  s.set(CHAVE_HTML, html);
  s.set(CHAVE_TABELAS, tabelas);
  try {
    await new Promise((ok, falha) =>
      s.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? ok() : falha(r.error)))
    );
  } catch (e) {
    // arquivo somente leitura, por exemplo: aplica na tela mesmo assim
    status("Aplicado, mas não foi salvo na pasta: " + (e && e.message ? e.message : e), true);
    $("editor").hidden = true;
    $("btnEditar").hidden = false;
    await renderizar(true);
    return;
  }
  fecharEditor();
  await renderizar(true);
}

/* ---------- dados ---------- */

async function lerTabelas(nomes) {
  const resultado = {};
  if (!nomes.length) return resultado;

  await Excel.run(async (ctx) => {
    const objs = nomes.map((n) => ctx.workbook.tables.getItemOrNullObject(n));
    await ctx.sync();

    const faixas = objs.map((t) =>
      t.isNullObject
        ? null
        : { cab: t.getHeaderRowRange().load("values"), corpo: t.getDataBodyRange().load("values") }
    );
    await ctx.sync();

    faixas.forEach((f, i) => {
      if (!f) { resultado[nomes[i]] = null; return; }
      const cab = f.cab.values[0].map(String);
      resultado[nomes[i]] = f.corpo.values
        .filter((linha) => linha.some((v) => v !== "" && v !== null))
        .map((linha) => Object.fromEntries(cab.map((c, j) => [c, linha[j]])));
    });
  });
  return resultado;
}

async function monitorar(nomes) {
  const novos = nomes.filter((n) => !monitoradas.has(n));
  await Excel.run(async (ctx) => {
    const objs = novos.map((n) => ctx.workbook.tables.getItemOrNullObject(n));
    await ctx.sync();
    objs.forEach((t, i) => {
      if (t.isNullObject) return;
      const nome = novos[i];
      t.onChanged.add(async () => { if (estado.tabelas.includes(nome)) agendar(); });
      monitoradas.add(nome);
    });
    // fórmulas recalculadas não disparam onChanged da tabela; o recálculo cobre esse caso
    if (!calculoMonitorado) {
      ctx.workbook.worksheets.onCalculated.add(async () => agendar());
      calculoMonitorado = true;
    }
    await ctx.sync();
  });
}

function agendar() {
  clearTimeout(timer);
  timer = setTimeout(() => renderizar(false).catch(mostrarErro), 600);
}

/* ---------- renderização ---------- */

async function renderizar(forcar) {
  if (!estado.html) return;

  const todas = await lerTabelas(estado.tabelas);
  const json = JSON.stringify(todas);
  if (!forcar && json === ultimoJson) return; // nada mudou: não pisca a tela
  ultimoJson = json;

  const faltando = estado.tabelas.filter((n) => todas[n] === null);
  if (faltando.length) status("Tabela não encontrada: " + faltando.join(", "), true);

  const primeira = estado.tabelas.length ? todas[estado.tabelas[0]] : null;
  $("view").srcdoc = montarDocumento(estado.html, primeira, todas);

  monitorar(estado.tabelas).catch((e) => console.error(e));
}

function montarDocumento(html, primeira, todas) {
  const seguro = (v) =>
    JSON.stringify(v === undefined ? null : v)
      .replace(/</g, "\\u003c")
      .replace(new RegExp("\\u2028", "g"), "\\u2028")
      .replace(new RegExp("\\u2029", "g"), "\\u2029");

  const js =
    "window.DADOS=" + seguro(primeira) + ";" +
    "window.TABELAS=" + seguro(todas) + ";" +
    "window.dataExcel=function(n){if(typeof n!=='number')return n;" +
    "var d=new Date(Math.round((n-25569)*864e5));" +
    "return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate(),d.getUTCHours(),d.getUTCMinutes());};";
  const injecao = LT + "script>" + js + LT + "/script>";

  const reHead = new RegExp(LT + "head[^>]*>", "i");
  const reHtml = new RegExp(LT + "html[^>]*>", "i");
  if (reHead.test(html)) return html.replace(reHead, (m) => m + injecao);
  if (reHtml.test(html)) return html.replace(reHtml, (m) => m + LT + "head>" + injecao + LT + "/head>");
  return injecao + html;
}

/* ---------- mensagens ---------- */

function status(texto, erro) {
  const el = $("status");
  el.textContent = texto;
  el.className = erro ? "erro" : "";
}

function mostrarErro(e) {
  console.error(e);
  const msg = e && e.message ? e.message : String(e);
  if ($("editor").hidden) abrirEditor(); // não apaga o que está sendo digitado
  status("Erro: " + msg, true);
}

function mostrarMensagem(texto) {
  $("view").srcdoc =
    LT + "body style='font-family:Segoe UI,sans-serif;color:#6b7280;padding:12px'>" + texto;
}

/* ---------- exemplo ---------- */

const EXEMPLO = [
  LT + "!doctype html>",
  LT + "html lang='pt-BR'>" + LT + "head>" + LT + "meta charset='utf-8'>",
  LT + "script src='https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js'>" + LT + "/script>",
  LT + "style>",
  "  body{font-family:Segoe UI,sans-serif;margin:12px;color:#1f2937}",
  "  h3{margin:0 0 8px;font-size:14px}",
  "  .kpi{font-size:22px;font-weight:700;color:#1a7f37;margin-bottom:8px}",
  "  .area{position:relative;height:calc(100vh - 90px)}",
  LT + "/style>" + LT + "/head>",
  LT + "body>",
  LT + "h3 id='titulo'>Exemplo de visual" + LT + "/h3>",
  LT + "div class='kpi' id='kpi'>" + LT + "/div>",
  LT + "div class='area'>" + LT + "canvas id='g'>" + LT + "/canvas>" + LT + "/div>",
  LT + "script>",
  "const linhas = window.DADOS || [];",
  "if (!linhas.length) {",
  "  document.getElementById('titulo').textContent = 'Informe o nome de uma tabela do Excel no editor (⚙).';",
  "} else {",
  "  const cols = Object.keys(linhas[0]);",
  "  const rotulo = cols[0];",
  "  const valor = cols.find(c => c !== rotulo && typeof linhas[0][c] === 'number') || cols[1];",
  "  const total = linhas.reduce((s, l) => s + (Number(l[valor]) || 0), 0);",
  "  document.getElementById('titulo').textContent = valor + ' por ' + rotulo;",
  "  document.getElementById('kpi').textContent = 'Total: ' + total.toLocaleString('pt-BR');",
  "  new Chart(document.getElementById('g'), {",
  "    type: 'bar',",
  "    data: { labels: linhas.map(l => l[rotulo]), datasets: [{ label: valor, data: linhas.map(l => l[valor]), backgroundColor: '#1a7f37' }] },",
  "    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }",
  "  });",
  "}",
  LT + "/script>",
  LT + "/body>" + LT + "/html>",
].join("\n");
