// ============================================================================
// Planejamento VD (venda direta) · módulo separado do painel do turno.
// Reproduz a aba "RESUMO COMPLETO" da planilha ACOMPANHAMENTO DE PRODUÇÃO:
//   planejamento de carregamento (NFs por veículo)
//   + faturamento (pedido+sub e peso bruto por NF, opcional)
//   + Pedidos Query (status de cada pedido no WMS)
//   = quantidade de pedidos e peso por status.
// Pedido que não aparece mais na Pedidos Query conta como "Conferido",
// igual à fórmula da planilha. Dados ficam salvos no navegador.
// ============================================================================
import * as XLSX from "xlsx";
import LOGO_W from "./assets/logo-branco.png";
import DUO from "./assets/pet-dupla.png";
import { entregar, imagemDe } from "./compartilhar.js";

const $ = (id) => document.getElementById(id);
const fmt = (n, d = 0) => (n == null || isNaN(n) ? "—" : n.toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d }));
const pct = (a, b) => (b ? fmt((a / b) * 100, 1) + "%" : "—");
const semAc = (x) => String(x ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
const chave = (x) => semAc(x).replace(/[\s_]+/g, "");
const esc = (x) => String(x ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const titulo = (s) => String(s || "").toLowerCase().replace(/(^|\s)\S/g, (m) => m.toUpperCase()).replace(/\b(Da|De|Do|Dos|Das|E|Me|Ltda|Eireli)\b/g, (m) => (m.length > 3 ? m.toUpperCase() : m.toLowerCase()));
const hora = (iso) => (iso ? new Date(iso).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : "—");
const diaTxt = (iso, o = { weekday: "long", day: "2-digit", month: "2-digit" }) => { if (!iso) return "—"; const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d).toLocaleDateString("pt-BR", o); };

// ordem e nomes iguais ao RESUMO COMPLETO
const STATUS = ["Conferido", "Em Curso de Conferência", "Em transição", "Separado", "Em Curso de Separação", "Planejamento concluído", "Reservado"];
const GRUPOS = [
  ["Pendente", ["Em Curso de Separação", "Planejamento concluído", "Reservado"]],
  ["Em Curso", ["Em Curso de Conferência", "Em transição", "Separado"]],
  ["Finalizado", ["Conferido"]],
];
const COR = {
  "Conferido": "#1f8a4c", "Em Curso de Conferência": "#2f6aa8", "Em transição": "#5b8fc7", "Separado": "#8db3dc",
  "Em Curso de Separação": "#e8a84a", "Planejamento concluído": "#ff6900", "Reservado": "#b4530a", "Outros": "#8a8f96", "Sem status": "#b5ada6",
};
const COR_G = { Pendente: "#ff6900", "Em Curso": "#2f6aa8", Finalizado: "#1f8a4c" };
const CANON = Object.fromEntries(STATUS.map((s) => [chave(s), s]));
const canon = (desc) => CANON[chave(desc)] || "Outros";
const grupoDe = (st) => (GRUPOS.find(([, l]) => l.includes(st)) || [""])[0];

// ---------- armazenamento ----------
const KP = "planejamentoVD.planos", KQ = "planejamentoVD.query", KH = "planejamentoVD.historico", KF = "planejamentoVD.faturamento";
const ler = (k, padrao) => { try { const v = JSON.parse(localStorage.getItem(k) || "null"); return v ?? padrao; } catch (e) { return padrao; } };
const gravar = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };
let planos = ler(KP, {});          // { "2026-10-13": {iso, titulo, arquivo, em, veiculos:[], linhas:[]} }
let query = ler(KQ, null);         // { arquivo, em, n, mapa:{ "O-283246/1": "RESERVADO" } }
let historico = ler(KH, []);       // [{ em, iso, n, kg, g:{Pendente:[n,kg],...} }]
let fatMeta = ler(KF, null);       // { arquivo, em, n }  (o faturamento em si fica só na memória)
let fat = null;                    // Map NF -> { ped, peso }
let selIso = null, filtro = "Todos";

function salvarPlanos() {
  const isos = Object.keys(planos).sort();
  while (isos.length > 10) delete planos[isos.shift()]; // guarda os 10 planejamentos mais recentes
  gravar(KP, planos);
}

// ---------- leitura das planilhas ----------
const num = (v) => { if (typeof v === "number") return v; let x = String(v ?? "").replace(/kg/i, "").trim(); if (!x) return NaN; if (/,\d+$/.test(x)) x = x.replace(/\./g, "").replace(",", "."); return parseFloat(x); };
const pedKey = (ped, sub) => { let p = String(ped ?? "").trim(); if (!p) return ""; if (p.includes("/")) return p; p = p.replace(/-\d{2}$/, ""); return p + "/" + (sub == null || sub === "" ? 1 : Number(sub) || sub); };
const idx = (row, ...nomes) => { const r = row.map(chave); for (const n of nomes) { const i = r.indexOf(chave(n)); if (i >= 0) return i; } return -1; };

function lerPlano(rows, arquivo) {
  // blocos por veículo: "Sessão | Sessão+rota | Veículo | Motorista | Rota | Peso", depois "Tipo de Veículo | Exportado | Nome Rota",
  // depois "Sequência | Nota Fiscal | Pedido | Razão Social | ... | Peso | Tipo cliente | Volume"
  let iso = null, tituloTxt = "";
  const veiculos = [], linhas = [];
  let v = null, cab = null;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!iso) {
      const t = r.find((c) => typeof c === "string" && /\d{2}\/\d{2}\/\d{4}/.test(c) && /ENTREGA/i.test(c));
      if (t) { const m = t.match(/(\d{2})\/(\d{2})\/(\d{4})/); iso = `${m[3]}-${m[2]}-${m[1]}`; tituloTxt = t.trim(); }
    }
    const iSes = idx(r, "Sessão"), iVei = idx(r, "Veículo");
    if (iSes >= 0 && iVei >= 0 && rows[i + 1]) {
      const n = rows[i + 1], iMot = idx(r, "Motorista"), iRota = idx(r, "Rota"), iPeso = idx(r, "Peso");
      v = { sessao: n[iSes], placa: String(n[iVei] ?? "").trim(), motorista: String(n[iMot] ?? "").trim(), rota: n[iRota], peso: num(n[iPeso]), tipo: "", nomeRota: "" };
      veiculos.push(v); cab = null; continue;
    }
    const iTipo = idx(r, "Tipo de Veículo");
    if (iTipo >= 0 && v && rows[i + 1]) { const n = rows[i + 1], iNome = idx(r, "Nome Rota"); v.tipo = String(n[iTipo] ?? "").trim(); v.nomeRota = String(n[iNome] ?? "").trim(); continue; }
    const iSeq = idx(r, "Sequência"), iNf = idx(r, "Nota Fiscal"), iPed = idx(r, "Pedido");
    if (iSeq >= 0 && iNf >= 0 && iPed >= 0) {
      cab = { iSeq, iNf, iPed, iCli: idx(r, "Razão Social"), iCid: idx(r, "Cidade"), iPeso: r.findIndex((c, k) => k > iPed && chave(c) === "PESO"), iTipoCli: idx(r, "Tipo cliente"), iVol: idx(r, "Volume") };
      continue;
    }
    if (cab) {
      const nf = num(r[cab.iNf]);
      if (!isFinite(nf) || !String(r[cab.iPed] ?? "").trim()) { if (String(r[cab.iSeq] ?? "").trim()) cab = null; continue; }
      linhas.push({ vi: veiculos.length - 1, seq: r[cab.iSeq], nf, pedPlano: String(r[cab.iPed]).trim(), pesoPlano: num(r[cab.iPeso]) || 0,
        cli: String(r[cab.iCli] ?? "").trim(), cid: String(r[cab.iCid] ?? "").trim(), tipoCli: String(r[cab.iTipoCli] ?? "").trim(), vol: num(r[cab.iVol]) || 0 });
    }
  }
  if (!linhas.length) return null;
  return { iso, titulo: tituloTxt, arquivo, em: new Date().toISOString(), veiculos, linhas };
}

function lerFaturamento(rows) {
  const hi = rows.findIndex((r) => idx(r, "Nota Fiscal") >= 0 && idx(r, "Pedido") >= 0 && r.some((c) => /PESO.*BRUTO/.test(semAc(c))) && idx(r, "Sequência") < 0);
  if (hi < 0) return null;
  const h = rows[hi], iNf = idx(h, "Nota Fiscal"), iPed = idx(h, "Pedido"), iSub = idx(h, "SUB"), iPS = idx(h, "PEDIDO+SUB");
  const iPeso = h.findIndex((c) => /PESO.*BRUTO/.test(semAc(c)));
  const mapa = new Map(), conta = {};
  for (const r of rows.slice(hi + 1)) {
    const nf = num(r[iNf]); const ped = String(r[iPed] ?? "").trim();
    if (!isFinite(nf) || !ped) continue;
    conta[ped] = (conta[ped] || 0) + 1; // mesmo critério do CONT.SE acumulado da planilha
    const ps = iPS >= 0 && String(r[iPS] ?? "").includes("/") ? String(r[iPS]).trim() : `${ped}/${iSub >= 0 && r[iSub] !== "" ? r[iSub] : conta[ped]}`;
    if (!mapa.has(nf)) mapa.set(nf, { ped: ps, peso: num(r[iPeso]) });
  }
  return mapa.size ? mapa : null;
}

function lerQuery(rows) {
  const hi = rows.findIndex((r) => idx(r, "PEDIDO") >= 0 && idx(r, "DESCRIÇÃO_STATUS", "DESCRICAO STATUS", "DESCRIÇÃO STATUS") >= 0 && idx(r, "Sequência") < 0);
  if (hi < 0) return null;
  const h = rows[hi], iPed = idx(h, "PEDIDO"), iSub = idx(h, "SUB"), iDesc = idx(h, "DESCRIÇÃO_STATUS", "DESCRICAO STATUS", "DESCRIÇÃO STATUS");
  const mapa = {};
  for (const r of rows.slice(hi + 1)) { const k = pedKey(r[iPed], iSub >= 0 ? r[iSub] : 1); if (k && r[iDesc]) mapa[k] = String(r[iDesc]).trim(); }
  const n = Object.keys(mapa).length;
  return n ? { mapa, n } : null;
}

function enriquecer(plano) { // pedido+sub e peso bruto vindos do faturamento, como as colunas O e P da planilha
  for (const l of plano.linhas) {
    const f = fat && fat.get(l.nf);
    l.ped = f ? f.ped : l.pedPlano; l.peso = f && isFinite(f.peso) ? f.peso : l.pesoPlano; l.fonte = f ? "fat" : "plano";
  }
}

async function lerArquivos(lista) {
  const st = $("plStatus"); st.innerHTML = `<div class="prev">Lendo ${lista.length > 1 ? lista.length + " arquivos" : esc(lista[0].name)}…</div>`;
  const achados = { planos: [], fat: null, query: null }; const nomes = { fat: "", query: "" }; const ignorados = [];
  try {
    for (const file of lista) {
      let tabelas = [];
      if (/\.csv$/i.test(file.name)) {
        const txt = await file.text(); const l0 = txt.split("\n")[0];
        const sc = (l0.match(/;/g) || []).length > (l0.match(/,/g) || []).length ? ";" : l0.includes("\t") ? "\t" : ",";
        tabelas.push(txt.split(/\r?\n/).map((x) => x.split(sc).map((c) => c.replace(/^"|"$/g, "").trim())));
      } else {
        const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellFormula: false, cellHTML: false, cellStyles: false });
        tabelas = wb.SheetNames.filter((sn) => !/^RESUMO|DEMANDA/i.test(sn)).map((sn) => XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: true, blankrows: false }));
      }
      let usou = false;
      for (const rows of tabelas) {
        const q = lerQuery(rows); if (q) { achados.query = q; nomes.query = file.name; usou = true; continue; }
        const f = lerFaturamento(rows); if (f) { achados.fat = f; nomes.fat = file.name; usou = true; continue; }
        const p = lerPlano(rows, file.name); if (p) { achados.planos.push(p); usou = true; }
      }
      if (!usou) ignorados.push(file.name);
    }
  } catch (e) {
    st.innerHTML = `<div class="prev err">Não consegui ler o arquivo. Confira se é .xlsx, .xlsm ou .csv e se não está protegido por senha.</div>`; return;
  }
  const feito = [];
  if (achados.fat) { fat = achados.fat; fatMeta = { arquivo: nomes.fat, em: new Date().toISOString(), n: fat.size }; gravar(KF, fatMeta); Object.values(planos).forEach(enriquecer); feito.push(`faturamento (${fmt(fat.size)} NFs)`); }
  let semData = 0;
  for (const p of achados.planos) {
    if (!p.iso) { p.iso = proximoDia(); semData++; }
    enriquecer(p); planos[p.iso] = p; selIso = p.iso; feito.push(`planejamento de ${diaTxt(p.iso, { day: "2-digit", month: "2-digit" })} (${p.linhas.length} NFs)`);
  }
  if (achados.planos.length > 1) selIso = achados.planos.map((p) => p.iso).sort().pop();
  if (achados.planos.length || achados.fat) salvarPlanos();
  if (achados.query) {
    query = { arquivo: nomes.query, em: new Date().toISOString(), n: achados.query.n, mapa: achados.query.mapa };
    if (!gravar(KQ, query)) feito.push("(aviso: a Pedidos Query é grande demais para ficar salva neste navegador)");
    feito.unshift(`Pedidos Query (${fmt(achados.query.n)} pedidos)`);
    registrarHistorico();
  }
  $("plFile").value = "";
  st.innerHTML = feito.length
    ? `<div class="prev"><span class="ok-msg">Atualizado: ${feito.join(", ")}.</span>${semData ? `<span class="sub">Não achei a data de entrega no planejamento; usei ${diaTxt(proximoDia(), { day: "2-digit", month: "2-digit" })}.</span>` : ""}${ignorados.length ? `<span class="sub">Sem dados reconhecidos em: ${ignorados.map(esc).join(", ")}.</span>` : ""}</div>`
    : `<div class="prev err">Não reconheci nenhum relatório. O planejamento precisa das colunas Sequência, Nota Fiscal e Pedido; o faturamento, de Nota Fiscal, Pedido e Peso Total Bruto; a Pedidos Query, de PEDIDO e DESCRIÇÃO_STATUS.</div>`;
  render();
}
const proximoDia = () => { const d = new Date(); d.setDate(d.getDate() + 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };

// ---------- cálculo do resumo ----------
const statusDe = (l) => (!query ? "Sem status" : query.mapa[l.ped] ? canon(query.mapa[l.ped]) : "Conferido");
function resumo(plano) {
  const r = { st: {}, g: {}, n: 0, kg: 0, outros: {} };
  [...STATUS, "Outros", "Sem status"].forEach((s) => (r.st[s] = [0, 0]));
  GRUPOS.forEach(([g]) => (r.g[g] = [0, 0]));
  for (const l of plano.linhas) {
    const s = statusDe(l); r.st[s][0]++; r.st[s][1] += l.peso || 0; r.n++; r.kg += l.peso || 0;
    if (s === "Outros") r.outros[query.mapa[l.ped]] = (r.outros[query.mapa[l.ped]] || 0) + 1;
    const g = grupoDe(s); if (g) { r.g[g][0]++; r.g[g][1] += l.peso || 0; }
  }
  return r;
}
function registrarHistorico() {
  const iso = Object.keys(planos).sort().pop(); if (!iso) return;
  const r = resumo(planos[iso]);
  historico.push({ em: query.em, iso, n: r.n, kg: r.kg, g: r.g });
  historico = historico.slice(-200); gravar(KH, historico);
}
const anteriorDe = (iso) => Object.keys(planos).filter((k) => k < iso).sort().pop();

// ---------- telas ----------
function tabelaResumo(r, compacta) {
  const linha = (nome, [n, kg], cls = "", cor) => `<tr class="${cls}"><td class="l">${cor ? `<i class="pl-dot" style="background:${cor}"></i>` : ""}${nome}</td><td><span class="n">${fmt(n)}</span></td><td><span class="n">${fmt(kg, 2)}</span></td><td><span class="n">${pct(kg, r.kg)}</span></td></tr>`;
  const extra = ["Outros", "Sem status"].filter((s) => r.st[s][0]);
  return `<table class="pl-res-t${compacta ? " pl-c" : ""}"><thead><tr><th class="l"></th><th>Qtde Ped.</th><th>Peso (kg)</th><th>Porcentual</th></tr></thead><tbody>
    ${STATUS.map((s) => linha(s, r.st[s], r.st[s][0] ? "" : "zero", COR[s])).join("")}
    ${extra.map((s) => linha(s === "Outros" ? `Outros (${Object.keys(r.outros).map(esc).join(", ")})` : "Sem status", r.st[s], "", COR[s])).join("")}
    ${GRUPOS.map(([g]) => linha(g, r.g[g], "grp", COR_G[g])).join("")}
    ${linha("Total Geral", [r.n, r.kg], "tot")}
  </tbody></table>`;
}
function barraDemanda(r) {
  const ok = r.g["Finalizado"][1] + r.g["Em Curso"][1];
  return `<div class="pl-dem"><div class="pl-dem-bar">${GRUPOS.slice().reverse().map(([g]) => `<i style="width:${r.kg ? (r.g[g][1] / r.kg) * 100 : 0}%;background:${COR_G[g]}" title="${g}"></i>`).join("")}</div>
    <div class="pl-dem-leg"><span><b>${pct(ok, r.kg)}</b> concluído <small>(separado em diante)</small></span><span><b>${pct(r.g["Pendente"][1], r.kg)}</b> pendente</span></div></div>`;
}
function cardResumo(el, plano, rotulo) {
  if (!plano) { el.innerHTML = `<div class="pl-card-h"><h3>${rotulo}</h3></div><p class="sub pl-vazio">Sem planejamento salvo para comparar. Ele aparece aqui a partir do segundo dia de uso.</p>`; return null; }
  const r = resumo(plano);
  el.innerHTML = `<div class="pl-card-h"><h3>${rotulo} · ${diaTxt(plano.iso)}</h3><span class="sub">${esc(plano.titulo.replace(/^\d{2}\/\d{2}\/\d{4}\s*-\s*/, "")) || esc(plano.arquivo)}</span></div>${barraDemanda(r)}${tabelaResumo(r)}`;
  return r;
}

function render() {
  const isos = Object.keys(planos).sort().reverse();
  if (!selIso || !planos[selIso]) selIso = isos[0] || null;
  $("plData").innerHTML = isos.length ? isos.map((k) => `<option value="${k}" ${k === selIso ? "selected" : ""}>${diaTxt(k)}</option>`).join("") : `<option>Nenhum planejamento</option>`;
  $("plData").disabled = !isos.length; $("plEnviar").disabled = !selIso;

  // fontes carregadas
  const p = selIso && planos[selIso];
  const chip = (ok, nome, txt) => `<span class="pl-fonte ${ok ? "ok" : ""}"><b>${nome}</b> ${txt}</span>`;
  $("plFontes").innerHTML = [
    chip(!!p, "Planejamento", p ? `${diaTxt(p.iso, { day: "2-digit", month: "2-digit" })} · ${p.linhas.length} NFs · ${p.veiculos.length} veículos` : "não enviado"),
    chip(!!fatMeta, "Faturamento", fatMeta ? `${fmt(fatMeta.n)} NFs · ${hora(fatMeta.em)}` : "opcional"),
    chip(!!query, "Pedidos Query", query ? `${fmt(query.n)} pedidos · ${hora(query.em)}` : "não enviada"),
  ].join("");

  // alertas
  const al = [];
  if (!query) al.push("<b>Falta a Pedidos Query:</b> sem ela não dá para saber o status dos pedidos.");
  else if (new Date(query.em).toDateString() !== new Date().toDateString()) al.push(`<b>Pedidos Query de ${hora(query.em)}:</b> suba a do turno atual para o resumo ficar em dia.`);
  if (p && p.linhas.some((l) => l.fonte !== "fat")) al.push(`<b>${p.linhas.filter((l) => l.fonte !== "fat").length} NFs sem faturamento:</b> usei o pedido e o peso do próprio planejamento.`);
  $("plAlertas").innerHTML = al.length ? `<div class="callout"><div><span class="lbl" style="color:var(--warn)">Confira</span><ul>${al.map((a) => `<li>${a}</li>`).join("")}</ul></div></div>` : "";

  const r = cardResumo($("plAtual"), p, "Planejamento");
  const ant = p ? planos[anteriorDe(p.iso)] : null;
  const ra = cardResumo($("plAnterior"), ant, "Dia anterior");

  // comparação com a atualização anterior do mesmo planejamento
  const hs = historico.filter((h) => p && h.iso === p.iso);
  const prev = hs.length > 1 ? hs[hs.length - 2] : null;
  $("plComp").textContent = prev && r ? `Desde ${hora(prev.em)}: ${r.g["Finalizado"][0] - prev.g["Finalizado"][0] >= 0 ? "+" : ""}${r.g["Finalizado"][0] - prev.g["Finalizado"][0]} pedidos finalizados, pendente ${pct(prev.g["Pendente"][1], prev.kg)} → ${pct(r.g["Pendente"][1], r.kg)}` : "";

  renderVeiculos(p); renderPedidos(p); renderHist(); renderImagem(p, r, ant, ra);
}

function miniBarra(cont, total) {
  return `<span class="pl-mini">${[...STATUS, "Outros", "Sem status"].filter((s) => cont[s]).map((s) => `<i style="width:${(cont[s] / total) * 100}%;background:${COR[s]}" title="${s}: ${cont[s]}"></i>`).join("")}</span>`;
}
function renderVeiculos(p) {
  if (!p) { $("plVeic").innerHTML = ""; $("plVeicSub").textContent = ""; return; }
  const rows = p.veiculos.map((v, i) => {
    const ls = p.linhas.filter((l) => l.vi === i); const cont = {}; let kg = 0, fin = 0;
    ls.forEach((l) => { const s = statusDe(l); cont[s] = (cont[s] || 0) + 1; kg += l.peso || 0; if (s === "Conferido") fin++; });
    return { v, n: ls.length, kg, cont, fin };
  }).filter((x) => x.n);
  const completos = rows.filter((x) => x.fin === x.n).length;
  $("plVeicSub").textContent = query ? `${completos} de ${rows.length} veículos com todos os pedidos conferidos` : "";
  $("plVeic").innerHTML = `<thead><tr><th class="l">Rota</th><th class="l">Placa</th><th class="l">Transportadora / motorista</th><th class="l">Destino</th><th>NFs</th><th>Peso (kg)</th><th class="l">Status</th><th>Conferido</th></tr></thead><tbody>${rows.map((x) => `<tr>
    <td class="l"><span class="n">${esc(x.v.rota)}</span></td><td class="l"><b>${esc(x.v.placa)}</b><small class="pl-s">${esc(x.v.tipo)}</small></td><td class="l">${esc(titulo(x.v.motorista))}</td><td class="l">${esc(x.v.nomeRota)}</td>
    <td><span class="n">${x.n}</span></td><td><span class="n">${fmt(x.kg, 0)}</span></td><td class="l">${miniBarra(x.cont, x.n)}</td><td><span class="n ${x.fin === x.n ? "ok" : ""}">${x.fin}/${x.n}</span></td></tr>`).join("")}</tbody>`;
}
function renderPedidos(p) {
  if (!p) { $("plPed").innerHTML = ""; $("plFiltro").innerHTML = ""; return; }
  const ls = p.linhas.map((l) => ({ ...l, st: statusDe(l) }));
  const presentes = [...STATUS, "Outros", "Sem status"].filter((s) => ls.some((l) => l.st === s));
  if (filtro !== "Todos" && !presentes.includes(filtro)) filtro = "Todos";
  $("plFiltro").innerHTML = ["Todos", ...presentes].map((s) => `<button type="button" data-f="${esc(s)}" aria-pressed="${filtro === s}">${s === "Todos" ? `Todos (${ls.length})` : `${s} (${ls.filter((l) => l.st === s).length})`}</button>`).join("");
  const q = semAc($("plBusca").value);
  const vis = ls.filter((l) => (filtro === "Todos" || l.st === filtro) && (!q || semAc(`${l.nf} ${l.ped} ${l.cli} ${l.cid} ${(p.veiculos[l.vi] || {}).placa}`).includes(q)));
  $("plPed").innerHTML = `<thead><tr><th class="l">Status</th><th class="l">Rota · placa</th><th>Seq.</th><th>NF</th><th class="l">Pedido</th><th class="l">Cliente</th><th class="l">Cidade</th><th>Peso (kg)</th></tr></thead><tbody>${vis.slice(0, 600).map((l) => { const v = p.veiculos[l.vi] || {}; return `<tr>
    <td class="l"><span class="pl-badge" style="--c:${COR[l.st]}">${l.st}</span></td><td class="l">${esc(v.rota)} · ${esc(v.placa)}</td><td>${esc(l.seq)}</td><td>${l.nf}</td><td class="l">${esc(l.ped)}</td><td class="l">${esc(titulo(l.cli))}</td><td class="l">${esc(l.cid)}</td><td><span class="n">${fmt(l.peso, 2)}</span></td></tr>`; }).join("") || `<tr><td class="l" colspan="8">Nenhum pedido neste filtro.</td></tr>`}</tbody>`;
}
function renderHist() {
  const hs = historico.slice(-30).reverse();
  $("plHist").innerHTML = hs.length ? `<thead><tr><th class="l">Atualização</th><th class="l">Entregas de</th><th>Pedidos</th><th>Peso (kg)</th><th>Finalizado</th><th>Em curso</th><th>Pendente</th></tr></thead><tbody>${hs.map((h) => `<tr><td class="l">${hora(h.em)}</td><td class="l">${diaTxt(h.iso, { day: "2-digit", month: "2-digit" })}</td><td><span class="n">${h.n}</span></td><td><span class="n">${fmt(h.kg, 0)}</span></td>
    <td><span class="n">${pct(h.g["Finalizado"][1], h.kg)}</span></td><td><span class="n">${pct(h.g["Em Curso"][1], h.kg)}</span></td><td><span class="n">${pct(h.g["Pendente"][1], h.kg)}</span></td></tr>`).join("")}</tbody>`
    : `<tbody><tr><td class="l">Nenhuma atualização ainda. Cada Pedidos Query enviada registra uma linha aqui.</td></tr></tbody>`;
}

// ---------- imagem para o WhatsApp ----------
function renderImagem(p, r, ant, ra) {
  const pg = $("plPg");
  if (!p || !r) { pg.innerHTML = ""; return; }
  const now = new Date();
  const bloco = (rr, pl, rot) => rr
    ? `<div class="x-sec"><div class="x-band"><b>${rot}</b><span class="x-tag ${rr.g["Pendente"][0] ? "warn" : "ok"}">${diaTxt(pl.iso, { weekday: "short", day: "2-digit", month: "2-digit" })}</span></div>${barraDemanda(rr)}${tabelaResumo(rr, true)}</div>`
    : `<div class="x-sec"><div class="x-band"><b>${rot}</b></div><div class="x-obs"><span class="na">Sem planejamento salvo do dia anterior.</span></div></div>`;
  pg.innerHTML = `<div class="x-h"><img class="lg" src="${LOGO_W}" alt="PremieRpet"><div class="t">PLANEJAMENTO VD</div><div class="d">Planejamento operacional · status de ${query ? hora(query.em) : "—"} · Estab. 40</div><div class="sw"></div><img class="pets" src="${DUO}" alt=""></div>
    <div class="x-b pl-foto">${bloco(r, p, "ENTREGAS DO DIA")}</div>
    <div class="x-f"><b>Planejamento VD · ${diaTxt(p.iso, { day: "2-digit", month: "2-digit" })}</b><span>Gerado em ${now.toLocaleDateString("pt-BR")} às ${now.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}</span></div>`;
  ajustar();
}
function ajustar() { const st = $("plSt"), w = st.clientWidth; if (!w) return; const sc = Math.min(1.25, w / 800); $("plFit").style.transform = `scale(${sc})`; st.style.height = $("plPg").offsetHeight * sc + 2 + "px"; }
window.addEventListener("resize", ajustar);

async function enviar() {
  const b = $("plEnviar"), txt = b.textContent; b.disabled = true; b.textContent = "Gerando…";
  try {
    const blob = await imagemDe($("plPg"));
    await entregar([{ nome: `planejamento-vd-${selIso}.png`, blob }], `Planejamento VD · ${diaTxt(selIso, { day: "2-digit", month: "2-digit" })}`);
  } catch (e) { $("plStatus").innerHTML = `<div class="prev err">Não consegui gerar a imagem. Tente de novo.</div>`; }
  b.disabled = false; b.textContent = txt;
}

// ---------- eventos ----------
$("plFile").addEventListener("change", (e) => e.target.files.length && lerArquivos([...e.target.files]));
const dz = $("plDrop");
["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", (e) => e.dataTransfer.files.length && lerArquivos([...e.dataTransfer.files]));
$("plData").addEventListener("change", (e) => { selIso = e.target.value; render(); });
$("plFiltro").addEventListener("click", (e) => { const b = e.target.closest("[data-f]"); if (b) { filtro = b.dataset.f; renderPedidos(planos[selIso]); } });
$("plBusca").addEventListener("input", () => renderPedidos(planos[selIso]));
$("plEnviar").addEventListener("click", enviar);
$("t-plan").addEventListener("click", () => setTimeout(ajustar, 0));

render();
