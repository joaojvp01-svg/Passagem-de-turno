// Sincroniza o que o site guarda no navegador (localStorage) com o banco compartilhado (/api/dados).
// O navegador continua sendo a cópia de trabalho: o site abre e funciona mesmo sem internet,
// e o que foi lançado sobe assim que a conexão volta.
// Os dados de cada dia são gravados por turno, para dois turnos não sobrescreverem um ao outro.

const API = "api/dados";
const KMETA = "passagemTurno.nuvem", KCOD = "passagemTurno.codigo";
const DIA = /^passagemTurno\.dia\.(\d{4}-\d{2}-\d{2})$/;
const SIMPLES = {
  "passagemTurno.equipe": "equipe", "passagemTurno.logins": "logins", "passagemTurno.afast": "afast",
  "planejamentoVD.planos": "plan.planos", "planejamentoVD.query": "plan.query",
  "planejamentoVD.historico": "plan.historico", "planejamentoVD.faturamento": "plan.faturamento",
};
const GLOBAIS = Object.values(SIMPLES);
const localDe = Object.fromEntries(Object.entries(SIMPLES).map(([l, r]) => [r, l]));

const setOrig = Storage.prototype.setItem;
const getL = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
const setL = (k, v) => { try { setOrig.call(localStorage, k, v); } catch (e) {} };

let meta = {}; try { meta = JSON.parse(getL(KMETA) || "{}") || {}; } catch (e) {}
const salvarMeta = () => setL(KMETA, JSON.stringify(meta));
const hash = (s) => { if (s == null) return ""; let h = 0x811c9dc5; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36) + ":" + s.length; };

// estado: "?" testando, "on" sincronizando, "local" sem banco, "off" sem conexão
const puxados = new Set(); // só envia o que já foi comparado com o banco nesta sessão
let modo = "?", ouvintes = [], aoStatus = () => {}, pendentes = new Set(), timer = null, ultimo = null;
let txtAtual = "Conectando…";
const status = (m, txt) => { modo = m; txtAtual = txt; aoStatus(m, txt); };

// ---- valor local de uma chave remota
function lerLocal(rk) {
  const d = rk.match(/^dia\.(\d{4}-\d{2}-\d{2})\.([123])$/);
  if (d) { try { const o = JSON.parse(getL("passagemTurno.dia." + d[1]) || "null"); return o && o[d[2]] ? JSON.stringify(o[d[2]]) : null; } catch (e) { return null; } }
  return getL(localDe[rk]);
}
function gravarLocal(rk, v) {
  const d = rk.match(/^dia\.(\d{4}-\d{2}-\d{2})\.([123])$/);
  if (d) { const k = "passagemTurno.dia." + d[1]; let o = {}; try { o = JSON.parse(getL(k) || "{}") || {}; } catch (e) {} o[d[2]] = v; setL(k, JSON.stringify(o)); return; }
  setL(localDe[rk], JSON.stringify(v));
}
const remotasDe = (k) => { const d = k.match(DIA); if (d) return [1, 2, 3].map((t) => `dia.${d[1]}.${t}`); return SIMPLES[k] ? [SIMPLES[k]] : []; };

// ---- intercepta as gravações do site
Storage.prototype.setItem = function (k, v) {
  setOrig.call(this, k, v);
  if (this === localStorage) { const rs = remotasDe(k); if (rs.length) { rs.forEach((r) => pendentes.add(r)); agendar(); } }
};
function agendar(ms = 900) { clearTimeout(timer); timer = setTimeout(enviarPendentes, ms); }

async function chamar(url, opt = {}) {
  const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), 15000);
  try {
    const r = await fetch(url, { ...opt, signal: ctrl.signal, headers: { "Content-Type": "application/json", "x-codigo": getL(KCOD) || "", ...(opt.headers || {}) } });
    if (r.status === 401) {
      const c = window.prompt("Digite o código de acesso do site da passagem de turno:");
      if (c) { setL(KCOD, c.trim()); return chamar(url, opt); }
      throw Object.assign(new Error("codigo"), { semBanco: true });
    }
    if (r.status === 503 || r.status === 404 || !(r.headers.get("content-type") || "").includes("json")) throw Object.assign(new Error("sem banco"), { semBanco: true });
    const j = await r.json(); if (!r.ok || !j.ok) throw new Error(j.erro || r.status); return j;
  } finally { clearTimeout(to); }
}
const falha = (e) => status(e && e.semBanco ? "local" : "off", e && e.semBanco ? "Salvo só neste navegador" : "Sem conexão: salvo neste aparelho, envia quando voltar");

async function enviarPendentes() {
  if (modo === "local" || !pendentes.size) return;
  const lista = [...pendentes].filter((rk) => puxados.has(rk)); lista.forEach((rk) => pendentes.delete(rk));
  if (!lista.length) return;
  for (const rk of lista) {
    const s = lerLocal(rk), h = hash(s);
    if (s == null || (meta[rk] && meta[rk].h === h)) continue;
    try {
      const j = await chamar(API, { method: "POST", body: JSON.stringify({ k: rk, v: JSON.parse(s) }) });
      meta[rk] = { t: j.t, h }; salvarMeta(); ultimo = new Date();
    } catch (e) { lista.filter((x) => !meta[x] || meta[x].h !== hash(lerLocal(x))).forEach((x) => pendentes.add(x)); falha(e); return; }
  }
  status("on", "Salvo para toda a equipe" + (ultimo ? ` · ${ultimo.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}` : ""));
}

// ---- busca o que mudou em outros aparelhos
export async function sincronizar(dia) {
  if (modo === "local") return;
  if (location.protocol === "file:") { status("local", "Salvo só neste navegador"); return; }
  const rks = [...GLOBAIS, ...(dia ? [1, 2, 3].map((t) => `dia.${dia}.${t}`) : [])];
  let j; try { j = await chamar(`${API}?k=${rks.join(",")}`); } catch (e) { falha(e); return; }
  const mudou = [];
  for (const rk of rks) {
    puxados.add(rk);
    const r = j.dados[rk], s = lerLocal(rk), h = hash(s), m = meta[rk];
    if (!r) { if (s != null && (!m || m.h !== h)) pendentes.add(rk); continue; }
    if (!m || r.t > m.t) {
      const alterouAqui = s != null && m && m.h !== h; // mudou neste aparelho desde a última sincronização
      if (alterouAqui) { pendentes.add(rk); continue; }
      gravarLocal(rk, r.v); meta[rk] = { t: r.t, h: hash(lerLocal(rk)) }; mudou.push(rk);
    } else if (m.h !== h) pendentes.add(rk);
  }
  salvarMeta();
  if (mudou.length) ouvintes.forEach((f) => { try { f(mudou); } catch (e) { console.error(e); } });
  if (pendentes.size) await enviarPendentes(); else status("on", "Salvo para toda a equipe" + (ultimo ? ` · ${ultimo.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}` : ""));
}
export const aoMudarRemoto = (f) => ouvintes.push(f);
export const aoMudarStatus = (f) => { aoStatus = f; f(modo, txtAtual); };
export const modoAtual = () => modo;

// confere de novo quando a pessoa volta para a aba e a cada minuto com a aba aberta
let diaAtual = () => null;
export const definirDia = (f) => { diaAtual = f; };
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") sincronizar(diaAtual()); else enviarPendentes(); });
setInterval(() => { if (document.visibilityState === "visible" && modo !== "local") sincronizar(diaAtual()); }, 60000);
window.addEventListener("online", () => sincronizar(diaAtual()));
