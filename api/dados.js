// Armazenamento compartilhado do site (Vercel Function + Upstash Redis).
// Sem isso, cada navegador guarda só os próprios dados.
//
// Variáveis de ambiente (criadas pela integração Upstash da Vercel):
//   KV_REST_API_URL + KV_REST_API_TOKEN   ou   UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN
// Opcional:
//   CODIGO_ACESSO  -> se definido, o site pede esse código uma vez em cada aparelho.
//
// GET  /api/dados?k=a,b,c   -> { ok:true, dados:{ a:{t,v}|null, ... } }
// POST /api/dados  {k, v, t} -> grava (v null apaga)

const URL_ = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
const PREFIXO = "passagem:";
const CHAVE_OK = /^(dia\.\d{4}-\d{2}-\d{2}\.[123]|equipe|logins|afast|plan\.(planos|query|historico|faturamento))$/;

async function redis(cmds) {
  const r = await fetch(`${URL_}/pipeline`, {
    method: "POST",
    headers: { Authorization: `Bearer ${TOKEN}`, "Content-Type": "application/json" },
    body: JSON.stringify(cmds),
  });
  if (!r.ok) throw new Error(`Upstash ${r.status}`);
  return r.json();
}

async function corpo(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") return JSON.parse(req.body || "{}");
  const partes = []; for await (const c of req) partes.push(c);
  return JSON.parse(Buffer.concat(partes).toString("utf8") || "{}");
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  if (!URL_ || !TOKEN) return res.status(503).json({ ok: false, erro: "banco-nao-configurado" });
  const codigo = process.env.CODIGO_ACESSO;
  if (codigo && req.headers["x-codigo"] !== codigo) return res.status(401).json({ ok: false, erro: "codigo" });
  try {
    if (req.method === "GET") {
      const ks = String(req.query.k || "").split(",").filter((k) => CHAVE_OK.test(k)).slice(0, 50);
      if (!ks.length) return res.status(200).json({ ok: true, dados: {} });
      const out = await redis(ks.map((k) => ["GET", PREFIXO + k]));
      const dados = {};
      ks.forEach((k, i) => { const v = out[i] && out[i].result; dados[k] = v ? JSON.parse(v) : null; });
      return res.status(200).json({ ok: true, dados });
    }
    if (req.method === "POST") {
      const b = await corpo(req);
      if (!CHAVE_OK.test(b.k || "")) return res.status(400).json({ ok: false, erro: "chave" });
      const t = Date.now();
      if (b.v == null) await redis([["DEL", PREFIXO + b.k]]);
      else await redis([["SET", PREFIXO + b.k, JSON.stringify({ t, v: b.v })]]);
      return res.status(200).json({ ok: true, t });
    }
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ ok: false });
  } catch (e) {
    return res.status(500).json({ ok: false, erro: String(e.message || e) });
  }
}
