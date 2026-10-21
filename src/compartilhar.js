// Entrega de arquivos, usada pelo painel do turno e pelo Planejamento VD.
// No celular abre o compartilhamento (WhatsApp); no computador baixa;
// dentro do Claude (pré-visualização) passa pela permissão de download do Claude.
import html2canvas from "html2canvas";

export const MIME = { png: "image/png", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", csv: "text/csv" };
const ehCelular = () => matchMedia("(pointer: coarse)").matches;

let claudeDl = null;
if (window.claude && typeof window.claude.use === "function") {
  window.claude.use("downloads").then((d) => { claudeDl = d; }).catch(() => {});
}

export function baixarArquivo(nome, blob) {
  if (claudeDl) return claudeDl.save({ filename: nome, data: blob }).catch(() => {});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// arquivos: [{nome, blob}] · devolve "compartilhado", "baixado" ou "cancelado"
export async function entregar(arquivos, titulo) {
  const files = arquivos.map((f) => new File([f.blob], f.nome, { type: MIME[f.nome.split(".").pop()] || f.blob.type }));
  if (ehCelular() && navigator.canShare && navigator.canShare({ files })) {
    try { await navigator.share({ files, title: titulo }); return "compartilhado"; }
    catch (e) { if (e && e.name === "AbortError") return "cancelado"; }
  }
  if (claudeDl) {
    for (const f of files) {
      try { await claudeDl.save({ filename: f.name, data: f }); }
      catch (e) { if (e && e.code === "declined") return "cancelado"; }
    }
    return "baixado";
  }
  files.forEach((f, i) => setTimeout(() => baixarArquivo(f.name, f), i * 400));
  return "baixado";
}

// gera o PNG de uma página (elemento com 800 px de largura) em 2x
export async function imagemDe(el) {
  const canvas = await html2canvas(el, {
    scale: 2, backgroundColor: "#ffffff", useCORS: true,
    onclone: (d) => {
      d.querySelectorAll(".pg-fit").forEach((f) => (f.style.transform = "none"));
      d.querySelectorAll(".pg-stage").forEach((f) => (f.style.height = "auto"));
    },
  });
  return await new Promise((r) => canvas.toBlob(r, "image/png"));
}
