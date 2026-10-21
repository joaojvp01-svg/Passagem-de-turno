import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";

// npm run build          -> site normal (Vercel), pasta dist
// npm run build:arquivo  -> um único index.html com tudo dentro, pasta dist-arquivo
//                           (abre com clique duplo, sem servidor)
export default defineConfig(({ mode }) => {
  if (mode === "arquivo") {
    return {
      plugins: [viteSingleFile()],
      build: { outDir: "dist-arquivo", assetsInlineLimit: 100000000, chunkSizeWarningLimit: 5000 },
    };
  }
  return {
    build: {
      outDir: "dist",
      // html2canvas e xlsx são grandes; separar ajuda o cache do navegador
      rollupOptions: { output: { manualChunks: { html2canvas: ["html2canvas"], xlsx: ["xlsx"] } } },
      chunkSizeWarningLimit: 1200,
    },
  };
});
