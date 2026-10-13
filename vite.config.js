import { defineConfig } from "vite";

export default defineConfig({
  build: {
    outDir: "dist",
    // html2canvas e xlsx são grandes; separar ajuda o cache do navegador
    rollupOptions: {
      output: {
        manualChunks: { html2canvas: ["html2canvas"], xlsx: ["xlsx"] },
      },
    },
    chunkSizeWarningLimit: 1200,
  },
});
