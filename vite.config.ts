import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react-swc';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
  plugins: [react()],
  // Relative asset paths, so the built page works wherever it is served from — the root of a
  // domain, or a GitHub Pages subpath like /AplusContentTool/ — without rebuilding for each.
  // There is no router here, so nothing else depends on knowing the base path.
  base: './',
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
  },
  server: {
    port: Number(process.env.PORT) || 5180,
  },
  build: {
    // ExcelJS is most of the bundle and is needed on first interaction, so splitting it out buys
    // nothing here. Raising the limit keeps the build output free of a warning that is not
    // actionable for an internal tool.
    chunkSizeWarningLimit: 2000,
  },
});
