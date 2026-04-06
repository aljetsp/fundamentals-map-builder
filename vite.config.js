import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Replace 'fundamentals-map-builder' with your actual GitHub repo name
// e.g. if your repo is github.com/yourname/my-map-tool, set base: '/my-map-tool/'
export default defineConfig({
  base: '/fundamentals-map-builder/',
  plugins: [react()],
});
