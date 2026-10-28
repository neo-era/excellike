import { defineConfig } from 'vite';
// --mode demo: gộp toàn bộ mã (kể cả phần .xlsx) vào một tệp để chạy thử dạng một trang HTML
// chính sách bảo mật nội dung (bản build web/PWA; bản demo gộp mã nội tuyến nên không gắn)
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; manifest-src 'self'; object-src 'none'; base-uri 'none'";
const csp = { name: 'csp', apply: 'build' as const, transformIndexHtml: (html: string) => html.replace('<meta charset="utf-8">', `<meta charset="utf-8">\n  <meta http-equiv="Content-Security-Policy" content="${CSP}" />`) };
export default defineConfig(({ mode }) => ({
  base: './',
  plugins: mode === 'demo' ? [] : [csp],
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 4000,
    outDir: mode === 'demo' ? 'dist-demo' : 'dist',
    rollupOptions: mode === 'demo' ? { output: { inlineDynamicImports: true } } : {},
  },
  test: { globals: true, environment: 'node', include: ['tests/**/*.test.ts'] },
}) as any);
