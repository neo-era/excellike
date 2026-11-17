import { defineConfig } from 'vite';
// --mode demo: gộp toàn bộ mã (kể cả phần .xlsx) vào một tệp để chạy thử dạng một trang HTML
// chính sách bảo mật nội dung (bản build web/PWA; bản demo gộp mã nội tuyến nên không gắn)
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; manifest-src 'self'; object-src 'none'; base-uri 'none'";
const csp = { name: 'csp', apply: 'build' as const, transformIndexHtml: (html: string) => html.replace('<meta charset="utf-8">', `<meta charset="utf-8">\n  <meta http-equiv="Content-Security-Policy" content="${CSP}" />`) };
// số → chữ (toán tử &, CONCATENATE, LEFT…): như Excel tiếng Việt — tối đa 15 chữ số có nghĩa, dấu thập phân ','
// (HyperFormula dùng Number.toString: 0,55*100 → "55.00000000000001"). Hàm đổi đặt ở globalThis.__xlNumText (workbook.ts)
let hfPatched = 0, hfBuild = false;
const hfNumText = {
  name: 'hf-num-text',
  enforce: 'pre' as const,
  configResolved(c: { command: string }) { hfBuild = c.command === 'build'; },
  buildStart() { hfPatched = 0; },
  // bản build mà không sửa được tệp nào (HyperFormula đổi đường dẫn) → dừng, không cho ra bản sai âm thầm
  buildEnd(err?: Error) { if (!err && hfBuild && hfPatched === 0) throw new Error('hf-num-text: chưa sửa được HyperFormula (ArithmeticHelper) — kiểm tra lại phiên bản'); },
  transform(code: string, id: string) {
    if (!/hyperformula[\\/](es|commonjs)[\\/]interpreter[\\/]ArithmeticHelper\.m?js$/.test(id)) return null;
    const re = /return (\(0, _InterpreterValue\.getRawValue\)|getRawValue)\(arg\)\.toString\(\);/g;
    const n = (code.match(re) ?? []).length;
    if (n !== 1) throw new Error('hf-num-text: không tìm thấy chỗ đổi số → chữ trong HyperFormula (đổi phiên bản?)');
    hfPatched++;
    return { code: code.replace(re, (_m, g) => `return (globalThis.__xlNumText ?? String)(${g}(arg));`), map: null };
  },
};
export default defineConfig(({ mode }) => ({
  base: './',
  // dev: không gói sẵn HyperFormula bằng esbuild để plugin hf-num-text vẫn áp dụng
  optimizeDeps: { exclude: ['hyperformula'] },
  plugins: mode === 'demo' ? [hfNumText] : [hfNumText, csp],
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 4000,
    outDir: mode === 'demo' ? 'dist-demo' : 'dist',
    rollupOptions: mode === 'demo' ? { output: { inlineDynamicImports: true } } : {},
  },
  test: { globals: true, environment: 'node', include: ['tests/**/*.test.ts'], server: { deps: { inline: [/hyperformula/] } } },
}) as any);
