// Gộp bản build thành một trang HTML duy nhất để chạy thử (không cần máy chủ)
import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
const dist = new URL('../dist-demo/', import.meta.url).pathname;
const assets = readdirSync(dist + 'assets');
const js = readFileSync(dist + 'assets/' + assets.find((f) => f.endsWith('.js')), 'utf8').replace(/<\/script/gi, '<\\/script').replace(/\uFFFD/g, '\\uFFFD');
const css = readFileSync(dist + 'assets/' + assets.find((f) => f.endsWith('.css')), 'utf8');
const html = readFileSync(dist + 'index.html', 'utf8');
const body = /<body>([\s\S]*?)<\/body>/.exec(html)[1].replace(/<script[^>]*src=[^>]*><\/script>/g, '');
const icon = 'data:image/png;base64,' + readFileSync(dist + 'icons/icon-192.png').toString('base64');
const out = `<title>ExcelLike</title>
<meta name="description" content="Bản chạy thử ExcelLike – Đợt 1">
<link rel="icon" href="${icon}">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap">
<style>${css}
html, body { height: 100%; }
.app { height: 100%; padding-top: 0; }
</style>
${body}
<script type="module">${js}</script>
`;
writeFileSync(new URL('../dist-demo/excellike-demo.html', import.meta.url), out);
console.log('demo', (out.length / 1024).toFixed(0), 'KB');
