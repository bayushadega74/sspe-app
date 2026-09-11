// Ekspor slide-1..6.html menjadi PNG 1080x1350.
// Jalankan dari folder repo: node output/tools/render-slides.js
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

(async () => {
  const dir = path.resolve(__dirname, '..', 'carousel-ig-orangtua');
  const outDir = path.join(dir, 'png');
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1350 }, deviceScaleFactor: 1 });
  for (let i = 1; i <= 6; i++) {
    await page.goto('file://' + path.join(dir, `slide-${i}.html`), { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts.ready);
    await page.screenshot({ path: path.join(outDir, `slide-${i}.png`), clip: { x: 0, y: 0, width: 1080, height: 1350 } });
    console.log(`slide-${i}.png`);
  }
  await browser.close();
})();
