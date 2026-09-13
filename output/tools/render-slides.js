// Ekspor slide-1..6.html menjadi PNG 1080x1350.
// Jalankan dari folder repo: node output/tools/render-slides.js
//
// Variabel lingkungan opsional:
//   CHROMIUM_PATH   path executable Chromium bila Playwright tidak mengunduh browser sendiri.
//   LOCAL_FONTS_DIR folder berisi Montserrat-700.woff2 dan Lora-400.woff2, dipakai bila
//                   mesin tidak bisa mengakses Google Fonts (font disuntik sebagai data URI).
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

(async () => {
  const dir = path.resolve(__dirname, '..', 'carousel-ig-orangtua');
  const outDir = path.join(dir, 'png');
  fs.mkdirSync(outDir, { recursive: true });
  const launchOpts = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {};
  const browser = await chromium.launch(launchOpts);
  const page = await browser.newPage({ viewport: { width: 1080, height: 1350 }, deviceScaleFactor: 1 });

  let fontCss = null;
  if (process.env.LOCAL_FONTS_DIR) {
    const b64 = f => fs.readFileSync(path.join(process.env.LOCAL_FONTS_DIR, f)).toString('base64');
    fontCss = `@font-face{font-family:"Montserrat";font-weight:100 900;src:url(data:font/woff2;base64,${b64('Montserrat-700.woff2')}) format("woff2")}
               @font-face{font-family:"Lora";font-weight:100 900;src:url(data:font/woff2;base64,${b64('Lora-400.woff2')}) format("woff2")}`;
  }

  for (let i = 1; i <= 6; i++) {
    await page.goto('file://' + path.join(dir, `slide-${i}.html`), { waitUntil: 'networkidle' });
    if (fontCss) await page.addStyleTag({ content: fontCss });
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(300);
    const ok = await page.evaluate(() => document.fonts.check('800 20px Montserrat') && document.fonts.check('500 20px Lora'));
    await page.screenshot({ path: path.join(outDir, `slide-${i}.png`), clip: { x: 0, y: 0, width: 1080, height: 1350 } });
    console.log(`slide-${i}.png (font Montserrat/Lora ${ok ? 'termuat' : 'TIDAK termuat, memakai fallback'})`);
  }
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
