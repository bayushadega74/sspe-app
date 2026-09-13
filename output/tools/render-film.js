// Render film.html menjadi MP4 (1080x1920, 25 fps).
// Jalankan dari folder repo: node output/tools/render-film.js
//
// Butuh: Node.js + Playwright (npm i playwright), dan ffmpeg.
//   - ffmpeg dari PATH, atau set FFMPEG=/path/ke/ffmpeg
//   - jika tak ada ffmpeg: pip install imageio-ffmpeg, lalu
//       FFMPEG=$(python3 -c "import imageio_ffmpeg;print(imageio_ffmpeg.get_ffmpeg_exe())") node output/tools/render-film.js
// Opsional: CHROMIUM_PATH untuk executable Chromium tertentu.
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const FPS = 25, DUR = 52.0, N = Math.round(FPS * DUR);
const dir = path.resolve(__dirname, '..', 'film-orangtua');
const file = path.join(dir, 'film.html');
const framesDir = path.join(dir, '_frames');
const outMp4 = path.join(dir, 'film-keselamatan-anak-twentyswim.mp4');

(async () => {
  fs.rmSync(framesDir, { recursive: true, force: true });
  fs.mkdirSync(framesDir, { recursive: true });
  const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {};
  const browser = await chromium.launch(launch);
  const page = await browser.newPage({ viewport: { width: 1080, height: 1920 }, deviceScaleFactor: 1 });
  await page.goto('file://' + file + '?capture=1', { waitUntil: 'load' });
  await page.waitForFunction(() => window.__ready === true, { timeout: 10000 });
  const clip = { x: 0, y: 0, width: 1080, height: 1920 };
  for (let i = 0; i < N; i++) {
    await page.evaluate(ms => window.__seek(ms), Math.round(i * 1000 / FPS));
    await page.screenshot({ path: path.join(framesDir, 'f' + String(i).padStart(5, '0') + '.png'), clip });
    if (i % 100 === 0) process.stdout.write(i + '..');
  }
  await browser.close();
  console.log('\n' + N + ' frame selesai. Encoding MP4...');

  const ff = process.env.FFMPEG || 'ffmpeg';
  execFileSync(ff, ['-y', '-hide_banner', '-loglevel', 'error',
    '-framerate', String(FPS), '-i', path.join(framesDir, 'f%05d.png'),
    '-c:v', 'libx264', '-preset', 'medium', '-crf', '22',
    '-pix_fmt', 'yuv420p', '-profile:v', 'high', '-level', '4.0',
    '-movflags', '+faststart', '-r', String(FPS), outMp4], { stdio: 'inherit' });
  fs.rmSync(framesDir, { recursive: true, force: true });
  console.log('Selesai:', outMp4);
})().catch(e => { console.error(e); process.exit(1); });
