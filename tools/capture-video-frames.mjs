import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');

const videoPath =
  process.argv[2] ??
  'C:\\Users\\cleyt\\Downloads\\WhatsApp Video 2026-04-13 at 18.52.12.mp4';
const outputDir = process.argv[3] ?? 'video-frames';

const videoUrl = pathToFileURL(videoPath).href;

const executablePath =
  process.env.BROWSER_PATH ??
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const headless = process.env.HEADLESS !== 'false';

const browserWithMode = await chromium.launch({ executablePath, headless });
const pageWithMode = await browserWithMode.newPage({
  viewport: { width: 420, height: 900 },
  deviceScaleFactor: 1,
});

await pageWithMode.goto(videoUrl, { waitUntil: 'domcontentloaded' });

const meta = await pageWithMode.evaluate(
  () =>
    new Promise((resolve, reject) => {
      const video = document.querySelector('video');
      if (!video) {
        reject(new Error('No video element found'));
        return;
      }
      const done = () =>
        resolve({
          duration: video.duration,
          width: video.videoWidth,
          height: video.videoHeight,
        });

      if (video.readyState >= 1) {
        done();
        return;
      }

      const timeout = window.setTimeout(
        () => reject(new Error(`Timed out loading ${video.currentSrc}`)),
        15000,
      );
      video.onloadedmetadata = () => {
        window.clearTimeout(timeout);
        done();
      };
      video.onerror = () =>
        reject(
          new Error(
            `Unable to load video: ${video.error?.code ?? 'unknown'} ${video.error?.message ?? ''}`,
          ),
        );
      video.load();
    }),
);

console.log(JSON.stringify(meta));

const times = [
  0.1,
  meta.duration * 0.16,
  meta.duration * 0.32,
  meta.duration * 0.48,
  meta.duration * 0.64,
  meta.duration * 0.8,
  Math.max(0.1, meta.duration - 0.5),
];

for (const [index, time] of times.entries()) {
  await pageWithMode.evaluate(
    (seekTime) =>
      new Promise((resolve, reject) => {
        const video = document.querySelector('video');
        video.onseeked = () => resolve();
        video.onerror = () => reject(new Error('Unable to seek video'));
        video.currentTime = seekTime;
      }),
    time,
  );

  const filename = `frame-${String(index + 1).padStart(2, '0')}.png`;
  await pageWithMode.screenshot({ path: path.resolve(outputDir, filename), fullPage: true });
  console.log(`${filename} ${time.toFixed(2)}s`);
}

await browserWithMode.close();
