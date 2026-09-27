// Draws a simple 256x256 document icon (page with folded corner, text lines, and a grid).
import fs from "node:fs";
import { PNG } from "pngjs";
const S = 256, png = new PNG({ width: S, height: S });
const set = (x, y, [r, g, b, a = 255]) => { const i = (y * S + x) * 4; png.data[i] = r; png.data[i + 1] = g; png.data[i + 2] = b; png.data[i + 3] = a; };
const rect = (x0, y0, x1, y1, c) => { for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) set(x, y, c); };
png.data.fill(0);
// rounded background
const bg = [37, 99, 235];
for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
  const r = 48, cx = Math.min(Math.max(x, r), S - 1 - r), cy = Math.min(Math.max(y, r), S - 1 - r);
  if ((x - cx) ** 2 + (y - cy) ** 2 <= r * r) set(x, y, bg);
}
// page with folded corner
const L = 64, T = 36, R = 192, B = 220, F = 40;
for (let y = T; y < B; y++) for (let x = L; x < R; x++) if (!(x > R - F && y < T + F && (x - (R - F)) > (y - T))) set(x, y, [255, 255, 255]);
for (let y = T; y < T + F; y++) for (let x = R - F; x < R; x++) if (x - (R - F) <= y - T) set(x, y, [191, 208, 250]);
// text lines
for (const [y, w] of [[92, 96], [110, 80], [128, 96]]) rect(80, y, 80 + w, y + 8, [100, 116, 139]);
// grid (spreadsheet)
const g = [22, 163, 74];
for (let x = 80; x <= 176; x += 32) rect(x, 152, x + 3, 204, g);
for (let y = 152; y <= 204; y += 17) rect(80, y, 179, y + 3, g);
// Upscale 2x to the recommended 512x512.
const out = new PNG({ width: S * 2, height: S * 2 });
for (let y = 0; y < S * 2; y++) for (let x = 0; x < S * 2; x++) png.data.copy(out.data, (y * S * 2 + x) * 4, ((y >> 1) * S + (x >> 1)) * 4, ((y >> 1) * S + (x >> 1)) * 4 + 4);
fs.writeFileSync(new URL("../icon.png", import.meta.url), PNG.sync.write(out));
