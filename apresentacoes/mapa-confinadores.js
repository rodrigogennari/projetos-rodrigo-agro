// Gera o mapa do Brasil por estado (PNG) a partir de @svg-maps/brazil (CC BY 4.0)
const sharp = require("sharp");
const fs = require("fs");
const src = fs.readFileSync(__dirname + "/brazil-map.mjs", "utf8").replace("export default ", "");
const MAP = JSON.parse(src.trim().replace(/;\s*$/, ""));
const D = { ac: 78, al: 3, am: 34, ba: 117, df: 4, es: 13, go: 93, ma: 1, mg: 200, ms: 161, mt: 341, pa: 77, pe: 3, pi: 3, pr: 94, rj: 8, rn: 11, ro: 1022, rr: 8, rs: 8, sc: 18, se: 5, sp: 461, to: 34 };
const cls = (v) => !v ? ["EEF0F3", "7A858F"] : v <= 20 ? ["DCE7F3", "4A5560"] : v <= 100 ? ["9DBBDF", "12295F"] : v <= 300 ? ["4A79BF", "FFFFFF"] : ["1B3A8C", "FFFFFF"];
function centroid(path) {
  const t = path.trim().split(/\s+/); let x = 0, y = 0, sx = 0, sy = 0, cmd = "m", polys = [], cur = null;
  for (const tok of t) {
    if (tok === "m" || tok === "z") { if (tok === "z") { x = sx; y = sy; } cmd = tok; continue; }
    const [dx, dy] = tok.split(",").map(Number);
    if (cmd === "m") { x += dx; y += dy; sx = x; sy = y; cur = [[x, y]]; polys.push(cur); cmd = "l"; } else { x += dx; y += dy; cur.push([x, y]); }
  }
  let best = null;
  for (const p of polys) {
    let a = 0, cx = 0, cy = 0;
    for (let i = 0; i < p.length; i++) { const [x1, y1] = p[i], [x2, y2] = p[(i + 1) % p.length]; const f = x1 * y2 - x2 * y1; a += f; cx += (x1 + x2) * f; cy += (y1 + y2) * f; }
    a /= 2; if (!best || Math.abs(a) > Math.abs(best.a)) best = { a, x: cx / (6 * a), y: cy / (6 * a) };
  }
  return best;
}
const NUDGE = { go: [-6, 8], mg: [8, 0], ba: [6, 0], pa: [0, 10], am: [0, 10], sp: [6, 2], pr: [0, 0], ms: [0, 0], to: [0, 0], ac: [0, 2] };
const LABEL = ["rs", "rr", "ma", "pi", "ac", "am", "pa", "ro", "mt", "to", "ba", "go", "mg", "ms", "sp", "pr", "sc"];
let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${MAP.viewBox}" width="1226" height="1278">`;
for (const l of MAP.locations) svg += `<path d="${l.path}" fill="#${cls(D[l.id])[0]}" stroke="#FFFFFF" stroke-width="1.2"/>`;
for (const l of MAP.locations) {
  if (!LABEL.includes(l.id)) continue;
  const c = centroid(l.path), n = NUDGE[l.id] || [0, 0], col = cls(D[l.id])[1], big = D[l.id] >= 300;
  const x = c.x + n[0], y = c.y + n[1];
  svg += `<text x="${x}" y="${y - 2}" font-family="Carlito" font-size="${big ? 15 : 12}" font-weight="bold" fill="#${col}" text-anchor="middle">${l.id.toUpperCase()}</text>`;
  svg += `<text x="${x}" y="${y + (big ? 15 : 12)}" font-family="Carlito" font-size="${big ? 15 : 11.5}" fill="#${col}" text-anchor="middle">${D[l.id].toLocaleString("pt-BR")}</text>`;
}
svg += "</svg>";
sharp(Buffer.from(svg)).png().toFile(__dirname + "/mapa-brasil.png").then(() => console.log("ok", Object.values(D).reduce((a, b) => a + b, 0)));
