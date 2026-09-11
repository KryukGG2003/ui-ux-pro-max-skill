/**
 * Генерирует градиентные фоны для слайдов.
 *
 * pptxgenjs не умеет градиентные заливки, поэтому фоны рендерятся
 * в Chromium и подставляются в слайды картинками.
 *
 * Запуск:  node make-backgrounds.mjs
 * Результат: assets/bg/*.jpg
 */
import { chromium } from "playwright";
import { mkdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, "assets", "bg");
await mkdir(outDir, { recursive: true });

/**
 * Каждый фон = база + световые пятна + СКРИМ + зерно.
 *
 * Скрим — затемняющий слой поверх пятен под текстовой зоной. Без него
 * яркий градиент опускает контраст белого текста ниже 4.5:1. Цвет при
 * этом остаётся по краям кадра, а середина становится читаемой.
 */
const VARIANTS = {
  // Титул и финал: оба поля светятся одинаково сильно
  cover: {
    base: "#0A0B1E",
    orbs: [
      { c: "#3D6BFF", x: "16%", y: "24%", r: "70%", a: 0.85 },
      { c: "#FF5E3A", x: "88%", y: "80%", r: "66%", a: 0.78 },
      { c: "#8B4BFF", x: "58%", y: "4%",  r: "52%", a: 0.50 },
    ],
    // на титуле текст слева — гасим только левую часть
    scrim: "linear-gradient(100deg, rgba(5,7,14,.90) 0%, rgba(5,7,14,.72) 34%, rgba(5,7,14,.20) 60%, rgba(5,7,14,0) 78%)",
  },
  // Объективная подструктура: холодная, синяя доминанта
  obj: {
    base: "#080C1E",
    orbs: [
      { c: "#3D7BFF", x: "10%", y: "10%", r: "68%", a: 0.78 },
      { c: "#22C8E8", x: "94%", y: "90%", r: "54%", a: 0.42 },
      { c: "#7A4BEA", x: "72%", y: "22%", r: "46%", a: 0.36 },
    ],
    scrim: "radial-gradient(ellipse 96% 86% at 44% 52%, rgba(5,7,14,.90) 0%, rgba(5,7,14,.74) 40%, rgba(5,7,14,.28) 68%, rgba(5,7,14,0) 88%)",
  },
  // Субъективная подструктура: тёплая, коралл и фиолет
  subj: {
    base: "#140A1C",
    orbs: [
      { c: "#FF6340", x: "12%", y: "84%", r: "68%", a: 0.78 },
      { c: "#9B45EE", x: "90%", y: "12%", r: "62%", a: 0.66 },
      { c: "#FFB03D", x: "62%", y: "96%", r: "46%", a: 0.44 },
    ],
    scrim: "radial-gradient(ellipse 96% 86% at 44% 52%, rgba(5,7,14,.91) 0%, rgba(5,7,14,.76) 40%, rgba(5,7,14,.30) 68%, rgba(5,7,14,0) 88%)",
  },
  // Нейтральные слайды: спокойный индиго
  neutral: {
    base: "#0A0C1C",
    orbs: [
      { c: "#4560E8", x: "86%", y: "14%", r: "62%", a: 0.60 },
      { c: "#E85B33", x: "8%",  y: "95%", r: "54%", a: 0.44 },
    ],
    scrim: "radial-gradient(ellipse 96% 86% at 44% 52%, rgba(5,7,14,.88) 0%, rgba(5,7,14,.70) 42%, rgba(5,7,14,.24) 70%, rgba(5,7,14,0) 88%)",
  },
};

function pageHtml(v) {
  const layers = v.orbs
    .map((o) => `radial-gradient(ellipse ${o.r} ${o.r} at ${o.x} ${o.y}, ${hexA(o.c, o.a)} 0%, ${hexA(o.c, o.a * 0.45)} 38%, ${hexA(o.c, 0)} 72%)`)
    .join(",");
  return `<!doctype html><meta charset="utf-8"><style>
  html,body{margin:0;height:100%;overflow:hidden}
  .bg{position:relative;width:1920px;height:1080px;background:${v.base}}
  .orbs{position:absolute;inset:0;background-image:${layers}}
  .scrim{position:absolute;inset:0;background-image:${v.scrim || "none"}}
  /* лёгкое зерно: убирает бандинг на плавных тёмных градиентах */
  .grain{position:absolute;inset:0;opacity:.05;
    background-image:radial-gradient(#fff 0.5px, transparent 0.6px);
    background-size:3px 3px}
  </style><div class="bg"><div class="orbs"></div><div class="scrim"></div><div class="grain"></div></div>`;
}

function hexA(hex, a) {
  const h = hex.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${a})`;
}

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });

for (const [name, v] of Object.entries(VARIANTS)) {
  await page.setContent(pageHtml(v), { waitUntil: "load" });
  await page.waitForTimeout(120);
  const file = join(outDir, `${name}.jpg`);
  await page.locator(".bg").screenshot({ path: file, type: "jpeg", quality: 92 });
  console.log("создан", file);
}

await browser.close();
