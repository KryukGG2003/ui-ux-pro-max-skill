/**
 * Собирает standalone-страницу landing.html из tilda-embed.html.
 *
 * tilda-embed.html — источник правды: его же содержимое вставляется
 * в блок T123 в Tilda. Здесь оно просто оборачивается в полноценный
 * HTML-документ, чтобы страницу можно было открыть локально или
 * залить на любой хостинг.
 *
 * Запуск:  node build-landing.mjs
 */
import { readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const embed = await readFile(join(here, "tilda-embed.html"), "utf8");

const page = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Структура конфликта — объективная и субъективная подструктуры</title>
<meta name="description" content="Объективная и субъективная структура конфликта по А. Я. Анцупову и А. И. Шипилову: участники, объект, предмет, микро- и макросреда, мотивы, цели, образы и прогноз результатов.">
<meta property="og:title" content="Структура конфликта">
<meta property="og:description" content="Объективная и субъективная подструктуры конфликта — разбор по элементам с примером.">
<meta property="og:type" content="website">
<style>
  html{scroll-behavior:smooth}
  body{margin:0;background:#0B0E14}
  @media (prefers-reduced-motion:reduce){html{scroll-behavior:auto}}
</style>
</head>
<body>
<!-- ВНИМАНИЕ: файл создаётся автоматически из tilda-embed.html — правьте источник, затем: node build-landing.mjs -->
${embed.trim()}
</body>
</html>
`;

await writeFile(join(here, "landing.html"), page, "utf8");
console.log("landing.html собран из tilda-embed.html");
