/**
 * Собирает презентацию «Объективная и субъективная структура конфликта» в .pptx.
 *
 * Запуск:
 *   npm install pptxgenjs
 *   node build-pptx.mjs
 *
 * Результат: struktura-konflikta.pptx — открывается в PowerPoint,
 * Google Презентациях, Keynote и LibreOffice.
 */
import PptxGenJS from "pptxgenjs";

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";          // 13.333 × 7.5 дюйма — ДО добавления слайдов
pres.author = "Кудякова Дарья, Бороздова Валерия";
pres.title = "Объективная и субъективная структура конфликта";
pres.subject = "Конфликтология. Структура конфликта по А. Я. Анцупову и А. И. Шипилову";

/* ---------------------------------------------------------------
   ПАЛИТРА
   Холодный синий = объективное (внешние факты).
   Тёплый коралловый = субъективное (внутренние образы).
   --------------------------------------------------------------- */
const C = {
  bg:       "0B0E14",
  card:     "161C2A",
  cardLine: "2A3547",
  text:     "EAEFF7",
  mute:     "9FADC4",
  faint:    "8494AE",
  obj:      "4DA3FF",
  subj:     "FF8A6B",
  objCard:  "12233A",
  objLine:  "2A4E7A",
  subjCard: "2B1C17",
  subjLine: "6E3A2A",
};

const F = { head: "Cambria", body: "Calibri" };

/* Геометрия полосы набора */
const W = 13.333, H = 7.5;
const M = 0.72;                 // боковые поля
const CW = W - M * 2;           // ширина контента = 11.893

let slideNo = 0;

/* ---------------------------------------------------------------
   ХЕЛПЕРЫ
   --------------------------------------------------------------- */
function newSlide(notes) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideNo++;
  if (notes) s.addNotes(notes);
  return s;
}

function eyebrow(s, txt, tone = "obj") {
  s.addShape(pres.ShapeType.ellipse, {
    x: M, y: 0.62, w: 0.1, h: 0.1, fill: { color: tone === "subj" ? C.subj : C.obj },
  });
  s.addText(txt.toUpperCase(), {
    x: M + 0.2, y: 0.52, w: CW - 0.2, h: 0.3,
    fontFace: F.body, fontSize: 10.5, bold: true, color: C.faint,
    charSpacing: 2.2, isTextBox: true, margin: 0, valign: "middle",
  });
}

function title(s, txt, size = 33) {
  s.addText(txt, {
    x: M, y: 0.92, w: CW, h: 0.9,
    fontFace: F.head, fontSize: size, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
}

function lead(s, runs, y = 1.92, h = 0.62, size = 13.5) {
  s.addText(runs, {
    x: M, y, w: CW, h,
    fontFace: F.body, fontSize: size, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
}

/** Карточка с заливкой по тону */
function card(s, { x, y, w, h, tone = "plain" }) {
  const fill = tone === "obj" ? C.objCard : tone === "subj" ? C.subjCard : C.card;
  const line = tone === "obj" ? C.objLine : tone === "subj" ? C.subjLine : C.cardLine;
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: fill },
    line: { color: line, width: 1 },
  });
}

/** Кружок с номером или буквой */
function badge(s, { x, y, label, tone = "obj", d = 0.4 }) {
  const col = tone === "subj" ? C.subj : C.obj;
  const fillCol = tone === "subj" ? C.subjCard : C.objCard;
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d,
    fill: { color: fillCol },
    line: { color: col, width: 1 },
  });
  s.addText(String(label), {
    x, y, w: d, h: d,
    fontFace: F.head, fontSize: 12, bold: true, color: col,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
}

/** Заголовок + описание внутри карточки */
function cardText(s, { x, y, w, head, body, tone = "plain", headSize = 14.5, bodySize = 11 }) {
  const headCol = tone === "obj" ? C.obj : tone === "subj" ? C.subj : C.text;
  s.addText(head, {
    x, y, w, h: 0.5,
    fontFace: F.body, fontSize: headSize, bold: true, color: headCol,
    isTextBox: true, margin: 0, valign: "top",
  });
  if (body) {
    s.addText(body, {
      x, y: y + 0.46, w, h: 1.6,
      fontFace: F.body, fontSize: bodySize, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  }
}

/** Маркированный список с чертой вместо точки */
function ticks(s, { x, y, w, items, tone = "obj", size = 12, gap = 0.52 }) {
  const col = tone === "subj" ? C.subj : C.obj;
  items.forEach((runs, i) => {
    const yy = y + i * gap;
    s.addShape(pres.ShapeType.rect, {
      x, y: yy + 0.1, w: 0.16, h: 0.025, fill: { color: col }, line: { color: col, width: 0 },
    });
    s.addText(runs, {
      x: x + 0.3, y: yy, w: w - 0.3, h: gap,
      fontFace: F.body, fontSize: size, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  });
}

function footer(s) {
  s.addText(String(slideNo), {
    x: W - M - 0.6, y: H - 0.62, w: 0.6, h: 0.3,
    fontFace: F.body, fontSize: 10, color: C.faint,
    align: "right", isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("Структура конфликта", {
    x: M, y: H - 0.62, w: 4, h: 0.3,
    fontFace: F.body, fontSize: 9.5, color: C.faint,
    charSpacing: 1.2, isTextBox: true, margin: 0, valign: "middle",
  });
}

/* Строчные акценты */
const b = (t) => ({ text: t, options: { bold: true, color: C.text } });
const o = (t) => ({ text: t, options: { color: C.obj, bold: true } });
const sj = (t) => ({ text: t, options: { color: C.subj, bold: true } });
const n = (t) => ({ text: t });

/* ===============================================================
   1. ТИТУЛЬНЫЙ
   =============================================================== */
{
  const s = newSlide(
    "Тема: объективная и субъективная структура конфликта по А. Я. Анцупову и А. И. Шипилову. " +
    "Цель — показать, что конфликт это не просто ссора, а система элементов, каждый из которых можно назвать и разобрать."
  );

  // Два пересекающихся поля — объективное и субъективное
  s.addShape(pres.ShapeType.ellipse, {
    x: 8.55, y: 1.35, w: 3.1, h: 3.1,
    fill: { color: C.obj, transparency: 92 }, line: { color: C.obj, width: 1.25 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.55, y: 2.5, w: 3.1, h: 3.1,
    fill: { color: C.subj, transparency: 92 }, line: { color: C.subj, width: 1.25 },
  });

  s.addShape(pres.ShapeType.ellipse, { x: M, y: 0.9, w: 0.11, h: 0.11, fill: { color: C.obj } });
  s.addText("КОНФЛИКТОЛОГИЯ · СТРУКТУРА КОНФЛИКТА", {
    x: M + 0.22, y: 0.8, w: 7, h: 0.3,
    fontFace: F.body, fontSize: 10.5, bold: true, color: C.faint,
    charSpacing: 2.2, isTextBox: true, margin: 0, valign: "middle",
  });

  s.addText("Объективная и субъективная\nструктура конфликта", {
    x: M, y: 1.5, w: 7.5, h: 2.2,
    fontFace: F.head, fontSize: 40, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addText([
    n("Разбор двух подструктур конфликта по "),
    b("А. Я. Анцупову"), n(" и "), b("А. И. Шипилову"),
  ], {
    x: M, y: 3.85, w: 6.4, h: 0.7,
    fontFace: F.body, fontSize: 15, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, {
    x: M, y: 4.85, w: 1.2, h: 0.02, fill: { color: C.cardLine }, line: { color: C.cardLine, width: 0 },
  });

  s.addText("Выполнили студентки группы 8.536 зу", {
    x: M, y: 5.1, w: 6, h: 0.3,
    fontFace: F.body, fontSize: 11, color: C.faint,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("Кудякова Дарья", {
    x: M, y: 5.5, w: 3, h: 0.32,
    fontFace: F.body, fontSize: 14, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("Бороздова Валерия", {
    x: M + 3.1, y: 5.5, w: 3, h: 0.32,
    fontFace: F.body, fontSize: 14, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
}

/* ===============================================================
   2. ПОНЯТИЕ «КОНФЛИКТ»
   =============================================================== */
{
  const s = newSlide(
    "Конфликт (лат. conflictus — столкновение) — отсутствие согласия между двумя или более сторонами, " +
    "проявляющееся в несовместимости позиций или действий сторон, воспринимаемых ими как угроза себе, собственной идентичности. " +
    "Каждая сторона делает всё, чтобы принята была её точка зрения, и мешает другой стороне делать то же самое. " +
    "При этом конфликт всегда представляет собой выбор: сохранять существующую систему взаимодействий в неизменном виде или изменить её."
  );
  eyebrow(s, "01 · Исходная точка");
  title(s, "Понятие «конфликт»");

  // Цитата слева
  s.addText("«Отсутствие согласия между сторонами, проявляющееся в несовместимости позиций или действий, воспринимаемых ими как угроза себе».", {
    x: M, y: 2.05, w: 5.5, h: 1.9,
    fontFace: F.head, fontSize: 19, italic: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("Каждая сторона делает всё, чтобы принята была её точка зрения, и мешает другой стороне делать то же самое.", {
    x: M, y: 4.1, w: 5.5, h: 0.8,
    fontFace: F.body, fontSize: 12, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.05, w: 3.5, h: 0.42, rectRadius: 0.2,
    fill: { color: C.objCard }, line: { color: C.objLine, width: 1 },
  });
  s.addText("лат. conflictus — столкновение", {
    x: M, y: 5.05, w: 3.5, h: 0.42,
    fontFace: F.body, fontSize: 11, bold: true, color: C.obj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });

  // Три признака справа
  const rx = 6.7, rw = CW - (rx - M);
  [
    ["1", "Несовместимость", "Позиции или действия сторон не могут быть реализованы одновременно.", "obj"],
    ["2", "Восприятие угрозы", "Стороны считывают ситуацию как посягательство на себя и свою идентичность.", "obj"],
    ["3", "Конфликт — это выбор", "Сохранить систему взаимодействий неизменной — или изменить её.", "subj"],
  ].forEach(([num, head, body, tone], i) => {
    const y = 2.05 + i * 1.45;
    card(s, { x: rx, y, w: rw, h: 1.22, tone: "plain" });
    badge(s, { x: rx + 0.28, y: y + 0.3, label: num, tone, d: 0.38 });
    cardText(s, { x: rx + 0.82, y: y + 0.22, w: rw - 1.1, head, body, tone: "plain", headSize: 14, bodySize: 10.5 });
  });
  footer(s);
}

/* ===============================================================
   3. СТРУКТУРА — ДВЕ ПОДСТРУКТУРЫ
   =============================================================== */
{
  const s = newSlide(
    "Анцупов и Шипилов рассматривают структуру конфликта как совокупность устойчивых связей, которые обеспечивают его целостность. " +
    "Они выделяют две подструктуры — объективную и субъективную, каждая из которых включает как явные, так и скрытые элементы."
  );
  eyebrow(s, "02 · Общая схема");
  title(s, "Структура конфликта");

  lead(s, [
    n("Структура — это "), b("совокупность устойчивых связей"),
    n(", которые обеспечивают целостность конфликта. Каждая подструктура включает как "),
    b("явные"), n(", так и "), b("скрытые"), n(" элементы."),
  ], 1.9, 0.7, 13.5);

  const cw = 3.55;
  card(s, { x: M, y: 2.9, w: cw, h: 2.5, tone: "obj" });
  badge(s, { x: M + 0.3, y: 3.15, label: "О", tone: "obj", d: 0.42 });
  cardText(s, {
    x: M + 0.3, y: 3.75, w: cw - 0.6,
    head: "Объективная", tone: "obj", headSize: 18,
    body: "Внешние, фактические обстоятельства. Существуют независимо от того, как их видят участники.",
    bodySize: 11.5,
  });

  card(s, { x: M + cw + 0.35, y: 2.9, w: cw, h: 2.5, tone: "subj" });
  badge(s, { x: M + cw + 0.65, y: 3.15, label: "С", tone: "subj", d: 0.42 });
  cardText(s, {
    x: M + cw + 0.65, y: 3.75, w: cw - 0.6,
    head: "Субъективная", tone: "subj", headSize: 18,
    body: "Внутренние, психологические компоненты. Складываются в сознании участников.",
    bodySize: 11.5,
  });

  // Схема пересечения
  s.addShape(pres.ShapeType.ellipse, {
    x: 8.5, y: 2.5, w: 2.5, h: 2.5,
    fill: { color: C.obj, transparency: 88 }, line: { color: C.obj, width: 1.2 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.85, y: 2.5, w: 2.5, h: 2.5,
    fill: { color: C.subj, transparency: 88 }, line: { color: C.subj, width: 1.2 },
  });
  s.addText("ФАКТЫ", {
    x: 8.42, y: 3.55, w: 1.2, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.obj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("ОБРАЗЫ", {
    x: 11.22, y: 3.55, w: 1.25, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.subj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("РЕАЛЬНЫЙ\nКОНФЛИКТ", {
    x: 9.85, y: 3.45, w: 1.15, h: 0.6,
    fontFace: F.body, fontSize: 9.5, bold: true, color: C.text,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  footer(s);
}

/* ===============================================================
   4. АВТОРЫ ПОДХОДА
   =============================================================== */
{
  const s = newSlide(
    "А. Я. Анцупов и А. И. Шипилов — авторы учебника «Конфликтология», одного из базовых в российской традиции. " +
    "Они предложили рассматривать структуру конфликта как систему из двух взаимосвязанных подструктур."
  );
  eyebrow(s, "03 · Авторы подхода");
  title(s, "Анцупов и Шипилов");

  lead(s, [
    n("Авторы учебника «Конфликтология» — одного из базовых в российской традиции. Они предложили рассматривать структуру конфликта как систему из "),
    b("двух взаимосвязанных подструктур"), n("."),
  ], 1.9, 0.7, 13.5);

  const cw = 5.75;
  [["А. Я. Анцупов", "АЯ", "obj", M], ["А. И. Шипилов", "АИ", "subj", M + cw + 0.4]].forEach(([nm, mono, tone, x]) => {
    card(s, { x, y: 3.0, w: cw, h: 1.85, tone });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.4, y: 3.38, w: 1.1, h: 1.1,
      fill: { color: C.bg }, line: { color: tone === "obj" ? C.objLine : C.subjLine, width: 1.2 },
    });
    s.addText(mono, {
      x: x + 0.4, y: 3.38, w: 1.1, h: 1.1,
      fontFace: F.head, fontSize: 20, bold: true, color: tone === "obj" ? C.obj : C.subj,
      align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(nm, {
      x: x + 1.72, y: 3.62, w: cw - 2.1, h: 0.45,
      fontFace: F.head, fontSize: 19, bold: true, color: C.text,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText("Соавтор структурного подхода к анализу конфликта", {
      x: x + 1.72, y: 4.08, w: cw - 2.1, h: 0.5,
      fontFace: F.body, fontSize: 11, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  });

  s.addText("Структура — совокупность устойчивых связей, обеспечивающих целостность конфликта. В каждой подструктуре есть явные и скрытые элементы.", {
    x: M, y: 5.25, w: CW, h: 0.5,
    fontFace: F.body, fontSize: 12, color: C.faint,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   5. ОБЪЕКТИВНАЯ ПОДСТРУКТУРА — ОБЗОР
   =============================================================== */
{
  const s = newSlide(
    "Объективная подструктура отражает внешние, фактические обстоятельства конфликта, независимые от индивидуального восприятия его участников. " +
    "К её элементам относятся: участники (субъекты) конфликта, объект конфликта, предмет конфликта, микросреда и макросреда."
  );
  eyebrow(s, "04 · Объективная подструктура");
  title(s, "Пять внешних элементов");
  lead(s, [
    n("Отражает внешние, фактические обстоятельства конфликта, "),
    b("независимые от индивидуального восприятия"), n(" его участников."),
  ], 1.88, 0.5, 13);

  const items = [
    ["1", "Участники", "Противоборствующие стороны, группы поддержки, наблюдатели."],
    ["2", "Объект", "То, из-за чего возникает противостояние: ресурс, статус, принцип."],
    ["3", "Предмет", "Основное противоречие, из-за которого стороны вступают в борьбу."],
    ["4", "Микросреда", "Непосредственное окружение: место, время, климат в группе."],
    ["5", "Макросреда", "Социальный, культурный, экономический и политический контекст."],
  ];
  const gap = 0.24;
  const cw = (CW - gap * 4) / 5;
  items.forEach(([num, head, body], i) => {
    const x = M + i * (cw + gap);
    card(s, { x, y: 2.75, w: cw, h: 2.3, tone: "obj" });
    badge(s, { x: x + 0.24, y: 3.0, label: num, tone: "obj", d: 0.4 });
    cardText(s, {
      x: x + 0.24, y: 3.58, w: cw - 0.48,
      head, body, tone: "obj", headSize: 14.5, bodySize: 10.5,
    });
  });
  footer(s);
}

/* ===============================================================
   6. УЧАСТНИКИ
   =============================================================== */
{
  const s = newSlide(
    "Участники конфликта — это противоборствующие стороны, а также группы поддержки (те, кто прямо или косвенно содействует одной из сторон), " +
    "наблюдатели и другие вовлечённые лица. Важный нюанс: сторона — это не просто жертва или наблюдатель, а тот, кто имеет свою позицию и силу. " +
    "Иногда за спиной прямых участников стоят третьи лица — подстрекатели, союзники, — но они не являются сторонами в строгом смысле, пока не начнут действовать открыто."
  );
  eyebrow(s, "04.1 · Объективная подструктура");
  title(s, "Участники конфликта");

  ticks(s, {
    x: M, y: 2.0, w: 6.3, gap: 0.66, size: 12.5, tone: "obj",
    items: [
      [b("Противоборствующие стороны"), n(" — те, кто непосредственно «столкнулся лбами».")],
      [b("Группы поддержки"), n(" — те, кто прямо или косвенно содействует одной из сторон.")],
      [b("Наблюдатели"), n(" и другие вовлечённые лица.")],
    ],
  });

  card(s, { x: M, y: 4.2, w: 6.3, h: 1.75, tone: "obj" });
  s.addText("Важный нюанс", {
    x: M + 0.3, y: 4.42, w: 5.7, h: 0.35,
    fontFace: F.body, fontSize: 14, bold: true, color: C.obj,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText([
    n("Сторона — это не просто жертва или наблюдатель. Это тот, кто имеет "),
    b("свою позицию и силу"),
    n(". Третьи лица (подстрекатели, союзники) не являются сторонами, пока не начнут действовать открыто."),
  ], {
    x: M + 0.3, y: 4.8, w: 5.7, h: 1.0,
    fontFace: F.body, fontSize: 11, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  // Схема ролей
  const bx = 7.6;
  s.addShape(pres.ShapeType.line, {
    x: bx + 2.55, y: 2.95, w: 0, h: 2.7,
    line: { color: C.cardLine, width: 1, dashType: "dash" },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 1.2, y: 3.25, w: 1.1, h: 1.1,
    fill: { color: C.objCard }, line: { color: C.obj, width: 1.4 },
  });
  s.addText("А", {
    x: bx + 1.2, y: 3.25, w: 1.1, h: 1.1,
    fontFace: F.head, fontSize: 16, bold: true, color: C.obj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 2.8, y: 3.25, w: 1.1, h: 1.1,
    fill: { color: C.subjCard }, line: { color: C.subj, width: 1.4 },
  });
  s.addText("Б", {
    x: bx + 2.8, y: 3.25, w: 1.1, h: 1.1,
    fontFace: F.head, fontSize: 16, bold: true, color: C.subj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 0.15, y: 4.72, w: 0.68, h: 0.68,
    fill: { color: C.card }, line: { color: C.cardLine, width: 1 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 4.28, y: 4.72, w: 0.68, h: 0.68,
    fill: { color: C.card }, line: { color: C.cardLine, width: 1 },
  });
  s.addText("группа\nподдержки", {
    x: bx - 0.15, y: 5.46, w: 1.3, h: 0.5,
    fontFace: F.body, fontSize: 9, color: C.faint,
    align: "center", valign: "top", isTextBox: true, margin: 0,
  });
  s.addText("группа\nподдержки", {
    x: bx + 3.98, y: 5.46, w: 1.3, h: 0.5,
    fontFace: F.body, fontSize: 9, color: C.faint,
    align: "center", valign: "top", isTextBox: true, margin: 0,
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 2.25, y: 2.2, w: 0.6, h: 0.6,
    fill: { color: C.bg }, line: { color: C.cardLine, width: 1, dashType: "dash" },
  });
  s.addText("наблюдатели", {
    x: bx + 1.8, y: 1.84, w: 1.5, h: 0.3,
    fontFace: F.body, fontSize: 9, color: C.faint,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  footer(s);
}

/* ===============================================================
   7. ОБЪЕКТ
   =============================================================== */
{
  const s = newSlide(
    "Объект конфликта — это то, что находится в центре борьбы: реальная ценность, которую хотят заполучить или сохранить. " +
    "Территория, деньги, должность, право голоса, ресурс. Объект может быть материальным (земля) или нематериальным (власть, уважение). " +
    "Объект может быть как реальным, так и идеальным. Если объекта нет, конфликт превращается в пустую перепалку."
  );
  eyebrow(s, "04.2 · Объективная подструктура");
  title(s, "Объект конфликта");
  lead(s, [
    n("То, что находится "), b("в центре борьбы"),
    n(": реальная ценность, которую хотят заполучить или сохранить. Объект может быть как реальным, так и идеальным."),
  ], 1.88, 0.55, 13);

  const trio = [
    ["Материальный", "территория · деньги · ресурс", "obj"],
    ["Статусный", "должность · право голоса", "obj"],
    ["Идеальный", "власть · уважение · справедливость", "subj"],
  ];
  const gap = 0.3;
  const cw = (CW - gap * 2) / 3;
  trio.forEach(([head, body, tone], i) => {
    const x = M + i * (cw + gap);
    card(s, { x, y: 2.75, w: cw, h: 1.5, tone });
    s.addText(head, {
      x: x + 0.3, y: 2.98, w: cw - 0.6, h: 0.4,
      fontFace: F.body, fontSize: 16, bold: true, color: tone === "subj" ? C.subj : C.obj,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(body, {
      x: x + 0.3, y: 3.42, w: cw - 0.6, h: 0.6,
      fontFace: F.body, fontSize: 11.5, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  });

  card(s, { x: M, y: 4.55, w: CW, h: 1.25, tone: "subj" });
  s.addText([
    b("Если объекта нет"),
    n(" — конфликт превращается в пустую перепалку: сторонам не за что бороться, остаётся только эмоциональное напряжение."),
  ], {
    x: M + 0.4, y: 4.88, w: CW - 0.8, h: 0.7,
    fontFace: F.body, fontSize: 13.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   8. ПРЕДМЕТ ≠ ОБЪЕКТ
   =============================================================== */
{
  const s = newSlide(
    "Предмет конфликта — это основное противоречие, из-за которого стороны вступают в противоборство; то, из-за чего спор возник здесь и сейчас. " +
    "Например, объект — квартира, а предмет — кто будет в ней жить после развода или как делить метры. " +
    "Предмет может меняться по ходу выяснения отношений, а объект остаётся неизменным. " +
    "Предмет не всегда совпадает с объектом и может быть скрытым или неосознанным."
  );
  eyebrow(s, "04.3 · Объективная подструктура");
  title(s, "Предмет ≠ объект");

  const cw = (CW - 0.5) / 2;
  const rows = [
    ["Вопрос", "Из-за чего в принципе идёт борьба?", "Из-за чего спор возник здесь и сейчас?"],
    ["Природа", "Ценность, которую хотят получить или сохранить", "Основное противоречие, «фокус» столкновения интересов"],
    ["Устойчивость", "Как правило, остаётся неизменным", "Может меняться по ходу выяснения отношений"],
    ["Пример", "Квартира", "Кто будет в ней жить после развода и как делить метры"],
  ];

  [["Объект", "obj", M], ["Предмет", "subj", M + cw + 0.5]].forEach(([head, tone, x], col) => {
    card(s, { x, y: 1.95, w: cw, h: 3.85, tone });
    s.addText(head, {
      x: x + 0.35, y: 2.2, w: cw - 0.7, h: 0.5,
      fontFace: F.head, fontSize: 21, bold: true, color: tone === "obj" ? C.obj : C.subj,
      isTextBox: true, margin: 0, valign: "top",
    });
    rows.forEach(([label, a, bb], i) => {
      const y = 2.85 + i * 0.73;
      s.addText(label.toUpperCase(), {
        x: x + 0.35, y, w: cw - 0.7, h: 0.24,
        fontFace: F.body, fontSize: 9, bold: true, color: C.faint,
        charSpacing: 1.6, isTextBox: true, margin: 0, valign: "top",
      });
      s.addText(col === 0 ? a : bb, {
        x: x + 0.35, y: y + 0.24, w: cw - 0.7, h: 0.46,
        fontFace: F.body, fontSize: 11.5, color: C.mute,
        isTextBox: true, margin: 0, valign: "top",
      });
    });
  });

  s.addText([
    n("Предмет не всегда совпадает с объектом и может быть "),
    b("скрытым или неосознанным"),
    n(" — стороны спорят о метрах, а на деле борются за признание."),
  ], {
    x: M, y: 6.0, w: CW, h: 0.5,
    fontFace: F.body, fontSize: 12.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   9. СРЕДА КОНФЛИКТА
   =============================================================== */
{
  const s = newSlide(
    "Микросреда — ближайшее окружение конфликта: семья, друзья, коллеги по отделу, соседи; локация, временные рамки, " +
    "социально-психологический климат в группе участников. Микросреда может подогревать страсти или, наоборот, гасить их. " +
    "Макросреда — более широкий социальный, культурный, экономический и политический контекст: законы, традиции, уровень напряжённости в обществе. " +
    "Если в стране кризис и безработица, конфликт на работе из-за премии протекает иначе, чем в благополучное время."
  );
  eyebrow(s, "04.4 · Объективная подструктура");
  title(s, "Среда конфликта");

  const cw = (CW - 0.45) / 2;
  const blocks = [
    ["Микро", "ближайшее окружение, влияющее «изнутри»", "obj", M,
      "Семья, друзья, коллеги по отделу, соседи. Локация, временные рамки, социально-психологический климат в группе участников.",
      "Может подогревать страсти («Да он тебя не уважает!») или, наоборот, гасить их. Это люди, с которыми стороны контактируют напрямую и чьё мнение для них важно."],
    ["Макро", "«большой мир», задающий правила игры", "subj", M + cw + 0.45,
      "Социальные, экономические, политические и культурные условия: законы, традиции, уровень напряжённости в обществе.",
      "Если в стране кризис и безработица, конфликт на работе из-за премии протекает иначе. Макросреда может сдерживать конфликт или раздувать его до неконтролируемых масштабов."],
  ];
  blocks.forEach(([big, sub, tone, x, p1, p2]) => {
    const col = tone === "obj" ? C.obj : C.subj;
    card(s, { x, y: 2.0, w: cw, h: 3.9, tone });
    s.addText(big, {
      x: x + 0.38, y: 2.28, w: 1.9, h: 0.65,
      fontFace: F.head, fontSize: 32, bold: true, color: col,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(sub, {
      x: x + 2.3, y: 2.42, w: cw - 2.7, h: 0.6,
      fontFace: F.body, fontSize: 11, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(p1, {
      x: x + 0.38, y: 3.2, w: cw - 0.76, h: 1.0,
      fontFace: F.body, fontSize: 12, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(p2, {
      x: x + 0.38, y: 4.3, w: cw - 0.76, h: 1.4,
      fontFace: F.body, fontSize: 12, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  });
  footer(s);
}

/* ===============================================================
   10. СУБЪЕКТИВНАЯ ПОДСТРУКТУРА — ОБЗОР
   =============================================================== */
{
  const s = newSlide(
    "В субъективную подструктуру входят внутренние, психологические компоненты, которые формируются в сознании участников " +
    "и во многом определяют динамику и характер конфликтного взаимодействия: психологические модели конфликтной ситуации, " +
    "мотивы действий сторон, цели участников, актуальные психические состояния, образы оппонента, самого себя, объекта и предмета конфликта, " +
    "а также вероятные результаты борьбы."
  );
  eyebrow(s, "05 · Субъективная подструктура", "subj");
  title(s, "Шесть внутренних элементов");
  lead(s, [
    n("Психологические компоненты, которые формируются "),
    b("в сознании участников"), n(" и определяют динамику конфликтного взаимодействия."),
  ], 1.88, 0.5, 13);

  const items = [
    ["1", "Психологические модели ситуации", "Своя картина происходящего у каждого участника."],
    ["2", "Мотивы действий", "Внутренние побуждения к конфликтному поведению."],
    ["3", "Цели участников", "Представления о том, чего они хотят добиться."],
    ["4", "Психические состояния", "Эмоциональный фон, напряжение, стрессовые реакции."],
    ["5", "Образы оппонента и себя", "Когнитивные конструкции, часто искажённые."],
    ["6", "Вероятные результаты", "Исходы борьбы, которые участники мысленно моделируют."],
  ];
  const gap = 0.28;
  const cw = (CW - gap * 2) / 3;
  const ch = 1.42;
  items.forEach(([num, head, body], i) => {
    const x = M + (i % 3) * (cw + gap);
    const y = 2.72 + Math.floor(i / 3) * (ch + 0.3);
    card(s, { x, y, w: cw, h: ch, tone: "subj" });
    badge(s, { x: x + 0.26, y: y + 0.24, label: num, tone: "subj", d: 0.36 });
    cardText(s, {
      x: x + 0.72, y: y + 0.2, w: cw - 0.98,
      head, body, tone: "subj", headSize: 13, bodySize: 10.5,
    });
  });
  footer(s);
}

/* ===============================================================
   11. МОДЕЛИ И ОБРАЗЫ
   =============================================================== */
{
  const s = newSlide(
    "У каждого участника складывается своя картина происходящего, которая не всегда совпадает с объективной реальностью. " +
    "Образы — это когнитивные конструкции: представления каждого участника о другой стороне, о собственных качествах, " +
    "о том, из-за чего возник спор, и о сути самого противоречия. Эти образы часто искажены — например, происходит «демонизация» противника."
  );
  eyebrow(s, "05.1 · Субъективная подструктура", "subj");
  title(s, "Модели и образы");

  lead(s, [
    n("У каждого участника складывается "), b("своя картина происходящего"),
    n(", которая не всегда совпадает с объективной реальностью. Образы — это конструкции о четырёх вещах:"),
  ], 1.9, 0.65, 12.5);

  ticks(s, {
    x: M, y: 2.72, w: 6.4, gap: 0.5, size: 12, tone: "subj",
    items: [
      [b("Образ оппонента"), n(" — каким видится другая сторона.")],
      [b("Образ себя"), n(" — представление о своих качествах и правоте.")],
      [b("Образ объекта"), n(" — из-за чего, по мнению участника, возник спор.")],
      [b("Образ предмета"), n(" — понимание сути противоречия.")],
    ],
  });

  card(s, { x: M, y: 4.85, w: 6.4, h: 1.25, tone: "subj" });
  s.addText([
    n("Эти образы часто "), b("искажены"),
    n(": происходит «демонизация» противника, а собственные действия выглядят вынужденными и оправданными."),
  ], {
    x: M + 0.32, y: 5.1, w: 5.76, h: 0.8,
    fontFace: F.body, fontSize: 11.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  // Схема искажения
  const bx = 7.85;
  s.addShape(pres.ShapeType.roundRect, {
    x: bx, y: 2.5, w: 2.1, h: 1.25, rectRadius: 0.1,
    fill: { color: C.objCard }, line: { color: C.obj, width: 1.2 },
  });
  s.addText("ЧТО ЕСТЬ", {
    x: bx, y: 2.78, w: 2.1, h: 0.3,
    fontFace: F.body, fontSize: 11, bold: true, color: C.obj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("объективно", {
    x: bx, y: 3.08, w: 2.1, h: 0.28,
    fontFace: F.body, fontSize: 9.5, color: C.mute,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: bx + 2.6, y: 4.45, w: 2.1, h: 1.25, rectRadius: 0.1,
    fill: { color: C.subjCard }, line: { color: C.subj, width: 1.2 },
  });
  s.addText("ЧТО ВИДНО", {
    x: bx + 2.6, y: 4.73, w: 2.1, h: 0.3,
    fontFace: F.body, fontSize: 11, bold: true, color: C.subj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("субъективно", {
    x: bx + 2.6, y: 5.03, w: 2.1, h: 0.28,
    fontFace: F.body, fontSize: 9.5, color: C.mute,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });

  s.addShape(pres.ShapeType.line, {
    x: bx + 2.1, y: 3.12, w: 1.55, h: 1.33,
    line: { color: C.faint, width: 1.4, dashType: "dash", endArrowType: "triangle" },
  });
  s.addText("искажение", {
    x: bx + 1.75, y: 3.68, w: 1.5, h: 0.28,
    fontFace: F.body, fontSize: 9.5, bold: true, color: C.faint,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  footer(s);
}

/* ===============================================================
   12. МОТИВЫ → ЦЕЛИ → ПРОГНОЗ
   =============================================================== */
{
  const s = newSlide(
    "Мотивы действий сторон — внутренние побуждения, которые толкают участников на конфликтное поведение. " +
    "Цели — представления о том, чего они хотят добиться в конфликте. " +
    "Вероятные результаты борьбы — прогнозируемые исходы, которые участники мысленно моделируют. " +
    "Расхождение между заявленной целью и реальным мотивом — одна из главных причин, по которой конфликт не разрешается по существу."
  );
  eyebrow(s, "05.2 · Субъективная подструктура", "subj");
  title(s, "Мотивы → цели → прогноз");

  const trio = [
    ["1", "Мотивы", "Внутренние побуждения, которые толкают участников на конфликтное поведение. Часто не осознаются и не проговариваются вслух."],
    ["2", "Цели", "Представления о том, чего участник хочет добиться. В отличие от мотивов — формулируются и предъявляются другой стороне."],
    ["3", "Вероятные результаты", "Прогнозируемые исходы борьбы. Оценка шансов напрямую влияет на готовность идти на обострение."],
  ];
  const gap = 0.32;
  const cw = (CW - gap * 2) / 3;
  trio.forEach(([num, head, body], i) => {
    const x = M + i * (cw + gap);
    card(s, { x, y: 2.1, w: cw, h: 2.75, tone: "subj" });
    badge(s, { x: x + 0.34, y: 2.38, label: num, tone: "subj", d: 0.42 });
    cardText(s, {
      x: x + 0.34, y: 3.0, w: cw - 0.68,
      head, body, tone: "subj", headSize: 18, bodySize: 11.5,
    });

    if (i < 2) {
      s.addShape(pres.ShapeType.line, {
        x: x + cw + 0.05, y: 3.45, w: gap - 0.1, h: 0,
        line: { color: C.subjLine, width: 1.6, endArrowType: "triangle" },
      });
    }
  });

  card(s, { x: M, y: 5.15, w: CW, h: 1.05, tone: "plain" });
  s.addText([
    n("Расхождение между "), sj("заявленной целью"), n(" и "), sj("реальным мотивом"),
    n(" — одна из главных причин, по которой конфликт не разрешается «по существу»."),
  ], {
    x: M + 0.4, y: 5.42, w: CW - 0.8, h: 0.55,
    fontFace: F.body, fontSize: 13, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   13. ПСИХИЧЕСКИЕ СОСТОЯНИЯ
   =============================================================== */
{
  const s = newSlide(
    "Актуальные психические состояния участников — эмоциональный фон, уровень напряжения, стрессовые реакции, которые влияют на поведение. " +
    "Чем выше напряжение, тем меньше вариантов выхода видит участник. Это единственный элемент структуры, который может измениться за минуты."
  );
  eyebrow(s, "05.3 · Субъективная подструктура", "subj");
  title(s, "Актуальные психические состояния");

  lead(s, [
    n("Эмоциональный фон, уровень напряжения и стрессовые реакции, которые "),
    b("влияют на поведение прямо в моменте"), n("."),
  ], 1.9, 0.6, 13);

  ticks(s, {
    x: M, y: 2.75, w: 6.3, gap: 0.72, size: 12.5, tone: "subj",
    items: [
      [n("Высокое напряжение "), b("сужает выбор"), n(" — участник видит меньше вариантов выхода.")],
      [n("Состояние меняется по ходу конфликта, поэтому одни и те же аргументы срабатывают по-разному.")],
      [n("Это единственный элемент структуры, который может измениться за минуты.")],
    ],
  });

  // Схема: напряжение растёт, варианты выхода сужаются
  const bx = 7.75, by = 2.55, bw = 4.3, bh = 2.9;
  card(s, { x: bx, y: by, w: bw, h: bh, tone: "plain" });

  s.addShape(pres.ShapeType.line, {
    x: bx + 0.55, y: by + 0.6, w: 0, h: 1.55, flipV: true,
    line: { color: C.subj, width: 2.4, endArrowType: "triangle" },
  });
  s.addText("напряжение\nрастёт", {
    x: bx + 0.85, y: by + 0.62, w: 1.4, h: 0.6,
    fontFace: F.body, fontSize: 11, bold: true, color: C.subj,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addShape(pres.ShapeType.line, {
    x: bx + 2.5, y: by + 0.6, w: 0, h: 1.55,
    line: { color: C.obj, width: 2.4, endArrowType: "triangle" },
  });
  s.addText("варианты\nвыхода\nсужаются", {
    x: bx + 2.8, y: by + 1.3, w: 1.4, h: 0.9,
    fontFace: F.body, fontSize: 11, bold: true, color: C.obj,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addText("время конфликта", {
    x: bx + 0.3, y: by + 2.32, w: bw - 0.6, h: 0.3,
    fontFace: F.body, fontSize: 10, color: C.faint,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  footer(s);
}

/* ===============================================================
   14. ВЗАИМОСВЯЗЬ ПОДСТРУКТУР
   =============================================================== */
{
  const s = newSlide(
    "Анцупов и Шипилов подчёркивают, что объективная и субъективная подструктуры не существуют изолированно. " +
    "Любая объективная причина конфликта неизбежно преломляется через субъективные восприятия и оценки участников. " +
    "И наоборот — субъективные мотивы и образы во многом детерминированы объективными обстоятельствами. " +
    "Поэтому для полноценного анализа конфликта необходимо рассматривать оба пласта."
  );
  eyebrow(s, "06 · Ключевой вывод");
  title(s, "Подструктуры не существуют изолированно", 30);

  s.addText([
    n("Любая "), o("объективная причина"),
    n(" конфликта неизбежно преломляется через субъективные восприятия и оценки участников. И наоборот — "),
    sj("субъективные мотивы и образы"),
    n(" во многом детерминированы объективными обстоятельствами."),
  ], {
    x: M, y: 2.0, w: 6.4, h: 1.5,
    fontFace: F.body, fontSize: 13.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  card(s, { x: M, y: 3.75, w: 6.4, h: 1.85, tone: "plain" });
  s.addText([
    n("Поэтому для полноценного анализа конфликта "),
    b("необходимо рассматривать оба пласта"),
    n(". Разбор только фактов даёт формальную картину, разбор только переживаний — не находит причины."),
  ], {
    x: M + 0.35, y: 4.05, w: 5.7, h: 1.3,
    fontFace: F.body, fontSize: 12.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  // Два поля и стрелки взаимного влияния
  const bx = 7.8;
  s.addShape(pres.ShapeType.ellipse, {
    x: bx, y: 2.15, w: 2.0, h: 2.0,
    fill: { color: C.obj, transparency: 88 }, line: { color: C.obj, width: 1.3 },
  });
  s.addText("ОБЪЕКТИВНОЕ", {
    x: bx, y: 2.95, w: 2.0, h: 0.3,
    fontFace: F.body, fontSize: 10.5, bold: true, color: C.obj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("факты", {
    x: bx, y: 3.25, w: 2.0, h: 0.26,
    fontFace: F.body, fontSize: 9.5, color: C.mute,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });

  s.addShape(pres.ShapeType.ellipse, {
    x: bx + 2.35, y: 3.55, w: 2.0, h: 2.0,
    fill: { color: C.subj, transparency: 88 }, line: { color: C.subj, width: 1.3 },
  });
  s.addText("СУБЪЕКТИВНОЕ", {
    x: bx + 2.35, y: 4.35, w: 2.0, h: 0.3,
    fontFace: F.body, fontSize: 10.5, bold: true, color: C.subj,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("образы", {
    x: bx + 2.35, y: 4.65, w: 2.0, h: 0.26,
    fontFace: F.body, fontSize: 9.5, color: C.mute,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });

  s.addShape(pres.ShapeType.line, {
    x: bx + 1.87, y: 3.40, w: 0.75, h: 0.6,
    line: { color: C.obj, width: 1.8, endArrowType: "triangle" },
  });
  s.addShape(pres.ShapeType.line, {
    x: bx + 1.54, y: 3.60, w: 0.72, h: 0.58, flipH: true, flipV: true,
    line: { color: C.subj, width: 1.8, endArrowType: "triangle" },
  });
  s.addText("взаимная детерминация", {
    x: bx + 0.6, y: 5.75, w: 3.5, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.faint,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  footer(s);
}

/* ===============================================================
   15. РАЗБОР КЕЙСА
   =============================================================== */
{
  const s = newSlide(
    "Сквозной пример. Объект (квартира) делится по закону, но конфликт не гаснет, пока не признан реальный предмет — " +
    "спор о признании вклада. Это показывает, зачем вообще нужен разбор по элементам: он вскрывает то, о чём стороны не говорят прямо."
  );
  eyebrow(s, "07 · Практика");
  title(s, "Разбор кейса по элементам");
  s.addText([
    n("Ситуация: "),
    { text: "бывшие супруги делят квартиру; каждый считает, что вложил больше.", options: { italic: true, color: C.text } },
  ], {
    x: M, y: 1.88, w: CW, h: 0.4,
    fontFace: F.body, fontSize: 13, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  const cw = (CW - 0.45) / 2;
  const objRows = [
    ["Участники:", " бывшие супруги; группы поддержки — родители с обеих сторон; наблюдатель — общий ребёнок."],
    ["Объект:", " квартира как имущественная ценность."],
    ["Предмет:", " кто останется в ней жить и как делить метры."],
    ["Микросреда:", " семьи, общие друзья, юристы, сроки судебного процесса."],
    ["Макросреда:", " семейное законодательство, цены на жильё, нормы о «правильном» разводе."],
  ];
  const subjRows = [
    ["Модель ситуации:", " у каждого своя версия того, «кто больше вложил»."],
    ["Мотив:", " не потерять лицо и получить признание вклада."],
    ["Цель:", " закрепить за собой бо́льшую долю."],
    ["Состояние:", " высокая тревога, усталость, обида после развода."],
    ["Образы:", " оппонент — «корыстный», себя — «пострадавшая сторона»."],
    ["Прогноз:", " «суд всё равно будет на моей стороне»."],
  ];

  [["Объективная подструктура", "obj", M, objRows], ["Субъективная подструктура", "subj", M + cw + 0.45, subjRows]]
    .forEach(([head, tone, x, rows]) => {
      card(s, { x, y: 2.45, w: cw, h: 3.4, tone });
      s.addText(head, {
        x: x + 0.32, y: 2.68, w: cw - 0.64, h: 0.4,
        fontFace: F.head, fontSize: 17, bold: true, color: tone === "obj" ? C.obj : C.subj,
        isTextBox: true, margin: 0, valign: "top",
      });
      const step = tone === "obj" ? 0.53 : 0.44;
      rows.forEach(([label, txt], i) => {
        s.addText([b(label), n(txt)], {
          x: x + 0.32, y: 3.18 + i * step, w: cw - 0.64, h: step,
          fontFace: F.body, fontSize: 10.5, color: C.mute,
          isTextBox: true, margin: 0, valign: "top",
        });
      });
    });

  s.addText([
    b("Что даёт разбор: "),
    n("объект (квартира) делится по закону, но конфликт не гаснет, пока не признан реальный предмет — "),
    sj("спор о признании вклада"), n("."),
  ], {
    x: M, y: 6.05, w: CW, h: 0.5,
    fontFace: F.body, fontSize: 12.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   16. ИТОГИ
   =============================================================== */
{
  const s = newSlide(
    "Главное: структура — это устойчивые связи, обеспечивающие целостность конфликта. " +
    "Объективная подструктура — участники, объект, предмет, микро- и макросреда. " +
    "Субъективная — модели, мотивы, цели, состояния, образы и прогноз результатов. " +
    "Оба пласта связаны, анализировать конфликт нужно одновременно по обоим. " +
    "В каждой подструктуре есть явные и скрытые элементы, и чаще всего именно скрытые определяют, разрешится конфликт или повторится."
  );
  eyebrow(s, "08 · Итоги");
  title(s, "Что важно запомнить");

  const items = [
    ["01", "Структура — это устойчивые связи, обеспечивающие целостность конфликта", "obj"],
    ["02", "Объективная подструктура: участники, объект, предмет, микро- и макросреда", "obj"],
    ["03", "Субъективная: модели, мотивы, цели, состояния, образы, прогноз результатов", "subj"],
    ["04", "Оба пласта связаны — анализировать конфликт нужно одновременно по обоим", "subj"],
  ];
  const gap = 0.32;
  const cw = (CW - gap) / 2;
  const ch = 1.45;
  items.forEach(([num, txt, tone], i) => {
    const x = M + (i % 2) * (cw + gap);
    const y = 2.1 + Math.floor(i / 2) * (ch + gap);
    card(s, { x, y, w: cw, h: ch, tone });
    s.addText(num, {
      x: x + 0.34, y: y + 0.26, w: 1.1, h: 0.75,
      fontFace: F.head, fontSize: 36, bold: true, color: tone === "obj" ? C.obj : C.subj,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(txt, {
      x: x + 1.42, y: y + 0.34, w: cw - 1.76, h: 0.9,
      fontFace: F.body, fontSize: 12.5, color: C.mute,
      isTextBox: true, margin: 0, valign: "top",
    });
  });

  s.addText([
    n("Каждая подструктура включает "), b("явные"), n(" и "), b("скрытые"),
    n(" элементы — и чаще всего именно скрытые определяют, разрешится конфликт или повторится."),
  ], {
    x: M, y: 5.7, w: CW, h: 0.55,
    fontFace: F.body, fontSize: 13, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });
  footer(s);
}

/* ===============================================================
   17. ФИНАЛ
   =============================================================== */
{
  const s = newSlide("Спасибо за внимание. Готовы ответить на вопросы.");

  s.addShape(pres.ShapeType.ellipse, {
    x: 8.7, y: 1.8, w: 2.8, h: 2.8,
    fill: { color: C.obj, transparency: 92 }, line: { color: C.obj, width: 1.2 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.6, y: 2.7, w: 2.8, h: 2.8,
    fill: { color: C.subj, transparency: 92 }, line: { color: C.subj, width: 1.2 },
  });

  s.addText("Спасибо за внимание", {
    x: M, y: 2.25, w: 7.2, h: 1.1,
    fontFace: F.head, fontSize: 40, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addText("ИСТОЧНИК", {
    x: M, y: 3.75, w: 6, h: 0.28,
    fontFace: F.body, fontSize: 10, bold: true, color: C.faint,
    charSpacing: 2, isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("Анцупов А. Я., Шипилов А. И. Конфликтология: учебник для вузов. — М.: ЮНИТИ-ДАНА.", {
    x: M, y: 4.06, w: 6.4, h: 0.6,
    fontFace: F.body, fontSize: 12.5, color: C.mute,
    isTextBox: true, margin: 0, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, {
    x: M, y: 5.0, w: 1.2, h: 0.02, fill: { color: C.cardLine }, line: { color: C.cardLine, width: 0 },
  });
  s.addText("Кудякова Дарья", {
    x: M, y: 5.25, w: 3, h: 0.32,
    fontFace: F.body, fontSize: 14, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("группа 8.536 зу", {
    x: M, y: 5.58, w: 3, h: 0.3,
    fontFace: F.body, fontSize: 11, color: C.faint,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("Бороздова Валерия", {
    x: M + 3.1, y: 5.25, w: 3, h: 0.32,
    fontFace: F.body, fontSize: 14, bold: true, color: C.text,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("группа 8.536 зу", {
    x: M + 3.1, y: 5.58, w: 3, h: 0.3,
    fontFace: F.body, fontSize: 11, color: C.faint,
    isTextBox: true, margin: 0, valign: "top",
  });
}

/* --------------------------------------------------------------- */
const out = "struktura-konflikta.pptx";
await pres.writeFile({ fileName: out });
console.log(`Готово: ${out} — ${slideNo} слайдов`);
