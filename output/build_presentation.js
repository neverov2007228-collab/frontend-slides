// Генератор презентации «Основные положения Международной конвенции
// об упрощении и гармонизации таможенных процедур».
// Запуск: node build.js <путь_к_выходному_файлу.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Академическая",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "333333", lt1: "FFFFFF", dk2: "1F3A5F", lt2: "F2F2F2",
    accent1: "1F3A5F", accent2: "595959", accent3: "7F7F7F",
    accent4: "A6A6A6", accent5: "BFBFBF", accent6: "D9D9D9",
    hlink: "1F3A5F", folHlink: "595959",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 × 7.5 дюйма, 16:9
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Основные положения Международной конвенции об упрощении и гармонизации таможенных процедур";
pres.lang = "ru-RU";
const C = pres.SchemeColor;

// Геометрия
const M = 0.6;               // поля
const W = 13.333 - 2 * M;    // ширина рабочей области
const BODY_Y = 1.75;         // верх рабочей области
const FOOT_Y = 6.95;         // строка подписи

// Размеры шрифта
const F = { title: 28, body: 20, scheme: 14, schemeHead: 18, foot: 10 };

// Макеты
const footerNum = { x: M, y: FOOT_Y, w: 0.4, h: 0.3, fontSize: F.foot, color: C.text1, align: "left", valign: "middle" };

pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.background1 },
  slideNumber: { ...footerNum },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 1.5, w: W, h: 1.9,
      fontSize: 30, bold: true, color: C.text1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 3.9, w: W, h: 2.4,
      fontSize: F.body, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  slideNumber: { ...footerNum },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.5, w: W, h: 1.05,
      fontSize: F.title, bold: true, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});

// Вспомогательные функции
function source(slide, text) {
  slide.addText("Источник: " + text, {
    x: M + 0.45, y: FOOT_Y, w: W - 0.45, h: 0.3, fontSize: F.foot, color: C.text1,
    align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Источник",
  });
}

function content(title, src) {
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText(title, { placeholder: "title" });
  source(s, src);
  return s;
}

// Прямоугольник схемы: filled — тёмно-синяя заливка, иначе белый с тёмно-синей рамкой
function box(slide, x, y, w, h, runs, { filled = false, name = "Блок схемы" } = {}) {
  slide.addText(runs, {
    shape: pres.shapes.RECTANGLE, x, y, w, h,
    fill: { color: filled ? C.accent1 : C.background1 },
    line: { color: C.accent1, width: 1.5 },
    color: filled ? C.background1 : C.text1,
    fontSize: F.scheme, align: "left", valign: "middle", margin: 8,
    objectName: name,
  });
}

// Заголовок блока (жирный) + пояснение 14 pt
function headRuns(head, desc) {
  const r = [{ text: head, options: { bold: true, fontSize: F.schemeHead, breakLine: !!desc } }];
  if (desc) r.push({ text: desc, options: { fontSize: F.scheme } });
  return r;
}

function arrow(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.LINE, {
    x, y, w, h, flipV: !!opts.flipV,
    line: { color: C.accent1, width: 2, endArrowType: "triangle" },
    objectName: "Стрелка",
  });
}

function bullets(slide, items, { y = BODY_Y, h = 4.6 } = {}) {
  slide.addText(items.map((t, i) => ({
    text: t,
    options: { bullet: { indent: 22 }, breakLine: i < items.length - 1, paraSpaceAfter: 16 },
  })), {
    x: M, y, w: W, h, fontSize: F.body, color: C.text1,
    align: "left", valign: "top", margin: 0, isTextBox: true, objectName: "Пункты",
  });
}

const SRC = {
  conv: "текст конвенции в редакции 1999 года, pavlodar.com/zakon/info?dok=04328",
  wco: "Всемирная таможенная организация, www.wcoomd.org",
  code: "Таможенный кодекс Евразийского экономического союза, www.consultant.ru",
  fts: "Федеральная таможенная служба, customs.gov.ru",
  wto: "Всемирная торговая организация, www.wto.org",
};

// 1. Титул
{
  const s = pres.addSlide({ masterName: "TITLE" });
  s.addText("Основные положения Международной конвенции об упрощении и гармонизации таможенных процедур", { placeholder: "title" });
  s.addText([
    { text: "Фамилия Имя студента", options: { breakLine: true } },
    { text: "Группа ____________", options: { breakLine: true } },
    { text: "Российский университет дружбы народов имени Патриса Лумумбы", options: { breakLine: true } },
    { text: "2026" },
  ], { placeholder: "body", paraSpaceAfter: 8 });
  source(s, SRC.wco);
  s.addNotes(
    "— Представиться, назвать тему.\n" +
    "— Конвенцию часто называют Киотской — по городу, где принята первая редакция.\n" +
    "— План: история, цели, структура, принципы, связь с правом Евразийского экономического союза."
  );
}

// 2. История документа — временная линия
{
  const s = content("История документа", SRC.wco + "; " + SRC.conv);
  const lineY = 3.05;
  s.addShape(pres.shapes.LINE, { x: M, y: lineY, w: W, h: 0, line: { color: C.accent1, width: 2 }, objectName: "Ось времени" });
  const colW = W / 3;
  const pts = [
    ["1973", "Первая редакция", "Киото, 18 мая 1973 года"],
    ["1999", "Пересмотренная редакция", "принята 26 июня 1999 года, Брюссель"],
    ["2006", "Вступление в силу", "пересмотренная редакция действует с 3 февраля 2006 года"],
  ];
  pts.forEach(([year, head, desc], i) => {
    const x = M + i * colW;
    s.addText(year, { x, y: 2.15, w: colW - 0.3, h: 0.7, fontSize: 30, bold: true, color: C.accent1,
      align: "left", valign: "bottom", margin: 0, isTextBox: true, objectName: "Год" });
    s.addShape(pres.shapes.OVAL, { x, y: lineY - 0.11, w: 0.22, h: 0.22, fill: { color: C.accent1 },
      line: { color: C.accent1, width: 1 }, objectName: "Точка" });
    s.addText([
      { text: head, options: { bold: true, fontSize: F.schemeHead, breakLine: true } },
      { text: desc, options: { fontSize: F.scheme } },
    ], { x, y: 3.35, w: colW - 0.4, h: 1.3, color: C.text1, align: "left", valign: "top", margin: 0,
      isTextBox: true, objectName: "Подпись точки" });
  });
  s.addText([
    { text: "Разработчик: ", options: {} },
    { text: "Всемирная таможенная организация", options: { bold: true, color: C.accent1 } },
    { text: "     Участников: ", options: {} },
    { text: "более 130", options: { bold: true, color: C.accent1 } },
  ], { x: M, y: 5.5, w: W, h: 0.6, fontSize: F.body, color: C.text1, align: "left", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Итог" });
  s.addNotes(
    "— 1973 год: первая редакция, подписана в Киото.\n" +
    "— 1999 год: полный пересмотр под новые технологии и объёмы торговли, вступил в силу в 2006 году.\n" +
    "— Разработчик — Всемирная таможенная организация, сегодня более 130 участников."
  );
}

// 3. Цели конвенции — нумерованный список
{
  const s = content("Цели конвенции", SRC.conv + ", преамбула");
  const goals = [
    "упрощение таможенных процедур",
    "гармонизация процедур и документов между странами",
    "предсказуемость для участников внешней торговли",
    "снижение издержек бизнеса",
  ];
  goals.forEach((g, i) => {
    const y = BODY_Y + 0.2 + i * 1.05;
    s.addText(String(i + 1), { x: M, y, w: 0.6, h: 0.6, fontSize: F.title, bold: true, color: C.accent1,
      align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Номер" });
    s.addText(g, { x: M + 0.75, y, w: W - 0.75, h: 0.6, fontSize: F.body, color: C.text1,
      align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Цель" });
  });
  s.addNotes(
    "— Главное — упростить и сделать одинаковыми процедуры в разных странах.\n" +
    "— Бизнес заранее знает правила: меньше времени и затрат.\n" +
    "— Одновременно сохраняется качество таможенного контроля."
  );
}

// 4. Структура конвенции — три слоя
{
  const s = content("Структура конвенции", SRC.conv);
  const rows = [
    ["Основная часть", "общие положения и обязательства сторон", true],
    ["Генеральное приложение", "обязательные стандарты, относятся ко всем таможенным процедурам", true],
    ["Специфические приложения", "рекомендуемые стандарты и практики по отдельным процедурам: прибытие товаров, ввоз, вывоз, транзит, свободные зоны, особые режимы", false],
  ];
  const labW = 3.7, gap = 0.3, rowH = 1.25, rowGap = 0.35;
  rows.forEach(([head, desc], i) => {
    const y = BODY_Y + 0.1 + i * (rowH + rowGap);
    box(s, M, y, labW, rowH, [{ text: head, options: { bold: true, fontSize: F.schemeHead } }], { filled: true, name: "Слой" });
    box(s, M + labW + gap, y, W - labW - gap, rowH, [{ text: desc, options: {} }], { name: "Содержание слоя" });
    if (i < rows.length - 1) arrow(s, M + labW / 2, y + rowH, 0, rowGap);
  });
  s.addNotes(
    "— Три уровня: основная часть, генеральное приложение, специфические приложения.\n" +
    "— Генеральное приложение обязательно для всех участников.\n" +
    "— Специфические приложения страна принимает по выбору."
  );
}

// 5. Основные принципы — сетка 2 × 3
{
  const s = content("Основные принципы", SRC.conv + ", генеральное приложение");
  const items = [
    "прозрачность и предсказуемость процедур",
    "стандартизация и упрощение",
    "максимальное использование информационных технологий",
    "управление рисками вместо сплошного контроля",
    "координация государственных органов, включая принцип «единого окна»",
    "партнёрство с бизнесом",
  ];
  const colGap = 0.4, rowGap = 0.3, cw = (W - colGap) / 2, rh = 1.25;
  items.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * (cw + colGap), y = BODY_Y + 0.1 + row * (rh + rowGap);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: rh, fill: { color: C.background1 },
      line: { color: C.accent1, width: 1.5 }, objectName: "Принцип" });
    s.addText(String(i + 1), { x: x + 0.2, y, w: 0.5, h: rh, fontSize: F.title, bold: true, color: C.accent1,
      align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Номер" });
    s.addText(t, { x: x + 0.8, y, w: cw - 1.0, h: rh, fontSize: F.schemeHead, color: C.text1,
      align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Текст принципа" });
  });
  s.addNotes(
    "— Шесть принципов, на которых построены все стандарты.\n" +
    "— Ключевые: информационные технологии и управление рисками.\n" +
    "— «Единое окно»: документы подаются один раз, в одно место."
  );
}

// 6. Управление рисками — схема с развилкой
{
  const s = content("Управление рисками", SRC.conv + ", генеральное приложение, глава 6");
  const midY = 3.55, bh = 1.2;
  box(s, M, midY - bh / 2, 2.9, bh, headRuns("Все декларации"), { name: "Вход" });
  arrow(s, M + 2.9, midY, 0.7, 0);
  box(s, M + 3.6, midY - bh / 2, 2.9, bh, headRuns("Анализ риска"), { filled: true, name: "Анализ" });
  const forkX = M + 6.5, outX = M + 7.6, outW = W - 7.6;
  box(s, outX, BODY_Y + 0.05, outW, bh, headRuns("Высокий риск", "проверка документов и товаров"), { name: "Высокий риск" });
  box(s, outX, 4.25, outW, bh, headRuns("Низкий риск", "упрощённый выпуск"), { name: "Низкий риск" });
  arrow(s, forkX, BODY_Y + 0.05 + bh / 2, outX - forkX, midY - (BODY_Y + 0.05 + bh / 2), { flipV: true });
  arrow(s, forkX, midY, outX - forkX, 4.25 + bh / 2 - midY);
  s.addText("Товары проверяются выборочно, по критериям риска", {
    x: M, y: 5.95, w: W, h: 0.5, fontSize: F.body, color: C.text1, align: "left", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Вывод" });
  s.addNotes(
    "— Сплошной контроль заменяется выборочным.\n" +
    "— Система оценивает каждую декларацию по критериям риска.\n" +
    "— Ресурсы таможни идут на опасные поставки, добросовестные проходят быстрее."
  );
}

// 7. Контроль после выпуска — две колонки
{
  const s = content("Контроль после выпуска", SRC.conv + ", генеральное приложение, глава 6");
  const colGap = 0.6, cw = (W - colGap) / 2;
  const bw = 1.75, ag = (cw - 3 * bw) / 2, bh = 1.1, by = 2.5;
  const cols = [
    { head: "Проверка до выпуска", steps: ["Подача декларации", "Проверка", "Выпуск товара"], hi: 1,
      res: ["задержка товара на границе", "больше времени на оформление"] },
    { head: "Проверка после выпуска", steps: ["Подача декларации", "Выпуск товара", "Проверка позже, выборочно"], hi: 2,
      res: ["товар выпускается сразу", "проверка проводится позже, выборочно"] },
  ];
  cols.forEach((c, ci) => {
    const x0 = M + ci * (cw + colGap);
    s.addText(c.head, { x: x0, y: BODY_Y, w: cw, h: 0.5, fontSize: F.body, bold: true, color: C.accent1,
      align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Подход" });
    c.steps.forEach((t, i) => {
      const x = x0 + i * (bw + ag);
      box(s, x, by, bw, bh, [{ text: t, options: {} }], { filled: i === c.hi, name: "Шаг" });
      if (i < 2) arrow(s, x + bw, by + bh / 2, ag, 0);
    });
    s.addText(c.res.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i < c.res.length - 1, paraSpaceAfter: 10 } })), {
      x: x0, y: 4.0, w: cw, h: 1.4, fontSize: F.schemeHead, color: C.text1, align: "left", valign: "top",
      margin: 0, isTextBox: true, objectName: "Результат" });
  });
  s.addNotes(
    "— Слева — классика: товар ждёт на границе, пока идёт проверка.\n" +
    "— Справа — подход конвенции: сначала выпуск, потом выборочная проверка документов и учёта.\n" +
    "— Работает вместе с управлением рисками."
  );
}

// 8. Уполномоченный экономический оператор
{
  const s = content("Уполномоченный экономический оператор", SRC.code + ", статья 430; " + SRC.fts);
  bullets(s, [
    "надёжный участник внешней торговли получает упрощения",
    "с момента включения в реестр оператор относится к категории низкого уровня риска",
    "Таможенный кодекс Евразийского экономического союза: раздел VIII, глава 61, статья 430",
    "в Российской Федерации упрощения для оператора определяет законодательство о таможенном регулировании",
  ]);
  s.addNotes(
    "— Идея конвенции: доверять надёжным компаниям больше.\n" +
    "— Статус даёт низкий уровень риска — связь со слайдом об управлении рисками.\n" +
    "— Правовая основа — статья 430 Таможенного кодекса Евразийского экономического союза."
  );
}

// 9. Место конвенции в праве Евразийского экономического союза
{
  const s = content("Место конвенции в праве Евразийского экономического союза", SRC.code);
  const bw = 3.9, bh = 1.9, y = 2.4;
  box(s, M, y, bw, bh, headRuns("Конвенция", "рекомендательная база для модернизации"), { name: "Конвенция" });
  box(s, M + W - bw, y, bw, bh, headRuns("Таможенный кодекс Евразийского экономического союза", "обязателен для государств-членов"), { filled: true, name: "Кодекс" });
  const ax = M + bw, aw = W - 2 * bw;
  arrow(s, ax + 0.15, y + bh / 2, aw - 0.3, 0);
  s.addText("учитывает принципы конвенции", { x: ax + 0.15, y: y + bh / 2 - 0.55, w: aw - 0.3, h: 0.45,
    fontSize: F.scheme, color: C.text1, align: "left", valign: "bottom", margin: 0, isTextBox: true, objectName: "Подпись стрелки" });
  s.addText([
    { text: "Таможенный кодекс Евразийского экономического союза вступил в силу с " },
    { text: "1\u00A0января 2018\u00A0года", options: { bold: true, color: C.accent1 } },
  ], { x: M, y: 5.3, w: W, h: 0.6, fontSize: F.schemeHead, color: C.text1, align: "left", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Вывод" });
  s.addNotes(
    "— Конвенция сама по себе даёт ориентиры, а не прямые нормы.\n" +
    "— Её принципы перенесены в Таможенный кодекс, обязательный для всех государств-членов.\n" +
    "— Кодекс действует с 1 января 2018 года."
  );
}

// 10. Сравнение с соглашением Всемирной торговой организации
{
  const s = content("Не путать: Соглашение Всемирной торговой организации об упрощении процедур торговли", SRC.wco + "; " + SRC.wto);
  const hdr = { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: F.schemeHead };
  const lab = { bold: true, color: C.text1, fontSize: F.schemeHead };
  const cell = { color: C.text1, fontSize: F.schemeHead };
  const rows = [
    [{ text: "", options: hdr }, { text: "Конвенция", options: hdr }, { text: "Соглашение", options: hdr }],
    [{ text: "Орган", options: lab }, { text: "Всемирная таможенная организация", options: cell }, { text: "Всемирная торговая организация", options: cell }],
    [{ text: "Год принятия", options: lab }, { text: "1999", options: { ...cell, bold: true, color: C.accent1 } }, { text: "2013 (Бали)", options: { ...cell, bold: true, color: C.accent1 } }],
    [{ text: "Предмет", options: lab }, { text: "таможенные процедуры", options: cell }, { text: "все процедуры внешней торговли, включая сборы и публикацию правил", options: cell }],
  ];
  s.addTable(rows, {
    x: M, y: 2.0, w: W, colW: [2.6, (W - 2.6) / 2, (W - 2.6) / 2],
    rowH: [0.6, 0.8, 0.8, 1.1], valign: "middle", align: "left", margin: 0.12,
    border: { type: "solid", pt: 1, color: "A6A6A6" }, objectName: "Сравнение",
  });
  s.addNotes(
    "— Названия похожи, но это разные документы разных организаций.\n" +
    "— Конвенция — только таможенные процедуры; соглашение шире — вся внешняя торговля.\n" +
    "— Документы дополняют друг друга."
  );
}

// 11. Выводы — три тезиса в колонках
{
  const s = content("Выводы", "обобщение по источникам, перечисленным на слайде 12");
  const theses = [
    "Конвенция задаёт единый язык таможенных процедур",
    "Её принципы работают через риск-ориентированный контроль и партнёрство с бизнесом",
    "В Евразийском экономическом союзе эти принципы реализуются через Таможенный кодекс",
  ];
  const gap = 0.5, cw = (W - 2 * gap) / 3;
  theses.forEach((t, i) => {
    const x = M + i * (cw + gap);
    s.addText(String(i + 1), { x, y: BODY_Y + 0.3, w: cw, h: 0.7, fontSize: F.title, bold: true, color: C.accent1,
      align: "left", valign: "bottom", margin: 0, isTextBox: true, objectName: "Номер" });
    s.addText(t, { x, y: BODY_Y + 1.15, w: cw, h: 2.6, fontSize: F.body, color: C.text1,
      align: "left", valign: "top", margin: 0, isTextBox: true, objectName: "Тезис" });
  });
  s.addNotes(
    "— Конвенция — общий стандарт для таможен мира.\n" +
    "— Главные инструменты: управление рисками и доверие к надёжному бизнесу.\n" +
    "— Для нас она важна через Таможенный кодекс Евразийского экономического союза. Спасибо за внимание."
  );
}

// 12. Источники
{
  const s = content("Источники", "составлено автором");
  const items = [
    "Международная конвенция об упрощении и гармонизации таможенных процедур (Брюссель, 26 июня 1999 года): pavlodar.com/zakon/info?dok=04328",
    "Всемирная таможенная организация: www.wcoomd.org",
    "Федеральная таможенная служба: customs.gov.ru",
    "Таможенный кодекс Евразийского экономического союза: www.consultant.ru",
    "Всемирная торговая организация: www.wto.org",
  ];
  s.addText(items.map((t, i) => ({
    text: t, options: { bullet: { type: "number" }, breakLine: i < items.length - 1, paraSpaceAfter: 12 },
  })), { x: M, y: BODY_Y, w: W, h: 4.8, fontSize: F.schemeHead, color: C.text1, align: "left", valign: "top",
    margin: 0, isTextBox: true, objectName: "Список источников" });
  s.addNotes(
    "— Основной источник — текст конвенции в редакции 1999 года.\n" +
    "— Перед сдачей заменить ссылки на официальные тексты, если найдутся.\n" +
    "— Готов ответить на вопросы."
  );
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", path.resolve(OUT));
})();
