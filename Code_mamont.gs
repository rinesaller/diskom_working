/**
 * Code_mamont.gs — парсер сайта https://mamont.by/ для Google Apps Script
 *
 * Логика работы:
 *  1. Лист "Brands": в колонке A находятся названия брендов (например: Alpina, Knauf, Ceresit).
 *     Скрипт находит страницу бренда на сайте (https://mamont.by/brands/<slug>/),
 *     при необходимости проверяя список всех брендов с https://mamont.by/brands/.
 *  2. Со страниц бренда (пагинация ?PAGEN_1000=N) собираются карточки товаров:
 *     название, цена, старая цена, URL изображения, URL товара.
 *  3. Для каждого товара открывается детальная страница, откуда берутся
 *     описание и характеристики.
 *  4. Результаты пишутся в лист "Products" со столбцами:
 *     A: Бренд | B: Название | C: Цена | D: Старая цена | E: URL изображения |
 *     F: Описание | G: Характеристики | H: URL товара | I: Дата парсинга
 *
 * Запуск:
 *  - parseAllBrands()         — пропарсить все бренды из листа Brands
 *  - parseBrandByName("Alpina") — пропарсить один бренд
 *  - fillBrandSlugs()         — (опционально) проставить в лист Brands найденные URL страниц брендов
 */

// ============================ НАСТРОЙКИ ============================

var BASE_URL = 'https://mamont.by';
var BRANDS_SHEET = 'Brands';
var PRODUCTS_SHEET = 'Products';

// Столбцы листа Products (1-индексация)
var COL = {
  BRAND: 1,      // Бренд
  NAME: 2,       // Название
  PRICE: 3,      // Цена
  OLD_PRICE: 4,  // Старая цена
  IMAGE: 5,      // URL изображения
  DESCRIPTION: 6,// Описание
  PROPERTIES: 7, // Характеристики
  URL: 8,        // URL товара
  DATE: 9        // Дата парсинга
};

var USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

var REQUEST_PAUSE_MS = 700;   // пауза между запросами, чтобы не грузить сайт
var MAX_BRAND_PAGES = 50;     // защита от бесконечной пагинации
var MAX_PRODUCTS_PER_BRAND = 500; // лимит товаров на бренд (0 = без лимита)

// ============================ ТОЧКИ ВХОДА ============================

/** Парсить все бренды из листа "Brands" (колонка A). */
function parseAllBrands() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var brandsSheet = ss.getSheetByName(BRANDS_SHEET);
  if (!brandsSheet) throw new Error('Лист "' + BRANDS_SHEET + '" не найден');

  var lastRow = brandsSheet.getLastRow();
  if (lastRow < 1) { Logger.log('Нет брендов для парсинга'); return; }

  var names = brandsSheet.getRange(1, 1, lastRow, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s && s.toLowerCase() !== 'бренд'; });

  Logger.log('Брендов к парсингу: ' + names.length);

  for (var i = 0; i < names.length; i++) {
    try {
      parseBrandByName(names[i]);
    } catch (e) {
      Logger.log('ОШИБКА при парсинге бренда "' + names[i] + '": ' + e);
    }
  }
}

/** Парсить один конкретный бренд. */
function parseBrandByName(brandName) {
  var brandUrl = findBrandUrl(brandName);
  if (!brandUrl) {
    Logger.log('Страница бренда не найдена: ' + brandName);
    return [];
  }
  Logger.log('Парсинг бренда "' + brandName + '" -> ' + brandUrl);

  var products = scrapeBrandProducts(brandUrl);
  Logger.log('Найдено товаров: ' + products.length);

  // обогащаем каждый товар описанием/характеристиками с детальной страницы
  for (var i = 0; i < products.length; i++) {
    if (MAX_PRODUCTS_PER_BRAND && i >= MAX_PRODUCTS_PER_BRAND) break;
    try {
      var html = fetchHtml(products[i].url);
      var details = parseProductDetails(html);
      products[i].description = details.description;
      products[i].properties = details.properties;
      if (details.brand) products[i].brand = details.brand;
    } catch (e) {
      Logger.log('Не удалось получить детали: ' + products[i].url + ' — ' + e);
    }
    Utilities.sleep(REQUEST_PAUSE_MS);
  }

  writeProducts(brandName, products);
  return products;
}

/** Опционально: проставляет в лист Brands найденные URL страниц брендов (колонка B). */
function fillBrandSlugs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BRANDS_SHEET);
  var lastRow = sheet.getLastRow();
  var known = getKnownBrandPages(); // { slug: displayUrl }
  for (var r = 1; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    var url = matchBrandUrl(name, known);
    if (url) sheet.getRange(r, 2).setValue(url);
  }
}

// ============================ ПОИСК СТРАНИЦЫ БРЕНДА ============================

/** Кэш известных страниц брендов с /brands/ */
var _brandPagesCache = null;

/** Возвращает карту {slug: абсолютный URL} для всех брендов из каталога /brands/. */
function getKnownBrandPages() {
  if (_brandPagesCache) return _brandPagesCache;
  var map = {};
  try {
    var html = fetchHtml(BASE_URL + '/brands/');
    var re = /href="(\/brands\/[a-z0-9_\-]+\/)"/gi;
    var m;
    while ((m = re.exec(html)) !== null) {
      var slug = m[1].replace(/\/$/, '').split('/').pop();
      map[slug.toLowerCase()] = BASE_URL + m[1];
    }
  } catch (e) {
    Logger.log('Не удалось загрузить /brands/: ' + e);
  }
  _brandPagesCache = map;
  return map;
}

/**
 * Пытается найти URL страницы бренда по его названию.
 * Порядок: прямое предположение slug -> поиск в каталоге брендов.
 */
function findBrandUrl(brandName) {
  var guessSlug = brandToSlug_(brandName);
  // 1. Прямая проверка предполагаемого адреса
  var direct = BASE_URL + '/brands/' + guessSlug + '/';
  if (urlExists_(direct)) return direct;

  // 2. Поиск по списку всех брендов сайта
  var known = getKnownBrandPages();
  var matched = matchBrandUrl(brandName, known);
  if (matched && urlExists_(matched)) return matched;

  return null;
}

function matchBrandUrl(brandName, knownMap) {
  var norm = normalizeForMatch_(brandName);
  // точное совпадение slug
  if (knownMap[norm]) return knownMap[norm];
  // частичное совпадение (например "Knauf Индастри" -> knauf)
  var best = null, bestLen = Infinity;
  for (var slug in knownMap) {
    if (slug.indexOf(norm) === 0 && slug.length < bestLen) {
      best = knownMap[slug];
      bestLen = slug.length;
    }
  }
  return best;
}

/** "Cerezit" -> "cerezit"; транслитерацию кириллицы делаем простой заменой пробелов/символов. */
function brandToSlug_(name) {
  return normalizeForMatch_(name).replace(/[^a-z0-9_\-]/g, '');
}

function normalizeForMatch_(s) {
  return String(s)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\s+/g, '_')
    .replace(/[^\w\-\u0400-\u04FF ]/g, '')
    .trim();
}

function urlExists_(url) {
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'User-Agent': USER_AGENT },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: false
    });
    var code = resp.getResponseCode();
    if (code !== 200) return false;
    // страница бренда должна содержать карточки товаров или имя бренда
    var body = resp.getContentText('UTF-8');
    return body.indexOf('item product') !== -1 || body.indexOf('item-popular') !== -1 ||
           body.toLowerCase().indexOf('/brands/') !== -1;
  } catch (e) {
    return false;
  }
}

// ============================ ЗАГРУЗКА HTML ============================

function fetchHtml(url) {
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'User-Agent': USER_AGENT, 'Accept-Language': 'ru-RU,ru;q=0.9' },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('HTTP ' + code + ' для ' + url);
  return resp.getContentText('UTF-8');
}

// ============================ СБОР ТОВАРОВ СТРАНИЦЕЙ ============================

/**
 * Собирает все товары со страницы бренда с учётом пагинации Bitrix:
 * ?PAGEN_1000=2, ?PAGEN_1000=3 ...
 */
function scrapeBrandProducts(brandUrl) {
  var products = [];
  var seen = {};
  var page = 1;

  while (page <= MAX_BRAND_PAGES) {
    var url = page === 1 ? brandUrl : brandUrl + (brandUrl.indexOf('?') === -1 ? '?' : '&') + 'PAGEN_1000=' + page;
    var html;
    try {
      html = fetchHtml(url);
    } catch (e) {
      Logger.log('Ошибка загрузки ' + url + ': ' + e);
      break;
    }

    var items = parseProductCards(html);
    if (!items.length) break;

    var addedNew = 0;
    for (var i = 0; i < items.length; i++) {
      var key = items[i].url;
      if (seen[key]) continue;
      seen[key] = true;
      products.push(items[i]);
      addedNew++;
    }

    // если на новой странице нет ни одного нового товара — пагинация закончилась
    if (addedNew === 0) break;

    // идём дальше только если на странице есть ссылка на следующую страницу PAGEN_1000
    var hasNext = html.indexOf('PAGEN_1000=' + (page + 1)) !== -1;
    if (!hasNext) break;

    page++;
    Utilities.sleep(REQUEST_PAUSE_MS);
  }

  return products;
}

/**
 * Разбирает HTML списка товаров: блоки <div class="... item product ...">.
 * Из карточки берём: название, ссылку, цену, старую цену, картинку.
 */
function parseProductCards(html) {
  var result = [];
  // режем документ на куски по началу карточки товара
  var parts = html.split(/<div class="[^"]*item[^"]*product[^"]*"/);
  for (var p = 1; p < parts.length; p++) {
    var chunk = parts[p];

    var nameM = chunk.match(/class="name"><span class="middle">([\s\S]*?)<\/span><\/a>/);
    var urlM = chunk.match(/href="(\/[^"]+?\/)" class="name"/);
    var priceM = chunk.match(/<span class="j-price"[^>]*>([\d.,]+)<\/span>/);
    var oldM = chunk.match(/<s class="discount">\s*([\d.,]+)\s*руб/);
    var imgM = chunk.match(/data-src="(\/upload\/[^"]+)"/);

    if (!urlM && !nameM) continue;

    var name = nameM ? decodeHtmlEntities_(stripTags_(nameM[1])) : '';
    var relUrl = urlM ? urlM[1] : '';
    if (!name || !relUrl) continue;

    result.push({
      name: name,
      url: BASE_URL + relUrl,
      price: priceM ? toNumber_(priceM[1]) : '',
      oldPrice: oldM ? toNumber_(oldM[1]) : '',
      image: imgM ? BASE_URL + imgM[1] : '',
      description: '',
      properties: ''
    });
  }
  return result;
}

// ============================ ДЕТАЛЬНАЯ СТРАНИЦА ТОВАРА ============================

/**
 * Разбирает страницу товара:
 *  - описание: <div class="changeDescription" data-first-value='...'> либо блок <div id="detailText">
 *  - характеристики: пары propertyName/propertyValue
 *  - бренд: ссылка /brands/<slug>/ рядом с классом brandImage
 */
function parseProductDetails(html) {
  var out = { description: '', properties: '', brand: '' };

  // --- Бренд ---
  var brandM = html.match(/<a href="\/brands\/([^"\/]+)\/" class="brandImage">/);
  if (brandM) out.brand = brandM[1];

  // --- Описание ---
  var descM = html.match(/<div class="changeDescription" data-first-value='([\s\S]*?)'>/);
  if (!descM) {
    descM = html.match(/<div id="detailText">[\s\S]*?<\/h2>([\s\S]*?)<\/div>\s*<\/div>/);
  }
  if (descM) {
    out.description = cleanTextBlock_(descM[1]);
  }

  // --- Характеристики ---
  var props = [];
  var propRe = /<div class="propertyName">([\s\S]*?)<\/div>\s*<div class="propertyValue">([\s\S]*?)<\/div>/g;
  var pm;
  while ((pm = propRe.exec(html)) !== null) {
    var pname = cleanTextBlock_(pm[1]);
    var pval = cleanTextBlock_(pm[2]);
    if (pname && pval) props.push(pname + ': ' + pval);
  }
  out.properties = props.join('\n');

  return out;
}

// ============================ ЗАПИСЬ В ЛИСТ ============================

/**
 * Добавляет строки в лист "Products". Если листа нет — создаёт с нужной шапкой.
 * Дубликаты (по URL товара) не добавляются повторно, а обновляются.
 */
function writeProducts(brandName, products) {
  if (!products.length) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRODUCTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRODUCTS_SHEET);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Бренд', 'Название', 'Цена', 'Старая цена', 'URL изображения',
      'Описание', 'Характеристики', 'URL товара', 'Дата парсинга'
    ]]);
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  var urlCol = sheet.getRange(2, COL.URL, Math.max(lastRow - 1, 0), 1).getValues();
  var rowIndexByUrl = {};
  for (var r = 0; r < urlCol.length; r++) {
    if (urlCol[r][0]) rowIndexByUrl[String(urlCol[r][0])] = r + 2;
  }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Minsk', 'yyyy-MM-dd HH:mm:ss');
  var updates = [];   // [rowIndex, rowValues]
  var appends = [];

  for (var i = 0; i < products.length; i++) {
    var pr = products[i];
    var values = [
      pr.brand ? capitalizeFirst_(pr.brand) : brandName, // Бренд
      pr.name,                                            // Название
      pr.price,                                           // Цена
      pr.oldPrice,                                        // Старая цена
      pr.image,                                           // URL изображения
      pr.description,                                     // Описание
      pr.properties,                                      // Характеристики
      pr.url,                                             // URL товара
      now                                                 // Дата парсинга
    ];
    if (rowIndexByUrl.hasOwnProperty(pr.url)) {
      updates.push([rowIndexByUrl[pr.url], values]);
    } else {
      appends.push(values);
    }
  }

  if (appends.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 9).setValues(appends);
  }
  for (var u = 0; u < updates.length; u++) {
    sheet.getRange(updates[u][0], 1, 1, 9).setValues([updates[u][1]]);
  }

  Logger.log('Записано/обновлено строк: ' + (appends.length + updates.length));
}

// ============================ УТИЛИТЫ ============================

function stripTags_(s) {
  return String(s).replace(/<[^>]+>/g, ' ');
}

/** Приводит HTML-фрагмент к читаемому тексту: теги -> переводы строк/пробелы. */
function cleanTextBlock_(htmlFragment) {
  var s = String(htmlFragment);
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<\/(li|p|div|tr|h[1-6]|ul)>/gi, '\n');
  s = s.replace(/<li[^>]*>/gi, '- ');
  s = s.replace(/<[^>]+>/g, ' ');
  s = decodeHtmlEntities_(s);
  s = s.replace(/[ \t\u00A0]+/g, ' ')          // схлопываем пробелы
       .replace(/ ?\n ?/g, '\n')                // чистим вокруг переносов
       .replace(/\n{2,}/g, '\n')
       .trim();
  // ячеек Sheets ограничен ~50k символов
  if (s.length > 49000) s = s.substring(0, 49000) + '…';
  return s;
}

function decodeHtmlEntities_(s) {
  s = String(s);
  var map = {
    '&nbsp;': ' ', '&quot;': '"', '&laquo;': '«', '&raquo;': '»',
    '&mdash;': '—', '&ndash;': '–', '&hellip;': '…', '&amp;': '&',
    '&lt;': '<', '&gt;': '>', '&#39;': "'", '&apos;': "'"
  };
  s = s.replace(/&(nbsp|quot|laquo|raquo|mdash|ndash|hellip|amp|lt|gt|apos);/g, function (m, g) {
    return map[m];
  });
  s = s.replace(/&#(\d+);/g, function (m, dec) {
    return String.fromCharCode(dec);
  });
  return s;
}

function toNumber_(s) {
  var n = parseFloat(String(s).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? '' : n;
}

function capitalizeFirst_(s) {
  s = String(s);
  return s.charAt(0).toUpperCase() + s.slice(1);
}

// ============================ ОТЛАДОЧНЫЕ ФУНКЦИИ ============================

/** Быстрая проверка: что нашлось по одному бренду (без записи в лист). */
function testParseBrand() {
  var brandUrl = findBrandUrl('Alpina');
  Logger.log('brandUrl: ' + brandUrl);
  if (!brandUrl) return;
  var products = scrapeBrandProducts(brandUrl);
  Logger.log('Товаров: ' + products.length);
  var html = fetchHtml(products[0].url);
  var det = parseProductDetails(html);
  Logger.log('Пример: ' + JSON.stringify({
    name: products[0].name,
    price: products[0].price,
    oldPrice: products[0].oldPrice,
    image: products[0].image,
    brand: det.brand,
    descPreview: det.description.substring(0, 200),
    propsPreview: det.properties.substring(0, 200)
  }, null, 2));
}
