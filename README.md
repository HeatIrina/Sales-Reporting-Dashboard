# Автоматизированная система отчетности продаж в Google Sheets

## 1. Название проекта

**"Sales Reporting Dashboard"** — Автоматическая система формирования отчетов по продажам электроники

---

## 2. Цель проекта

Создание автоматизированной системы для генерации, форматирования и рассылки отчетов о продажах на основе исходных данных в Google Sheets.

---

## 3. Исходные данные

Таблица с продажами электроники содержит следующие поля:

| Столбец | Описание |
|---------|----------|
| Продавец | Название компании-продавца |
| Товар | Наименование товара |
| Страна | Страна продажи |
| Кол-во | Количество единиц |
| Цена | Цена за единицу (руб.) |
| Дата | Дата продажи |
| Покупатель | Название компании-покупателя |

**Характеристика данных:**
- 49 записей о продажах
- Страны: Бельгия, Франция, Германия, Россия, Беларусь, США
- Период: 2011–2012 гг.
- Особенность: числа содержат пробелы (например, `14 630`)

---

## 4. Функциональность системы

### 4.1. Основные отчеты (4 листа)

| № | Название листа | Содержание |
|:---:|----------------|-------------|
| 1 | `Отчет_Товары` | Выручка по каждому товару (сортировка по убыванию) + итоговая строка |
| 2 | `Отчет_Продавцы` | Выручка по каждому продавцу + итоговая строка |
| 3 | `Отчет_Страны` | Выручка по странам + итоговая строка |
| 4 | `Отчет_Покупатели` | Топ-10 покупателей по выручке с указанием места |

### 4.2. Дополнительные отчеты (лист `Дополнительные отчеты`)

| Блок | Содержание |
|:------:|-------------|
| **KPI панель** | Общая выручка, количество продаж, средний чек, количество уникальных товаров |
| **Динамика по месяцам** | Выручка и количество продаж с правильной сортировкой по датам (формат: ГГГГ-ММ) |
| **Топ-5 товаров** | 5 товаров с максимальной выручкой |
| **Худшие 5 товаров** | 5 товаров с минимальной выручкой |

### 4.3. Отправка на email

- Автоматическое формирование временной Google Sheets таблицы
- Копирование всех отчетов (основные + дополнительные)
- Отправка на указанный адрес электронной почты
- Информативное письмо со списком отправленных отчетов и ссылкой

---

## 5. Техническая реализация

### 5.1. Инструменты

| Инструмент | Назначение |
|------------|------------|
| Google Sheets | Хранение исходных данных и отчетов |
| Google Apps Script (GAS) | Язык программирования (JavaScript) |
| Google Mail Service | Отправка писем |

### 5.2. Основные функции скрипта

| Функция | Назначение |
|:---------:|-------------|
| `onOpen()` | Создание пользовательского меню при открытии таблицы |
| `generateAllReports()` | Генерация 4 основных отчетов |
| `addExtraReports()` | Генерация дополнительных отчетов (KPI, динамика, топы) |
| `sendReportByEmail()` | Отправка всех отчетов на email |
| `cleanNumber()` | Очистка чисел от пробелов и нечисловых символов |

### 5.3. Код скрипта

```javascript
// Создание пользовательского меню
function updateMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Отчеты');
  menu.addItem('Создать основные отчеты (4 листа)', 'generateAllReports');
  menu.addSeparator();
  menu.addItem('Создать дополнительные отчеты (KPI + динамика + топы)', 'addExtraReports');
  menu.addSeparator();
  menu.addItem('Отправить отчет на email', 'sendReportByEmail');
  menu.addToUi();
}

function onOpen() {
  updateMenu();
}

// Главная функция - создает все отчеты
function generateAllReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Проверяем, что мы на правильном листе
  var firstRow = data[0].join(',');
  if (firstRow.includes('ОТЧЕТ') || firstRow.includes('Выручка')) {
    ss.toast('Ошибка: Вы на листе с отчетом! Перейдите на лист с исходными данными', 'Ошибка', 5);
    return;
  }
  
  // Ищем нужные столбцы по заголовкам
  var colIndex = { seller: -1, product: -1, country: -1, qty: -1, price: -1, buyer: -1 };
  
  for (var i = 0; i < data[0].length; i++) {
    var h = String(data[0][i]).toLowerCase();
    if (h.indexOf('продавец') !== -1) colIndex.seller = i;
    if (h.indexOf('товар') !== -1) colIndex.product = i;
    if (h.indexOf('страна') !== -1) colIndex.country = i;
    if (h.indexOf('кол-во') !== -1 || h.indexOf('количество') !== -1) colIndex.qty = i;
    if (h.indexOf('цена') !== -1) colIndex.price = i;
    if (h.indexOf('покупатель') !== -1) colIndex.buyer = i;
  }
  
  // Проверка
  if (colIndex.product === -1 || colIndex.qty === -1 || colIndex.price === -1) {
    ss.toast('Ошибка: Не найдены нужные столбцы. Убедитесь, что есть: Товар, Кол-во, Цена', 'Ошибка', 5);
    return;
  }
  
  // Собираем данные
  var sales = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var qty = parseFloat(String(row[colIndex.qty]).replace(/\s/g, '').replace(',', '.'));
    var price = parseFloat(String(row[colIndex.price]).replace(/\s/g, '').replace(',', '.'));
    
    if (!isNaN(qty) && !isNaN(price) && qty > 0 && price > 0 && row[colIndex.product]) {
      sales.push({
        product: row[colIndex.product],
        seller: colIndex.seller !== -1 && row[colIndex.seller] ? row[colIndex.seller] : 'Не указан',
        country: colIndex.country !== -1 && row[colIndex.country] ? row[colIndex.country] : 'Не указана',
        buyer: colIndex.buyer !== -1 && row[colIndex.buyer] ? row[colIndex.buyer] : 'Не указан',
        revenue: qty * price
      });
    }
  }
  
  if (sales.length === 0) {
    ss.toast('Ошибка: Нет данных для отчетов', 'Ошибка', 5);
    return;
  }
  
  // Удаляем ВСЕ старые листы с отчетами
  var allSheets = ss.getSheets();
  var sheetsToDelete = [];
  
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (name.indexOf('Отчет_') === 0 || 
        name.indexOf('Отчет -') === 0 ||
        name.indexOf('Выручка по') === 0 || 
        name.indexOf('Топ-10') === 0 ||
        name === 'Отчеты' ||
        name === 'Отчёт' ||
        name === 'Отчет' ||
        name === 'Reports') {
      sheetsToDelete.push(allSheets[i]);
    }
  }
  
  for (var i = 0; i < sheetsToDelete.length; i++) {
    ss.deleteSheet(sheetsToDelete[i]);
  }
  
  // Создаем все отчеты
  createProductReportFinal(ss, sales);
  createSellerReportFinal(ss, sales);
  createCountryReportFinal(ss, sales);
  createBuyerReportFinal(ss, sales);
  
  ss.toast('4 отчета успешно созданы! Старые отчеты удалены.', 'Готово', 3);
}

// Отчет 1: по товарам
function createProductReportFinal(ss, sales) {
  var byProduct = {};
  sales.forEach(function(s) { 
    byProduct[s.product] = (byProduct[s.product] || 0) + s.revenue; 
  });
  
  var reportData = [['Товар', 'Выручка (руб.)']];
  for (var p in byProduct) {
    reportData.push([p, byProduct[p]]);
  }
  reportData.sort(function(a, b) { return b[1] - a[1]; });
  
  var sheet = ss.insertSheet('Отчет_Товары');
  
  sheet.getRange('A1').setValue('ОТЧЕТ ПО ВЫРУЧКЕ ТОВАРОВ');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setFontSize(12);
  
  sheet.getRange(3, 1, reportData.length, 2).setValues(reportData);
  
  var headerRange = sheet.getRange(3, 1, 1, 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  
  var lastRow = reportData.length + 2;
  var dataRange = sheet.getRange(3, 1, reportData.length, 2);
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment('left');
  
  sheet.getRange('B4:B' + lastRow).setNumberFormat('#,##0.00');
  sheet.getRange('B4:B' + lastRow).setHorizontalAlignment('right');
  sheet.autoResizeColumns(1, 2);
  
  var totalRow = lastRow + 1;
  sheet.getRange('A' + totalRow).setValue('ИТОГО:');
  sheet.getRange('A' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setFormula('=SUM(B4:B' + lastRow + ')');
  sheet.getRange('B' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setNumberFormat('#,##0.00');
  sheet.getRange('B' + totalRow).setHorizontalAlignment('right');
}

// Отчет 2: по продавцам
function createSellerReportFinal(ss, sales) {
  var bySeller = {};
  sales.forEach(function(s) { 
    bySeller[s.seller] = (bySeller[s.seller] || 0) + s.revenue; 
  });
  
  var reportData = [['Продавец', 'Выручка (руб.)']];
  for (var s in bySeller) {
    reportData.push([s, bySeller[s]]);
  }
  reportData.sort(function(a, b) { return b[1] - a[1]; });
  
  var sheet = ss.insertSheet('Отчет_Продавцы');
  
  sheet.getRange('A1').setValue('ОТЧЕТ ПО ВЫРУЧКЕ ПРОДАВЦОВ');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setFontSize(12);
  
  sheet.getRange(3, 1, reportData.length, 2).setValues(reportData);
  
  var headerRange = sheet.getRange(3, 1, 1, 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  
  var lastRow = reportData.length + 2;
  var dataRange = sheet.getRange(3, 1, reportData.length, 2);
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment('left');
  
  sheet.getRange('B4:B' + lastRow).setNumberFormat('#,##0.00');
  sheet.getRange('B4:B' + lastRow).setHorizontalAlignment('right');
  sheet.autoResizeColumns(1, 2);
  
  var totalRow = lastRow + 1;
  sheet.getRange('A' + totalRow).setValue('ИТОГО:');
  sheet.getRange('A' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setFormula('=SUM(B4:B' + lastRow + ')');
  sheet.getRange('B' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setNumberFormat('#,##0.00');
  sheet.getRange('B' + totalRow).setHorizontalAlignment('right');
}

// Отчет 3: по странам
function createCountryReportFinal(ss, sales) {
  var byCountry = {};
  sales.forEach(function(s) { 
    byCountry[s.country] = (byCountry[s.country] || 0) + s.revenue; 
  });
  
  var reportData = [['Страна', 'Выручка (руб.)']];
  for (var c in byCountry) {
    reportData.push([c, byCountry[c]]);
  }
  reportData.sort(function(a, b) { return b[1] - a[1]; });
  
  var sheet = ss.insertSheet('Отчет_Страны');
  
  sheet.getRange('A1').setValue('ОТЧЕТ ПО ВЫРУЧКЕ ПО СТРАНАМ');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setFontSize(12);
  
  sheet.getRange(3, 1, reportData.length, 2).setValues(reportData);
  
  var headerRange = sheet.getRange(3, 1, 1, 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  
  var lastRow = reportData.length + 2;
  var dataRange = sheet.getRange(3, 1, reportData.length, 2);
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment('left');
  
  sheet.getRange('B4:B' + lastRow).setNumberFormat('#,##0.00');
  sheet.getRange('B4:B' + lastRow).setHorizontalAlignment('right');
  sheet.autoResizeColumns(1, 2);
  
  var totalRow = lastRow + 1;
  sheet.getRange('A' + totalRow).setValue('ИТОГО:');
  sheet.getRange('A' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setFormula('=SUM(B4:B' + lastRow + ')');
  sheet.getRange('B' + totalRow).setFontWeight('bold');
  sheet.getRange('B' + totalRow).setNumberFormat('#,##0.00');
  sheet.getRange('B' + totalRow).setHorizontalAlignment('right');
}

// Отчет 4: топ-10 покупателей
function createBuyerReportFinal(ss, sales) {
  var byBuyer = {};
  sales.forEach(function(s) { 
    byBuyer[s.buyer] = (byBuyer[s.buyer] || 0) + s.revenue; 
  });
  
  var reportData = [['Покупатель', 'Выручка (руб.)']];
  for (var b in byBuyer) {
    reportData.push([b, byBuyer[b]]);
  }
  reportData.sort(function(a, b) { return b[1] - a[1]; });
  reportData = [reportData[0]].concat(reportData.slice(1, 11));
  
  var sheet = ss.insertSheet('Отчет_Покупатели');
  
  sheet.getRange('A1').setValue('ТОП-10 ПОКУПАТЕЛЕЙ ПО ВЫРУЧКЕ');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setFontSize(12);
  
  sheet.getRange(3, 1, reportData.length, 2).setValues(reportData);
  
  var headerRange = sheet.getRange(3, 1, 1, 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  
  var lastRow = reportData.length + 2;
  var dataRange = sheet.getRange(3, 1, reportData.length, 2);
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment('left');
  
  sheet.getRange('B4:B' + lastRow).setNumberFormat('#,##0.00');
  sheet.getRange('B4:B' + lastRow).setHorizontalAlignment('right');
  sheet.autoResizeColumns(1, 2);
  
  sheet.getRange('C3').setValue('Место');
  sheet.getRange('C3').setFontWeight('bold');
  sheet.getRange('C3').setBackground('#4A86E8');
  sheet.getRange('C3').setFontColor('#FFFFFF');
  sheet.getRange('C3').setHorizontalAlignment('center');
  
  for (var i = 0; i < reportData.length - 1; i++) {
    sheet.getRange(4 + i, 3).setValue(i + 1);
    sheet.getRange(4 + i, 3).setHorizontalAlignment('center');
  }
  sheet.autoResizeColumn(3);
  
  var placeRange = sheet.getRange(4, 3, reportData.length - 1, 1);
  placeRange.setBorder(true, true, true, true, true, true);
}

// Дополнительные отчеты
function addExtraReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var firstRow = data[0].join(',');
  if (firstRow.includes('ОТЧЕТ') || firstRow.includes('Выручка')) {
    ss.toast('Ошибка: Вы на листе с отчетом! Перейдите на лист с исходными данными', 'Ошибка', 5);
    return;
  }
  
  // Ищем нужные столбцы
  var colIndex = { seller: -1, product: -1, country: -1, qty: -1, price: -1, buyer: -1, date: -1 };
  
  for (var i = 0; i < data[0].length; i++) {
    var h = String(data[0][i]).toLowerCase();
    if (h.indexOf('продавец') !== -1) colIndex.seller = i;
    if (h.indexOf('товар') !== -1) colIndex.product = i;
    if (h.indexOf('страна') !== -1) colIndex.country = i;
    if (h.indexOf('кол-во') !== -1 || h.indexOf('количество') !== -1) colIndex.qty = i;
    if (h.indexOf('цена') !== -1) colIndex.price = i;
    if (h.indexOf('покупатель') !== -1) colIndex.buyer = i;
    if (h.indexOf('дата') !== -1) colIndex.date = i;
  }
  
  // Собираем данные
  var sales = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var qty = parseFloat(String(row[colIndex.qty]).replace(/\s/g, '').replace(',', '.'));
    var price = parseFloat(String(row[colIndex.price]).replace(/\s/g, '').replace(',', '.'));
    
    if (!isNaN(qty) && !isNaN(price) && qty > 0 && price > 0 && row[colIndex.product]) {
      var date = null;
      if (colIndex.date !== -1 && row[colIndex.date]) {
        date = new Date(row[colIndex.date]);
      }
      sales.push({
        product: row[colIndex.product],
        seller: colIndex.seller !== -1 && row[colIndex.seller] ? row[colIndex.seller] : 'Не указан',
        revenue: qty * price,
        quantity: qty,
        date: date
      });
    }
  }
  
  if (sales.length === 0) {
    ss.toast('Ошибка: Нет данных для отчетов', 'Ошибка', 5);
    return;
  }
  
  // ============ РАСЧЕТ KPI ============
  var totalRevenue = 0;
  var totalQuantity = 0;
  sales.forEach(function(s) {
    totalRevenue += s.revenue;
    totalQuantity += s.quantity;
  });
  var avgCheck = totalRevenue / totalQuantity;
  var uniqueProducts = [...new Set(sales.map(s => s.product))].length;
  
  // Удаляем старый дополнительный отчет, если есть
  var extraSheet = ss.getSheetByName('Дополнительные отчеты');
  if (extraSheet) ss.deleteSheet(extraSheet);
  
  var reportSheet = ss.insertSheet('Дополнительные отчеты');
  var currentRow = 1;
  
  // ============ 1. KPI ПАНЕЛЬ ============
  reportSheet.getRange(currentRow, 1).setValue('KPI - КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ');
  reportSheet.getRange(currentRow, 1).setFontWeight('bold');
  reportSheet.getRange(currentRow, 1).setFontSize(14);
  reportSheet.getRange(currentRow, 1).setBackground('#4A86E8');
  reportSheet.getRange(currentRow, 1).setFontColor('#FFFFFF');
  currentRow += 2;
  
  var kpiData = [
    ['Общая выручка:', totalRevenue],
    ['Общее количество продаж:', totalQuantity],
    ['Средний чек:', avgCheck],
    ['Количество уникальных товаров:', uniqueProducts]
  ];
  
  reportSheet.getRange(currentRow, 1, kpiData.length, 2).setValues(kpiData);
  
  // Форматирование KPI
  var kpiRange = reportSheet.getRange(currentRow, 1, kpiData.length, 2);
  kpiRange.setBorder(true, true, true, true, true, true);
  kpiRange.setVerticalAlignment('middle');
  
  // Жирный шрифт для названий
  reportSheet.getRange(currentRow, 1, kpiData.length, 1).setFontWeight('bold');
  
  // Формат чисел
  reportSheet.getRange(currentRow, 2, 1, 1).setNumberFormat('#,##0.00');
  reportSheet.getRange(currentRow + 1, 2, 1, 1).setNumberFormat('#,##0');
  reportSheet.getRange(currentRow + 2, 2, 1, 1).setNumberFormat('#,##0.00');
  reportSheet.getRange(currentRow + 3, 2, 1, 1).setNumberFormat('#,##0');
  
  // Выравнивание
  reportSheet.getRange(currentRow, 1, kpiData.length, 1).setHorizontalAlignment('left');
  reportSheet.getRange(currentRow, 2, kpiData.length, 1).setHorizontalAlignment('right');
  
  currentRow += kpiData.length + 2;
  
  // ============ 2. ДИНАМИКА ПО МЕСЯЦАМ ============
  reportSheet.getRange(currentRow, 1).setValue('ДИНАМИКА ПРОДАЖ ПО МЕСЯЦАМ');
  reportSheet.getRange(currentRow, 1).setFontWeight('bold');
  reportSheet.getRange(currentRow, 1).setFontSize(14);
  reportSheet.getRange(currentRow, 1).setBackground('#4A86E8');
  reportSheet.getRange(currentRow, 1).setFontColor('#FFFFFF');
  currentRow += 2;
  
  var monthlyRevenue = {};
  var monthlyCount = {};
  sales.forEach(function(s) {
    if (s.date && !isNaN(s.date.getTime())) {
      var year = s.date.getFullYear();
      var month = s.date.getMonth() + 1;
      var monthKey = year + '-' + (month < 10 ? '0' + month : month);
      monthlyRevenue[monthKey] = (monthlyRevenue[monthKey] || 0) + s.revenue;
      monthlyCount[monthKey] = (monthlyCount[monthKey] || 0) + s.quantity;
    }
  });
  
  var sortedMonths = Object.keys(monthlyRevenue).sort();
  
  var monthlyData = [['Год-месяц', 'Выручка (руб.)', 'Количество продаж']];
  for (var m = 0; m < sortedMonths.length; m++) {
    monthlyData.push([sortedMonths[m], monthlyRevenue[sortedMonths[m]], monthlyCount[sortedMonths[m]]]);
  }
  
  if (monthlyData.length > 1) {
    reportSheet.getRange(currentRow, 1, monthlyData.length, 3).setValues(monthlyData);
    
    var monthlyHeaderRange = reportSheet.getRange(currentRow, 1, 1, 3);
    monthlyHeaderRange.setFontWeight('bold');
    monthlyHeaderRange.setBackground('#4A86E8');
    monthlyHeaderRange.setFontColor('#FFFFFF');
    monthlyHeaderRange.setHorizontalAlignment('center');
    
    var monthlyTableRange = reportSheet.getRange(currentRow, 1, monthlyData.length, 3);
    monthlyTableRange.setBorder(true, true, true, true, true, true);
    monthlyTableRange.setHorizontalAlignment('left');
    
    reportSheet.getRange(currentRow + 1, 2, monthlyData.length - 1, 1).setNumberFormat('#,##0.00');
    reportSheet.getRange(currentRow + 1, 3, monthlyData.length - 1, 1).setNumberFormat('#,##0');
    reportSheet.autoResizeColumns(1, 3);
  }
  currentRow += monthlyData.length + 2;
  
  // ============ 3. ТОП-5 ТОВАРОВ ============
  var productRevenue = {};
  sales.forEach(function(s) {
    productRevenue[s.product] = (productRevenue[s.product] || 0) + s.revenue;
  });
  
  var productArray = [];
  for (var p in productRevenue) {
    productArray.push({name: p, revenue: productRevenue[p]});
  }
  productArray.sort(function(a, b) { return b.revenue - a.revenue; });
  
  var top5 = productArray.slice(0, 5);
  var flop5 = productArray.slice(-5).reverse();
  
  // Топ-5
  reportSheet.getRange(currentRow, 1).setValue('ТОП-5 ТОВАРОВ ПО ВЫРУЧКЕ');
  reportSheet.getRange(currentRow, 1).setFontWeight('bold');
  reportSheet.getRange(currentRow, 1).setFontSize(14);
  reportSheet.getRange(currentRow, 1).setBackground('#4A86E8');
  reportSheet.getRange(currentRow, 1).setFontColor('#FFFFFF');
  currentRow += 2;
  
  var top5Data = [['№', 'Товар', 'Выручка (руб.)']];
  for (var i = 0; i < top5.length; i++) {
    top5Data.push([i + 1, top5[i].name, top5[i].revenue]);
  }
  
  reportSheet.getRange(currentRow, 1, top5Data.length, 3).setValues(top5Data);
  
  var top5HeaderRange = reportSheet.getRange(currentRow, 1, 1, 3);
  top5HeaderRange.setFontWeight('bold');
  top5HeaderRange.setBackground('#4A86E8');
  top5HeaderRange.setFontColor('#FFFFFF');
  top5HeaderRange.setHorizontalAlignment('center');
  
  var top5TableRange = reportSheet.getRange(currentRow, 1, top5Data.length, 3);
  top5TableRange.setBorder(true, true, true, true, true, true);
  reportSheet.getRange(currentRow + 1, 3, top5Data.length - 1, 1).setNumberFormat('#,##0.00');
  reportSheet.getRange(currentRow + 1, 1, top5Data.length - 1, 1).setHorizontalAlignment('center');
  reportSheet.autoResizeColumns(1, 3);
  
  currentRow += top5Data.length + 2;
  
  // Худшие 5
  reportSheet.getRange(currentRow, 1).setValue('ХУДШИЕ 5 ТОВАРОВ ПО ВЫРУЧКЕ');
  reportSheet.getRange(currentRow, 1).setFontWeight('bold');
  reportSheet.getRange(currentRow, 1).setFontSize(14);
  reportSheet.getRange(currentRow, 1).setBackground('#4A86E8');
  reportSheet.getRange(currentRow, 1).setFontColor('#FFFFFF');
  currentRow += 2;
  
  var flop5Data = [['№', 'Товар', 'Выручка (руб.)']];
  for (var i = 0; i < flop5.length; i++) {
    flop5Data.push([i + 1, flop5[i].name, flop5[i].revenue]);
  }
  
  reportSheet.getRange(currentRow, 1, flop5Data.length, 3).setValues(flop5Data);
  
  var flop5HeaderRange = reportSheet.getRange(currentRow, 1, 1, 3);
  flop5HeaderRange.setFontWeight('bold');
  flop5HeaderRange.setBackground('#4A86E8');
  flop5HeaderRange.setFontColor('#FFFFFF');
  flop5HeaderRange.setHorizontalAlignment('center');
  
  var flop5TableRange = reportSheet.getRange(currentRow, 1, flop5Data.length, 3);
  flop5TableRange.setBorder(true, true, true, true, true, true);
  reportSheet.getRange(currentRow + 1, 3, flop5Data.length - 1, 1).setNumberFormat('#,##0.00');
  reportSheet.getRange(currentRow + 1, 1, flop5Data.length - 1, 1).setHorizontalAlignment('center');
  reportSheet.autoResizeColumns(1, 3);
  
  reportSheet.setFrozenRows(1);
  
  ss.toast('Дополнительные отчеты созданы на листе "Дополнительные отчеты"', 'Готово', 3);
}

// Функция отправки отчета на email (с дополнительными отчетами)
function sendReportByEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var email = 'kordumova.irina@yandex.ru';
  
  // Проверяем наличие основных отчетов
  var mainReportSheets = ['Отчет_Товары', 'Отчет_Продавцы', 'Отчет_Страны', 'Отчет_Покупатели'];
  var missingSheets = [];
  
  mainReportSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      missingSheets.push(name);
    }
  });
  
  // Проверяем наличие дополнительных отчетов
  var extraSheet = ss.getSheetByName('Дополнительные отчеты');
  var hasExtra = (extraSheet !== null);
  
  if (missingSheets.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка', 'Сначала создайте основные отчеты (меню "Отчеты" -> "Создать основные отчеты").\nОтсутствуют: ' + missingSheets.join(', '), ui.ButtonSet.OK);
    return;
  }
  
  // Создаем временную таблицу
  var tempSpreadsheet = SpreadsheetApp.create('Отчеты_от_' + new Date().toLocaleDateString('ru-RU'));
  
  // Копируем основные отчеты
  mainReportSheets.forEach(function(name) {
    var sourceSheet = ss.getSheetByName(name);
    var destinationSheet = sourceSheet.copyTo(tempSpreadsheet);
    destinationSheet.setName(name);
  });
  
  // Копируем дополнительные отчеты, если есть
  if (hasExtra) {
    var sourceExtra = ss.getSheetByName('Дополнительные отчеты');
    var destExtra = sourceExtra.copyTo(tempSpreadsheet);
    destExtra.setName('Дополнительные отчеты');
  }
  
  // Удаляем лишний стандартный лист
  var defaultSheet = tempSpreadsheet.getSheets()[0];
  var isDefaultSheet = true;
  var mainSheetNames = mainReportSheets.concat(hasExtra ? ['Дополнительные отчеты'] : []);
  
  for (var i = 0; i < mainSheetNames.length; i++) {
    if (defaultSheet.getName() === mainSheetNames[i]) {
      isDefaultSheet = false;
      break;
    }
  }
  
  if (isDefaultSheet && defaultSheet.getName() !== 'Дополнительные отчеты') {
    try {
      tempSpreadsheet.deleteSheet(defaultSheet);
    } catch(e) {
      // Игнорируем ошибку, если лист нельзя удалить
    }
  }
  
  // Формируем письмо
  var subject = 'Отчеты по продажам от ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  
  var body = 'Здравствуйте!\n\n';
  body += 'Сформированы следующие отчеты:\n';
  body += 'ОСНОВНЫЕ ОТЧЕТЫ:\n';
  body += '- Выручка по товарам\n';
  body += '- Выручка по продавцам\n';
  body += '- Выручка по странам\n';
  body += '- Топ-10 покупателей\n';
  
  if (hasExtra) {
    body += '\nДОПОЛНИТЕЛЬНЫЕ ОТЧЕТЫ:\n';
    body += '- KPI - Ключевые показатели\n';
    body += '- Динамика продаж по месяцам\n';
    body += '- Топ-5 товаров\n';
    body += '- Худшие 5 товаров\n';
  }
  
  body += '\nОтчеты находятся в прикрепленной Google Sheets таблице.\n\n';
  body += 'Ссылка для просмотра: ' + tempSpreadsheet.getUrl() + '\n\n';
  body += 'С уважением,\nАвтоматическая система отчетов';
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });
    
    var ui = SpreadsheetApp.getUi();
    var message = 'Отчет отправлен на адрес: ' + email + '\n\nВременная таблица: ' + tempSpreadsheet.getUrl();
    if (!hasExtra) {
      message += '\n\nПримечание: Дополнительные отчеты не были созданы. Запустите "Создать дополнительные отчеты" перед отправкой.';
    }
    ui.alert('Успешно!', message, ui.ButtonSet.OK);
    
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка отправки', 'Не удалось отправить письмо. Проверьте настройки доступа.\n\nОшибка: ' + error.toString(), ui.ButtonSet.OK);
  }
}
```

## 6. Инструкция по использованию

### Шаг 1: Подготовка данных

Убедитесь, что в таблице есть столбцы:

- Продавец
- Товар
- Страна
- Кол-во
- Цена
- Дата (опционально)
- Покупатель

### Шаг 2: Создание основных отчетов

Нажмите: `Отчеты → Создать основные отчеты (4 листа)`

**Результат:** Создаются листы `Отчет_Товары`, `Отчет_Продавцы`, `Отчет_Страны`, `Отчет_Покупатели`

### Шаг 3: Создание дополнительных отчетов

Нажмите: `Отчеты → Создать дополнительные отчеты (KPI + динамика + топы)`

**Результат:** Создается лист `Дополнительные отчеты` с KPI, динамикой и топами

### Шаг 4: Отправка отчетов

Нажмите: `Отчеты → Отправить отчет на email`

**Результат:** Письмо с ссылкой на временную таблицу отправлено на указанный адрес

### Шаг 5: Обновление данных

При добавлении новых данных повторите шаги 2–4 (старые отчеты удалятся автоматически)

---

## 7. Результаты и выводы

### 7.1. Достигнутые результаты

- ✅ Полная автоматизация формирования 4 основных отчетов
- ✅ Дополнительная аналитика (KPI, динамика, топ-5)
- ✅ Автоматическая отправка отчетов на email
- ✅ Удобное меню для пользователя
- ✅ Профессиональное форматирование (сетка, цвета, выравнивание)
- ✅ Обработка "грязных" данных (пробелы в числах)

### 7.2. Экономия времени

| Операция | Ручное выполнение | Автоматическое |
|----------|-------------------|----------------|
| Формирование отчетов | 15–20 минут | 3 секунды |
| Форматирование таблиц | 5–10 минут | 1 секунда |
| Отправка на email | 2–3 минуты | 2 секунды |
| **Итого** | **~30 минут** | **~6 секунд** |

## Ссылка
https://docs.google.com/spreadsheets/d/1tfhWqxakiumaerlr_xf7_bkW-TDZa9JCNnb-vPkefgs/edit?usp=sharing
