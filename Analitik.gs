function createDesignAnalytics() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('ТЗ САЙТА');
  if (!src) return SpreadsheetApp.getUi().alert('Лист "ТЗ САЙТА" не найден!');
  const lastRow = src.getLastRow();
  if (lastRow <= 1) return SpreadsheetApp.getUi().alert('Нет данных для анализа!');
  const values = src.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const total = values.length;
  const done = values.filter(v => v === true).length;
  const todo = total - done;
  const progress = ((done / total) * 100).toFixed(1);
  let sheet = ss.getSheetByName('📊 Аналитика дизайна');
  sheet ? sheet.clear() : sheet = ss.insertSheet('📊 Аналитика дизайна');
  const data = [
    [`📊 АНАЛИТИКА ДИЗАЙНА (${progress}%)`, '', '', `Обновлено: ${new Date().toLocaleString('ru-RU')}`],
    ['', '', '', 'Источник: ТЗ САЙТА'],
    ['', '', '', ''],
    ['Показатель', 'Количество', 'Статус', '% выполнения'],
    ['Всего задач', total, '📊', `${progress}%`],
    ['✅ Выполнено', done, 'Готово', `✓ ${done}/${total}`],
    ['❌ В работе', todo, 'В процессе', `${((todo / total) * 100).toFixed(1)}%`]
  ];
  sheet.getRange(1, 1, data.length, 4).setValues(data);
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#1a237e')
    .setFontColor('#fff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');
  sheet.getRange('B6').setBackground('#4caf50').setFontColor('#fff');
  sheet.getRange('B7').setBackground('#f44336').setFontColor('#fff');
  sheet.getRange('D5').setBackground('#2196f3').setFontColor('#fff');
  [150, 100, 100, 200].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A5:B7'))
    .setPosition(1, 6, 0, 0)
    .setOption('title', `Прогресс выполнения: ${progress}%`)
    .setOption('width', 450)
    .setOption('height', 350)
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#4caf50', '#f44336'])
    .setOption('pieSliceText', 'percentage')
    .build();
  sheet.insertChart(chart);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(`✅ Аналитика создана!\n\nПрогресс: ${progress}%\nВыполнено: ${done}/${total}`);
}

function setupAnalyticsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateAnalyticsTrigger')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateAnalyticsTrigger').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('✅ Триггер аналитики настроен!');
}

function updateAnalyticsTrigger() {
  createDesignAnalytics();
}
