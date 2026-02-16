function createDesignAnalytics(showAlert) {
  if (showAlert === undefined) showAlert = true;

  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('ТЗ САЙТА');

  if (!src) {
    if (showAlert) SpreadsheetApp.getUi().alert('Лист "ТЗ САЙТА" не найден!');
    return;
  }

  const lastRow = src.getLastRow();
  if (lastRow <= 1) {
    if (showAlert) SpreadsheetApp.getUi().alert('Нет данных для анализа!');
    return;
  }

  // Читаем статусы (колонка A) и типы (колонка B) одним запросом
  const data = src.getRange(2, 1, lastRow - 1, 2).getValues();

  let done = 0;
  let inProgress = 0;
  let notDone = 0;

  for (let i = 0; i < data.length; i++) {
    const status = data[i][0];
    const type = data[i][1];

    // Пропускаем строки-заголовки (Блок, Страница) — у них нет статуса
    if (type === 'Блок' || type === 'Страница') continue;

    if (status === 'СДЕЛАНО') {
      done++;
    } else if (status === 'В РАБОТЕ') {
      inProgress++;
    } else {
      notDone++;
    }
  }

  const total = done + inProgress + notDone;
  if (total === 0) {
    if (showAlert) SpreadsheetApp.getUi().alert('Нет элементов для анализа!');
    return;
  }

  const progress = ((done / total) * 100).toFixed(1);
  const inProgressPct = ((inProgress / total) * 100).toFixed(1);
  const notDonePct = ((notDone / total) * 100).toFixed(1);

  let sheet = ss.getSheetByName('📊 Аналитика дизайна');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('📊 Аналитика дизайна');
  }

  const analyticsData = [
    [`📊 АНАЛИТИКА ДИЗАЙНА (${progress}%)`, '', '', `Обновлено: ${new Date().toLocaleString('ru-RU')}`],
    ['', '', '', 'Источник: ТЗ САЙТА'],
    ['', '', '', ''],
    ['Показатель', 'Количество', 'Статус', '% выполнения'],
    ['Всего элементов', total, '📊', `${progress}%`],
    ['✅ Выполнено', done, 'Готово', `${progress}%`],
    ['🔄 В работе', inProgress, 'В процессе', `${inProgressPct}%`],
    ['❌ Не сделано', notDone, 'Ожидание', `${notDonePct}%`]
  ];

  sheet.getRange(1, 1, analyticsData.length, 4).setValues(analyticsData);

  sheet.getRange(1, 1, 1, 4)
    .setBackground('#1a237e')
    .setFontColor('#fff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  sheet.getRange('B6').setBackground('#4caf50').setFontColor('#fff');
  sheet.getRange('B7').setBackground('#ff9800').setFontColor('#fff');
  sheet.getRange('B8').setBackground('#f44336').setFontColor('#fff');
  sheet.getRange('D5').setBackground('#2196f3').setFontColor('#fff');

  [150, 100, 100, 200].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  sheet.setFrozenRows(1);

  sheet.getCharts().forEach(function(c) {
    sheet.removeChart(c);
  });

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A5:B8'))
    .setPosition(1, 6, 0, 0)
    .setOption('title', `Прогресс выполнения: ${progress}%`)
    .setOption('width', 450)
    .setOption('height', 350)
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#2196f3', '#4caf50', '#ff9800', '#f44336'])
    .setOption('pieSliceText', 'percentage')
    .build();

  sheet.insertChart(chart);
  ss.setActiveSheet(sheet);

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      `✅ Аналитика создана!\n\nПрогресс: ${progress}%\nВыполнено: ${done}/${total}\nВ работе: ${inProgress}\nНе сделано: ${notDone}`
    );
  }
}

function setupAnalyticsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'updateAnalyticsTrigger'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('updateAnalyticsTrigger').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('✅ Триггер аналитики настроен!');
}

function updateAnalyticsTrigger() {
  createDesignAnalytics(false);
}
