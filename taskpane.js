Office.onReady(() => {
    // Привязываем идентификаторы (из манифеста) к реальной функции
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
    Office.actions.associate("regularPrices", regularPrices);
    Office.actions.associate("competitivePrices", competitivePrices);
});

function generalSettings(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        {height: 40, width: 40, displayInIframe: true},
        function (asyncResult) {
            const dialog = asyncResult.value;

            // Обработка сообщений из диалога
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
                if (args.message === "createSheets") {

                    // === Получение текущего времени через HTTP ===
                    let currentTime = "";
                    try {
                        const response = await fetch("https://worldtimeapi.org/api/timezone/Europe/Moscow");
                        const data = await response.json();
                        currentTime = data.datetime;
                        console.log("Текущее время:", currentTime);
                    } catch (err) {
                        console.error("Ошибка при запросе времени:", err);
                    }
                    // ============================================

                    await Excel.run(async (context) => {
                        const workbook = context.workbook;

                        // Вставляем время в ячейку A1 листа Лист1
                        let timeSheet = workbook.worksheets.getItemOrNullObject("Лист1");
                        timeSheet.load("name");
                        await context.sync();

                        if (timeSheet.isNullObject) {
                            timeSheet = workbook.worksheets.add("Лист1");
                        }

                        timeSheet.getRange("A1").values = [[currentTime]];

                        const sheetsData = [
                            {
                                name: "Ассортимент",
                                values: [["Товар", "Цена", "Количество"], ["Товар1", 100, 10], ["Товар2", 200, 5]]
                            },
                            {
                                name: "Продажи",
                                values: [["Дата", "Товар", "Количество", "Сумма"], ["01.09.2025", "Товар1", 2, 200], ["01.09.2025", "Товар2", 1, 200]]
                            },
                            {
                                name: "Цены конкурентов",
                                values: [["Конкурент", "Товар", "Цена"], ["CompA", "Товар1", 105], ["CompB", "Товар2", 195]]
                            }
                        ];

                        for (const sheet of sheetsData) {
                            const ws = workbook.worksheets.getItemOrNullObject(sheet.name);
                            ws.load("name");
                            await context.sync();

                            if (!ws.isNullObject) ws.delete();

                            const newSheet = workbook.worksheets.add(sheet.name);
                            const range = newSheet.getRangeByIndexes(0, 0, sheet.values.length, sheet.values[0].length);
                            range.values = sheet.values;
                        }

                        await context.sync();
                    });

                    // Закрыть диалог после добавления листов
                    dialog.close();
                }
            });

            event.completed();
        }
    );
}

function newTemplate(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 40, width: 50, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}

function regularPrices(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/regular-prices.html",
        { height: 92, width: 44, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}

function competitivePrices(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/competitive-prices.html",
        { height: 40, width: 50, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}
