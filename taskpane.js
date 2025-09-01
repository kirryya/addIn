Office.onReady(() => {
    // Привязываем идентификаторы (из манифеста) к реальной функции
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
});

function generalSettings(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        {height: 40, width: 50, displayInIframe: true},
        function (asyncResult) {
            const dialog = asyncResult.value;

            // Обработка сообщений из диалога
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
                if (args.message === "createSheets") {
                    await Excel.run(async (context) => {
                        const workbook = context.workbook;

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


