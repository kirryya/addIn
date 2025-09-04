let licenseKey = "";

Office.onReady(() => {
    // Привязываем идентификаторы из манифеста к функциям
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
    Office.actions.associate("regularPrices", regularPrices);
    Office.actions.associate("competitivePrices", competitivePrices);
});

function generalSettings(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        {height: 44, width: 40, displayInIframe: true},
        (asyncResult) => {
            const dialog = asyncResult.value;

            // Слушаем сообщения из диалога
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                if (args.message === "licenseOk") {
                    licenseKey = "1234";
                    dialog.close();
                }
            });
        }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}

function newTemplate(event) {
    if (licenseKey !== "1234") {
        // Ключ не введён — сначала просим пользователя ввести
        Office.context.ui.displayDialogAsync(
            "https://kirryya.github.io/addIn/taskpane.html",
            { height: 44, width: 40, displayInIframe: true },
            (asyncResult) => {
                const dialog = asyncResult.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    if (args.message === "licenseOk") {
                        licenseKey = "1234";
                        dialog.close();

                        // после успешного ввода сразу открываем нужный диалог
                        openNewTemplate();
                    }
                });
            }
        );
    } else {
        openNewTemplate();
    }

    if (event && typeof event.completed === "function") event.completed();
}

function regularPrices(event) {
    if (licenseKey !== "1234") {
        // Ключ не введён — сначала просим пользователя ввести
        Office.context.ui.displayDialogAsync(
            "https://kirryya.github.io/addIn/taskpane.html",
            { height: 44, width: 40, displayInIframe: true },
            (asyncResult) => {
                const dialog = asyncResult.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    if (args.message === "licenseOk") {
                        licenseKey = "1234";
                        dialog.close();

                        // после успешного ввода сразу открываем нужный диалог
                        openRegularPricesDialog();
                    }
                });
            }
        );
    } else {
        openRegularPricesDialog();
    }

    if (event && typeof event.completed === "function") event.completed();
}

function openRegularPricesDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/regular-prices.html",
        { height: 92, width: 44, displayInIframe: true }
    );
}

function openNewTemplate () {
    (async () => {
        // Получаем текущее время
        let currentTime = "";
        try {
            const response = await fetch("https://worldtimeapi.org/api/timezone/Europe/Moscow");
            const data = await response.json();
            currentTime = data.datetime;
        } catch (err) {
            console.error("Ошибка при запросе времени:", err);
        }

        await Excel.run(async (context) => {
            const workbook = context.workbook;

            // Лист1 с временем
            let timeSheet = workbook.worksheets.getItemOrNullObject("Лист1");
            timeSheet.load("name");
            await context.sync();

            if (timeSheet.isNullObject) {
                timeSheet = workbook.worksheets.add("Лист1");
            }

            timeSheet.getRange("A1").values = [[currentTime]];

            // Данные для других листов
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

        if (event && typeof event.completed === "function") {
            event.completed();
        }
    })();
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
