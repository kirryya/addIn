Office.onReady(() => {
    // Привязываем идентификаторы из манифеста к функциям
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
    Office.actions.associate("regularPrices", regularPrices);
    Office.actions.associate("competitivePrices", competitivePrices);
});

function isLicenseOk() {
    return Office.context.document.settings.get("licenseKey");
}

function saveLicenseAndContinue(callback) {
    Office.context.document.settings.set("licenseKey", "1234");
    Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            if (callback) callback();
        } else {
            console.error("Ошибка сохранения ключа:", result.error);
        }
    });
}

function generalSettings(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        {height: 44, width: 40, displayInIframe: true},
        (asyncResult) => {
            const dialog = asyncResult.value;

            // Слушаем сообщения из диалога
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                if (args.message === "licenseOk") {
                    // Сохраняем ключ в Office Settings
                    Office.context.document.settings.set("licenseKey", "1234");
                    Office.context.document.settings.saveAsync();

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
    if (!isLicenseOk()) {
        // Ключ не введён — сначала просим пользователя ввести
        Office.context.ui.displayDialogAsync(
            "https://kirryya.github.io/addIn/taskpane.html",
            { height: 44, width: 40, displayInIframe: true },
            (asyncResult) => {
                const dialog = asyncResult.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    if (args.message === "licenseOk") {
                        saveLicenseAndContinue(() => {
                            dialog.close();
                            openNewTemplate();
                        });
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
    if (!isLicenseOk()) {
        // Ключ не введён — сначала просим пользователя ввести
        Office.context.ui.displayDialogAsync(
            "https://kirryya.github.io/addIn/taskpane.html",
            { height: 44, width: 40, displayInIframe: true },
            (asyncResult) => {
                const dialog = asyncResult.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    if (args.message === "licenseOk") {
                        saveLicenseAndContinue(() => {
                            dialog.close();
                            setTimeout(() => {
                                openRegularPricesDialog();
                            }, 1000);
                        });
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
            const files = [
                { name: "Ассортимент", path: "templates/TemplateCompetitors.xlsx" },
                { name: "Продажи", path: "templates/TemplateLastPrices.xlsx" },
                { name: "Цены конкурентов", path: "templates/TemplateSales_1.xlsx" }
            ];

            for (const file of files) {

                // Проверяем, есть ли лист с таким именем, и удаляем его
                const existingSheet = workbook.worksheets.getItemOrNullObject(file.name);
                existingSheet.load("name");
                await context.sync();
                if (!existingSheet.isNullObject) {
                    existingSheet.delete();
                    await context.sync();
                }

                // Загружаем файл как ArrayBuffer
                const response = await fetch(file.path);
                const arrayBuffer = await response.arrayBuffer();

                // Конвертируем в Base64
                const base64 = arrayBufferToBase64(arrayBuffer);

                // Добавляем лист из Base64
                const newSheet = workbook.worksheets.addFromBase64(base64);
                newSheet.load("name");
                await context.sync();

                // Переименовываем лист, если нужно
                newSheet.name = file.name;
            }

            await context.sync();
        });
    })();
}

function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
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
