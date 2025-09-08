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

function openNewTemplate (event) {
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

            // локальные файлы
            const files = [
                { name: "Ассортимент", path: "./Template1.xlsx", sheetName: "Лист1" },
                { name: "Продажи", path: "./Template2.xlsx", sheetName: "Лист1" },
                { name: "Цены конкурентов", path: "./Template3.xlsx", sheetName: "Лист1" }
            ];

            for (const file of files) {
                // удаляем старый лист, если есть
                const existing = workbook.worksheets.getItemOrNullObject(file.name);
                existing.load("name");
                await context.sync();
                if (!existing.isNullObject) {
                    workbook.worksheets.getFirst().activate(); // чтобы не удалять активный
                    existing.delete();
                    await context.sync();
                }

                // загружаем локальный xlsx
                const resp = await fetch(file.path);
                const arrayBuffer = await resp.arrayBuffer();
                const base64 = arrayBufferToBase64(arrayBuffer);

                // вставляем лист из файла
                workbook.insertWorksheetsFromBase64(base64, {
                    sheetNamesToInsert: [file.sheetName], // имя листа в шаблоне
                    positionType: Excel.InsertWorksheetPositionType.end
                });
                await context.sync();

                // сразу переименовываем вставленный лист
                const inserted = workbook.worksheets.getItem(file.sheetName);
                inserted.name = file.name;
                await context.sync();
            }
        });

        if (event && typeof event.completed === "function") {
            event.completed();
        }
    })();
}

// функция конвертации ArrayBuffer → Base64
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize);
        binary += String.fromCharCode.apply(null, chunk);
    }
    return btoa(binary);
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
