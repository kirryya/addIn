Office.onReady(() => {
    // Привязываем идентификаторы из манифеста к функциям
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
    Office.actions.associate("regularPrices", regularPrices);
    Office.actions.associate("competitivePrices", competitivePrices);
    Office.actions.associate("KVI", KVI);
    Office.actions.associate("CTM", CTM);
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
    console.log('click')
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
        { height: 92, width: 52, displayInIframe: true },
        (result) => {
            const dialog = result.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                if (args.message === "close") {
                    dialog.close();
                }
            });
        }
    );
}

function openNewTemplate(event) {
    (async () => {
        await Excel.run(async (context) => {
            const workbook = context.workbook;

            // данные берём из файлов
            const files = [
                { name: "Цены конкурентов", path: "https://kirryya.github.io/addIn/Template2.xlsx", sheetName: "Цены конкурентов" },
                { name: "Продажи", path: "https://kirryya.github.io/addIn/Template1.xlsx", sheetName: "Продажи" },
                { name: "Ассортимент", path: "https://kirryya.github.io/addIn/Template2.xlsx", sheetName: "Ассортимент" },
            ];

            for (const file of files) {
                // если лист уже есть — удаляем
                const existing = workbook.worksheets.getItemOrNullObject(file.name);
                existing.load("name");
                await context.sync();
                if (!existing.isNullObject) {
                    workbook.worksheets.getFirst().activate(); // чтобы не удалять активный
                    existing.delete();
                    await context.sync();
                }

                // загружаем xlsx
                const resp = await fetch(file.path);
                const arrayBuffer = await resp.arrayBuffer();
                const base64 = arrayBufferToBase64(arrayBuffer);

                // вставляем конкретный лист из файла
                workbook.insertWorksheetsFromBase64(base64, {
                    sheetNamesToInsert: [file.sheetName], // имя листа внутри TemplateN.xlsx
                });
                await context.sync();
            }

            await context.sync();
        });

        if (event && typeof event.completed === "function") {
            event.completed();
        }
    })();
}

// вспомогательная функция
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
        { height: 92, width: 52, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}

function KVI(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/kvi-prices.html",
        { height: 92, width: 52, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}

function CTM(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/ctm-prices.html",
        { height: 92, width: 52, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}
