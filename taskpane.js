// function openSettingsDialog() {
//     Office.context.ui.displayDialogAsync(
//         "https://kirryya.github.io/addIn/dailog.html",
//         { height: 60, width: 40, displayInIframe: true }
//     );
// }
//
// Office.actions.associate("openSettingsDialog", openSettingsDialog);

Office.onReady(() => {
    // Привязываем идентификатор insertSheets (из манифеста) к реальной функции
    Office.actions.associate("insertSheets", insertSheets);
});

async function insertSheets(event) {
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;

            const sheetsData = [
                {
                    name: "Ассортимент",
                    values: [
                        ["Товар", "Цена", "Количество"],
                        ["Товар1", 100, 10],
                        ["Товар2", 200, 5],
                    ]
                },
                {
                    name: "Продажи",
                    values: [
                        ["Дата", "Товар", "Количество", "Сумма"],
                        ["01.09.2025", "Товар1", 2, 200],
                        ["01.09.2025", "Товар2", 1, 200],
                    ]
                },
                {
                    name: "Цены конкурентов",
                    values: [
                        ["Конкурент", "Товар", "Цена"],
                        ["CompA", "Товар1", 105],
                        ["CompB", "Товар2", 195],
                    ]
                }
            ];

            for (const sheet of sheetsData) {
                // если лист уже есть — переиспользуем, иначе создаём
                let ws = workbook.worksheets.getItemOrNullObject(sheet.name);
                ws.load("name, isNullObject");
                await context.sync();

                if (ws.isNullObject) {
                    ws = workbook.worksheets.add(sheet.name);
                } else {
                    // очистим существующий диапазон перед записью (по желанию)
                    const used = ws.getUsedRangeOrNullObject();
                    used.load("address, isNullObject");
                    await context.sync();
                    if (!used.isNullObject) {
                        used.clear();
                    }
                }

                const rows = sheet.values.length;
                const cols = sheet.values[0].length;
                const range = ws.getRangeByIndexes(0, 0, rows, cols);
                range.values = sheet.values;
            }

            // активируем первый лист для наглядности
            workbook.worksheets.getItem(sheetsData[0].name).activate();

            await context.sync();
        });
        console.log("Листы добавлены!");
    } catch (err) {
        console.error("Ошибка при добавлении листов:", err);
    } finally {
        // ОБЯЗАТЕЛЬНО уведомляем Excel о завершении ExecuteFunction
        if (event && typeof event.completed === "function") {
            event.completed();
        }
    }
}
