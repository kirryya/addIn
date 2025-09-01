// function openSettingsDialog() {
//     Office.context.ui.displayDialogAsync(
//         "https://kirryya.github.io/addIn/dailog.html",
//         { height: 60, width: 40, displayInIframe: true }
//     );
// }
//
// Office.actions.associate("openSettingsDialog", openSettingsDialog);

Office.onReady((info) => {
    console.log("Office ready", info);

    // Регистрируем действие для кнопки
    Office.actions.associate("openSettingsDialog", insertSheets);
});

async function insertSheets() {
    await Excel.run(async (context) => {
        const workbook = context.workbook;

        // Данные для листов
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

        // Создаем листы и заполняем данными
        sheetsData.forEach(sheet => {
            const newSheet = workbook.worksheets.add(sheet.name);
            const range = newSheet.getRangeByIndexes(
                0,
                0,
                sheet.values.length,
                sheet.values[0].length
            );
            range.values = sheet.values;
        });

        await context.sync();
    });

    console.log("Листы добавлены!");
}
