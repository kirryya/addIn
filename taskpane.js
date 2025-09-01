function openSettingsDialog() {
    console.log("Кнопка 'Новый шаблон' нажата!");
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 40, width: 60, displayInIframe: false }
    );
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);

console.log("Taskpane.js загружен")
