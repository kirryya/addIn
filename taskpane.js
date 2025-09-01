Office.onReady(() => {
    // Привязываем идентификатор insertSheets к функции
    Office.actions.associate("insertSheets", insertSheets);
});

function insertSheets(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 60, width: 40, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}
