Office.onReady(() => {
    // Привязываем идентификатор insertSheets (из манифеста) к реальной функции
    Office.actions.associate("generalSettings", generalSettings);
    Office.actions.associate("newTemplate", newTemplate);
});

function generalSettings(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        { height: 40, width: 60, displayInIframe: true },
        (asyncResult) => {
            const dialog = asyncResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                if (args.message === "close") {
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
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 40, width: 60, displayInIframe: true }
    );

    if (event && typeof event.completed === "function") {
        event.completed();
    }
}


