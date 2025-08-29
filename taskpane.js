function openSettingsDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        { height: 60, width: 40, displayInIframe: false },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    console.log("Message from dialog:", arg.message);
                    dialog.close();
                });
            }
        }
    );
}

// делаем функцию глобальной для Ribbon
window.openSettingsDialog = openSettingsDialog;
