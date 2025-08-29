function openSettingsDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html", // URL вашего HTML
        { height: 50, width: 50, displayInIframe: true }, // можно регулировать размер
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                // dialog получаем через asyncResult.value
                const dialog = asyncResult.value;

                // Можно слушать сообщение из диалога
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    console.log("Message from dialog:", arg.message);
                    dialog.close(); // закрыть диалог при необходимости
                });
            }
        }
    );
}

window.openSettingsDialog = openSettingsDialog;
