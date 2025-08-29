Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        if (Office.context.ui.isDialog) {
            // This instance is running inside a dialog, so just display the modal content
            const modal = document.getElementById("settingsModal");
            modal.style.display = "flex"; // Ensure the modal content is visible
        } else {
            // This instance is running inside the task pane, so launch the dialog and close the task pane
            Office.context.ui.displayDialogAsync(
                'https://kirryya.github.io/addIn/taskpane.html', // Open taskpane.html as the dialog content
                { height: 50, width: 50, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error('Failed to open modal dialog:', asyncResult.error.message);
                    } else {
                        console.log('Modal dialog opened successfully.');
                    }
                }
            );
            if (Office.context.ui.closeContainer) {
                Office.context.ui.closeContainer();
            }
        }
    }
});

function openSettingsDialog(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        { height: 50, width: 50, displayInIframe: true }
    );

    // Обязательно уведомляем Excel, что функция завершена
    if (event) {
        event.completed();
    }
}

// Делаем глобальной
if (typeof window !== "undefined") {
    window.openSettingsDialog = openSettingsDialog;
}
