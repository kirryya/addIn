Office.onReady((info) => {
    console.log('Office.onReady called');
            if (info.host === Office.HostType.Excel) {
                console.log('Host is Excel');
                if (Office.context.ui.isDialog) {
                    console.log('Running as dialog');
                    // This instance is running inside a dialog, so just display the modal content
                    const modal = document.getElementById("settingsModal");
                    if (modal) {
                        modal.style.display = "flex"; // Ensure the modal content is visible
                        console.log('settingsModal display set to flex');
                    } else {
                        console.error('settingsModal element not found');
                    }
                } else {
                    console.log('Running as task pane');
            // This instance is running inside the task pane, so launch the dialog
            console.log('Calling openSettingsDialog from task pane');
            openSettingsDialog();
            if (Office.context.ui.closeContainer) {
                Office.context.ui.closeContainer();
            }
        }
    }
});

function openSettingsDialog(event) {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        { height: 400, width: 400, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to open modal dialog from openSettingsDialog:', asyncResult.error.message);
            } else {
                console.log('Modal dialog opened successfully from openSettingsDialog.');
            }
        }
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
