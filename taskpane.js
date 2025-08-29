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
