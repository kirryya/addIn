function openSettingsDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/taskpane.html",
        { height: 40, width: 60, displayInIframe: false }
    );
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);
