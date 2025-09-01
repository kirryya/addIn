function openSettingsDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 60, width: 40, displayInIframe: true }
    );
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);
