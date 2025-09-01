function openSettingsDialog() {
    Office.context.ui.displayDialogAsync("https://kirryya.github.io/addIn/dailog.html");
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);
