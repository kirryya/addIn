function openSettingsDialog() {
    console.log("openSettingsDialog called");
    Office.context.ui.displayDialogAsync("https://kirryya.github.io/addIn/dailog.html", function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog: ", result.error);
        } else {
            console.log("Dialog opened successfully");
        }
    });
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);
