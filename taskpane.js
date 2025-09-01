function openSettingsDialog() {
    Office.context.ui.displayDialogAsync(
        "https://kirryya.github.io/addIn/dailog.html",
        { height: 60, width: 40, displayInIframe: false },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                alert("Failed to open dialog: " + asyncResult.error.message);
            } else {
                alert("Dialog opened successfully!");
            }
        }
    );
}

Office.actions.associate("openSettingsDialog", openSettingsDialog);
