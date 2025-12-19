Office.onReady(() => {
  document.getElementById("login").onclick = openLogin;
});

function openLogin() {
  Office.context.ui.displayDialogAsync(
    "https://triskell-outlook-add-into-timesheet-crdd.onrender.com/dialog.html",
    { height: 60, width: 30 },
    result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(result.error);
        return;
      }

      const dialog = result.value;

      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        arg => {
          console.log("Auth result:", arg.message);
          dialog.close();
        }
      );
    }
  );
}
