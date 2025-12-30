Office.onReady(() => {
  console.log("Commands ready");
});

function askFromRibbon() {
  const settings = Office.context.roamingSettings;

  if (!settings.get("isConnected")) {
    Office.addin.show(); // ouvre la TaskPane
    return;
  }

  runExcelLogic();
}

async function runExcelLogic() {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.values = [["Hello depuis le Ruban 🚀"]];
  });
}
