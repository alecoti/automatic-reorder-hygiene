/**
 * Menu principale
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Generatore Scorte")
    .addItem("Esegui allocazione HW", "allocazioneScorteHW")
    .addItem("Esegui allocazione CONS", "allocazioneScorteCONS")
    .addToUi();
}
