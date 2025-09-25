/**
 * Common.js
 * 
 * Funzioni di supporto riutilizzabili tra HW.js e CONS.js
 */

/**
 * Calcola il totale donato da una lista di allocazioni.
 * @param {Array} allocazioni - Array di coppie [magazzino, quantità].
 * @returns {number}
 */
function calcolaTotaleDonato(allocazioni) {
  return (allocazioni || []).reduce((sum, a) => sum + (a[1] || 0), 0);
}

/**
 * Calcola il totale ricevuto da una lista di allocazioni.
 * @param {Array} allocazioni - Array di coppie [magazzino, quantità].
 * @returns {number}
 */
function calcolaTotaleRicevuto(allocazioni) {
  return (allocazioni || []).reduce((sum, a) => sum + (a[1] || 0), 0);
}

/**
 * Esegue il padding di una riga fino al numero di colonne necessarie.
 * @param {Array} rowData - La riga da completare.
 * @param {number} neededCols - Numero di colonne richiesto.
 * @returns {Array}
 */
function padRow(rowData, neededCols) {
  while (rowData.length < neededCols) {
    rowData.push("");
  }
  return rowData;
}
