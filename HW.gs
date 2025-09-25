/**
 * Funzione principale per HW
 */
/**
 * Funzione principale per HW
 */
function allocazioneScorteHW() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startRow = 1;        // Riga intestazioni principali
  const startDataRow = 2;    // Prima riga dati
  const startCol = 81;       // Colonna CC (indice 81)

  // Pulizia eventuali colonne oltre startCol
  const lastCol = sheet.getLastColumn();
  if (lastCol >= startCol) {
    sheet.deleteColumns(startCol, lastCol - startCol + 1);
  }

  // === DEBUG: controllo quante righe HW ci sono ===
  let countHW = 0;
  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];      // Colonna D
    const ubic = data[i][1];        // Colonna B
    const tipo = data[i][12];       // Colonna M
    const fabU = parseFloat(data[i][69]); // Colonna BR
    const fabS = parseFloat(data[i][70]); // Colonna BS
    if (tipo === "HW") {
      countHW++;
      Logger.log(
        "Riga %s | Codice: %s | Ubic: %s | Tipo: %s | fabU: %s | fabS: %s",
        i + 1, codice, ubic, tipo, fabU, fabS
      );
    }
  }
  Logger.log("Totale righe con HW trovate: %s", countHW);

  // === STEP 1: Generazione disponibilità da donatori (HW) ===
  const disponibilita = generaDisponibilitaHW(data, startDataRow);
  Logger.log("Disponibilità generate: U=%s, S=%s",
    JSON.stringify(disponibilita.disponibilitaU),
    JSON.stringify(disponibilita.disponibilitaS)
  );

  // === STEP 2: Allocazione fabbisogni (HW) ===
  const {
    risultati,
    trasferimenti,
    bwRicalcolato,
    maxAllocazioni
  } = allocaFabbisogniHW(data, startDataRow, disponibilita);

  Logger.log("Totale risultati generati: %s", risultati.length);
  Logger.log("Totale trasferimenti generati: %s", trasferimenti.length);
  Logger.log("Max allocazioni: %s", maxAllocazioni);

  // === STEP 3: Scrittura risultati (HW) ===
  scriviRisultatiHW(sheet, startRow, startDataRow, startCol, data, risultati, bwRicalcolato, maxAllocazioni);

  // === STEP 4: Scrittura trasferimenti (HW) ===
  scriviTrasferimentiHW(ss, trasferimenti);

  // === STEP 5: Report acquisti (HW) ===
  scriviReportAcquistiHW();

  // === STEP 6: Scrittura valore ordine (BX × BI in BY) ===
  scriviValoreOrdineHW(sheet, startDataRow);
}



/**
 * Crea mappe delle disponibilità donatori (HW).
 */
function generaDisponibilitaHW(data, startDataRow) {
  let disponibilitaU = {}; // Usati
  let disponibilitaS = {}; // Stock

  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];            // Colonna D (Articolo)
    const ubic = data[i][1];              // Colonna B (Ubicazione)
    const fabU = parseFloat(data[i][69]); // Colonna BR
    const fabS = parseFloat(data[i][70]); // Colonna BS

    if (!codice) continue;

    if (!isNaN(fabU) && fabU < 0) {
      if (!disponibilitaU[codice]) disponibilitaU[codice] = {};
      disponibilitaU[codice][ubic] = (disponibilitaU[codice][ubic] || 0) + Math.abs(fabU);
    }

    if (!isNaN(fabS) && fabS < 0) {
      if (!disponibilitaS[codice]) disponibilitaS[codice] = {};
      disponibilitaS[codice][ubic] = (disponibilitaS[codice][ubic] || 0) + Math.abs(fabS);
    }
  }

  return { disponibilitaU, disponibilitaS };
}

/**
 * Alloca i fabbisogni per HW con ottimizzazione spedizioni.
 */
/**
 * Alloca i fabbisogni per HW con logica "spedizioni minime".
 */
/**
 * Alloca i fabbisogni per HW con logica ottimizzata:
 * - Per ogni articolo, prima i riceventi con fabbisogno maggiore
 * - Si usano i donor più grandi (preferendo Usati a Stock in caso di parità)
 */

// BX-driven: i riceventi sono solo le righe HW con BX (col 76) > 0.
// Se una riga viene coperta in parte da trasferimenti e in parte da acquisti,
// l'azione viene marcata come "COMBINATO".
// BX-driven: i riceventi sono solo le righe HW con BX (col 76) > 0.
// Se una riga viene coperta in parte da trasferimenti e in parte da acquisti,
// l'azione viene marcata come "COMBINATO".
function allocaFabbisogniHW(data, startDataRow, disponibilita) {
  let trasferimenti = [];
  let bwRicalcolato = {};
  let risultati = [];
  let maxAllocazioni = 0;

  const { disponibilitaU, disponibilitaS } = disponibilita;

  // === 1) Costruisco i riceventi per articolo (solo HW con BX>0) ===
  let riceventiPerArticolo = {};
  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];                // D - Articolo
    const ubic = data[i][1];                  // B - Ubicazione
    const tipo = data[i][12];                 // M - Categoria
    const copertura = parseFloat(data[i][11]); // L - Copertura
    const ord = parseFloat(data[i][75]);      // BX - Quantità Ordine

    if (!codice || tipo !== "HW") continue;
    if (isNaN(ord) || ord <= 0) continue;     // ricevente solo se BX>0

    if (!riceventiPerArticolo[codice]) riceventiPerArticolo[codice] = [];
    riceventiPerArticolo[codice].push({
      index: i,
      codice,
      ubic,
      ord,
      copertura
    });
  }

  // === 2) Alloco: donor più capienti prima; preferisco Usati (_U) a parità ===
  for (let codice in riceventiPerArticolo) {
    // serve prima chi ha BX più alto
    let riceventi = riceventiPerArticolo[codice].sort((a, b) => (b.ord - a.ord));

    for (let rec of riceventi) {
      const key = rec.codice + "|" + rec.ubic;
      let fabBisogno = rec.ord;         // bisogno = BX
      let allocazioni = [];
      let azione = "";
      let motivo = "";
      let acquisto = 0;

      if (bwRicalcolato[key] === undefined) bwRicalcolato[key] = fabBisogno;

      // === NUOVO CHECK: copertura nazionale (colonna L) ===
      if (rec.copertura < 100) {
        azione = "ACQUISTARE";
        motivo = "STOCK NAZ <100 COPERTURA";
        acquisto = fabBisogno;
        bwRicalcolato[key] = 0;
      } else {
        // Donor disponibili per questo articolo (escludo stessa ubicazione)
        let donors = [];
        if (disponibilitaU[rec.codice]) {
          for (let [mag, qty] of Object.entries(disponibilitaU[rec.codice])) {
            if (mag !== rec.ubic && qty > 0) donors.push({ mag, qty, tipo: "_U" });
          }
        }
        if (disponibilitaS[rec.codice]) {
          for (let [mag, qty] of Object.entries(disponibilitaS[rec.codice])) {
            if (mag !== rec.ubic && qty > 0) donors.push({ mag, qty, tipo: "" });
          }
        }

        // Ordino donor: quantità desc; a parità preferisco Usati
        donors.sort((a, b) => {
          if (b.qty === a.qty) return (a.tipo === "_U" ? -1 : 1);
          return b.qty - a.qty;
        });

        // Trasferimenti greedy
        for (let donor of donors) {
          if (fabBisogno <= 0) break;
          const prelievo = Math.min(fabBisogno, donor.qty);
          if (prelievo <= 0) continue;

          allocazioni.push([donor.mag, prelievo, donor.tipo]);
          if (donor.tipo === "_U") {
            disponibilitaU[rec.codice][donor.mag] -= prelievo;
          } else {
            disponibilitaS[rec.codice][donor.mag] -= prelievo;
          }

          const donorKey = rec.codice + "|" + donor.mag;
          bwRicalcolato[donorKey] = (bwRicalcolato[donorKey] || 0) + prelievo;
          bwRicalcolato[key] -= prelievo;

          fabBisogno -= prelievo;
          trasferimenti.push([rec.codice + donor.tipo, donor.mag, rec.ubic, prelievo]);
        }

        // Residuo → acquisto; se c'è anche almeno un trasferimento diventa COMBINATO
        if (fabBisogno > 0) {
          acquisto = fabBisogno;
          azione = (allocazioni.length > 0) ? "COMBINATO" : "ACQUISTARE";
          motivo = (allocazioni.length > 0) ? "COMBINATO" : "Residuo non coperto (BX)";
          bwRicalcolato[key] = 0;
        } else {
          azione = (allocazioni.length > 0) ? "TRASFERIMENTO" : "";
        }
      }

      maxAllocazioni = Math.max(maxAllocazioni, allocazioni.length);

      risultati[rec.index] = {
        ubic: rec.ubic,
        codice: rec.codice,
        key,
        bwRicalcolato: bwRicalcolato[key],
        allocazioni,
        azione,
        motivo,
        acquisto
      };
    }
  }

  return { risultati, trasferimenti, bwRicalcolato, maxAllocazioni };
}


/**
 * Scrive sul foglio principale (solo HW).
 */

/**
 * Scrive i risultati HW in modo ottimizzato:
 * - costruisce tutte le righe in memoria
 * - scrive tutto con un'unica operazione setValues
 * - usa setFormulaR1C1 per applicare la formula "Valore Ordine" su tutte le righe
 * - logga ogni passaggio importante
 */


function scriviRisultatiHW(sheet, startRow, startDataRow, startCol, data, risultati, bwRicalcolato, maxAllocazioni) {
  const headers = [
    "Totale Donato", "Totale Ricevuto", "BW Ricalcolato",
    "Azione", "Motivo", "Quantità Acquisto", "Valore Ordine"
  ];
  let allocHeaders = [];
  for (let j = 0; j < maxAllocazioni; j++) {
    allocHeaders.push("Source " + (j + 1));
    allocHeaders.push("Qty " + (j + 1));
    allocHeaders.push("Tipo " + (j + 1));
  }

  const neededCols = headers.length + allocHeaders.length;

  // Assicuro che esistano abbastanza colonne a destra di startCol per scrivere l'output
  const lastCol = sheet.getLastColumn();
  const missing = (startCol + neededCols - 1) - lastCol;
  if (missing > 0) {
    sheet.insertColumnsAfter(lastCol, missing);
  }

  // Intestazioni
  sheet.getRange(startRow, startCol, 1, neededCols)
    .setValues([headers.concat(allocHeaders)]);

  // Costruzione output riga per riga in base ai dati del foglio,
  // garantendo l'allineamento con gli indici di "risultati"
  let out = [];
  for (let i = startDataRow - 1; i < data.length; i++) {
    const r = risultati[i];

    if (!r || !r.ubic) {
      out.push(new Array(neededCols).fill("")); // Riga vuota
      continue;
    }

    const allocazioni = r.allocazioni || [];
    const totDonato = allocazioni.reduce((sum, a) => sum + (a[1] || 0), 0);
    const totRicevuto = allocazioni.reduce((sum, a) => sum + (a[1] || 0), 0);
    const bwVal = (allocazioni.length > 0 || r.acquisto) ? (r.bwRicalcolato ?? "") : "";

    let rowData = [
      totDonato || "",
      totRicevuto || "",
      bwVal,
      r.azione || "",
      r.motivo || "",
      r.acquisto || "",   // Quantità Acquisto
      ""                  // Valore Ordine → formula inserita dopo
    ];

    // Aggiungo allocazioni
    allocazioni.forEach(a => rowData.push(a[0], a[1], a[2]));
    while (rowData.length < neededCols) rowData.push("");

    out.push(rowData);
  }

  // Scrittura massiva
  const range = sheet.getRange(startDataRow, startCol, out.length, neededCols);
  range.setValues(out);
  range.setHorizontalAlignment("center");

  // Auto resize colonne
  for (let col = startCol; col < startCol + neededCols; col++) {
    sheet.autoResizeColumn(col);
  }

  // Inserimento formula "Valore Ordine" (BY = Quantità Acquisto × Prezzo Unitario BI)
  const lastRowOut = startDataRow + out.length - 1;
  const colAzione = startCol + 3;      // Azione
  const colQta = startCol + 5;         // Quantità Acquisto
  const colValore = startCol + 6;      // Valore Ordine
  const colPrezzo = 61;                // Colonna BI (prezzo unitario)

  for (let r = startDataRow; r <= lastRowOut; r++) {
    const formula = `=IF($${colToLetter(colAzione)}${r}="ACQUISTARE", $${colToLetter(colQta)}${r}*$${colToLetter(colPrezzo)}${r}, "")`;
    sheet.getRange(r, colValore).setFormula(formula);
  }
}



/**
 * Utility: converte numero colonna → lettera colonna
 */
function colToLetter(col) {
  let temp, letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

/**
 * Utility: converte numero colonna → lettera colonna
 */
function colToLetter(col) {
  let temp, letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
/**
 * Scrive i trasferimenti HW
 */
function scriviTrasferimentiHW(ss, trasferimenti) {
  let tSheet = ss.getSheetByName("TrasferimentiHW");
  if (!tSheet) {
    tSheet = ss.insertSheet("TrasferimentiHW");
  } else {
    tSheet.clear();
  }

  tSheet.getRange(1, 1, 1, 4).setValues([["Articolo", "Da", "A", "Quantità"]]);

  if (trasferimenti.length > 0) {
    const range = tSheet.getRange(2, 1, trasferimenti.length, 4);
    range.setValues(trasferimenti);
    range.setHorizontalAlignment("center");

    // Ordinamento
    const lastRow = tSheet.getLastRow();
    tSheet.getRange(2, 1, lastRow - 1, 4).sort([
      { column: 2, ascending: true },
      { column: 3, ascending: true },
      { column: 1, ascending: true }
    ]);

    // Colori alternati per Da->A
    const values = tSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let colors = [];
    let currentCombo = null;
    let useGreen = true;

    for (let i = 0; i < values.length; i++) {
      const combo = values[i][1] + "->" + values[i][2];
      if (combo !== currentCombo) {
        currentCombo = combo;
        useGreen = !useGreen;
      }
      let color = useGreen ? "#e6f4ea" : "#e6f0f9";
      colors.push([color, color, color, color]);
    }

    tSheet.getRange(2, 1, lastRow - 1, 4).setBackgrounds(colors);

    // Report
    let totals = {};
    let combinazioni = {};

    values.forEach(([cod, da, a, qty]) => {
      if (!totals[da]) totals[da] = { donato: 0, ricevuto: 0 };
      if (!totals[a]) totals[a] = { donato: 0, ricevuto: 0 };
      totals[da].donato += qty;
      totals[a].ricevuto += qty;

      const combo = da + "->" + a;
      if (!combinazioni[combo]) {
        combinazioni[combo] = { pezzi: 0, articoli: new Set() };
      }
      combinazioni[combo].pezzi += qty;
      combinazioni[combo].articoli.add(cod);
    });

    let rSheet = ss.getSheetByName("ReportTrasferimentiHW");
    if (!rSheet) {
      rSheet = ss.insertSheet("ReportTrasferimentiHW");
    } else {
      rSheet.clear();
    }

    let summary = [["Magazzino", "Totale Donato", "Totale Ricevuto"]];
    for (let mag in totals) {
      summary.push([mag, totals[mag].donato, totals[mag].ricevuto]);
    }
    rSheet.getRange(1, 1, summary.length, 3).setValues(summary);

    let comboReport = [["Da", "A", "Totale Pezzi", "Codici Unici"]];
    for (let combo in combinazioni) {
      const [da, a] = combo.split("->");
      comboReport.push([da, a, combinazioni[combo].pezzi, combinazioni[combo].articoli.size]);
    }

    const startCombo = summary.length + 2;
    rSheet.getRange(startCombo, 1, comboReport.length, 4).setValues(comboReport);
  }
}

/**
 * Report degli acquisti HW
 */
function scriviReportAcquistiHW() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startDataRow = 2; 
  let rSheet = ss.getSheetByName("ReportAcquistiHW");
  if (!rSheet) {
    rSheet = ss.insertSheet("ReportAcquistiHW");
  } else {
    rSheet.clear();
  }

  let acquistiPerLoc = {};
  let totaleGenerale = 0;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];    
    const ubic = data[i][1];      
    const qta = parseFloat(data[i][75]); 
    const prezzo = parseFloat(data[i][60]); 

    if (!codice || isNaN(qta) || qta <= 0 || isNaN(prezzo)) continue;

    const totale = qta * prezzo;
    totaleGenerale += totale;

    if (!acquistiPerLoc[ubic]) acquistiPerLoc[ubic] = [];
    acquistiPerLoc[ubic].push([codice, qta, prezzo, totale]);
  }

  let output = [];
  let headers = ["Articolo", "Quantità", "Prezzo Unitario", "Totale Riga"];

  for (let loc in acquistiPerLoc) {
    output.push([`📦 Location: ${loc}`, "", "", ""]);
    output.push(headers);

    let totaleLoc = 0;
    acquistiPerLoc[loc].forEach(row => {
      output.push(row);
      totaleLoc += row[3];
    });

    output.push(["", "", "Totale Location", totaleLoc]);
    output.push(["", "", "", ""]);
  }

  output.push(["", "", "Totale Generale", totaleGenerale]);

  rSheet.getRange(1, 1, output.length, 4).setValues(output);
  rSheet.getRange(1, 1, output.length, 4).setHorizontalAlignment("center");

  for (let col = 1; col <= 4; col++) {
    rSheet.autoResizeColumn(col);
  }
}

/**
 * Scrive la colonna "Valore Ordine" (BY = BX × BI)
 */
function scriviValoreOrdineHW(sheet, startDataRow) {
  Logger.log("[HW] === INIZIO scriviValoreOrdineHW ===");

  try {
    const lastRow = sheet.getLastRow();
    const colQta = 76; // BX
    const colPrezzo = 61; // BI
    const colValore = colQta + 1; // BY

    Logger.log("[HW] Ultima riga dati: %s", lastRow);
    Logger.log("[HW] Colonne -> Qta: %s | Prezzo: %s | Valore: %s", colQta, colPrezzo, colValore);

    // intestazione
    sheet.getRange(1, colValore).setValue("Valore Ordine");
    Logger.log("[HW] Intestazione scritta in colonna %s", colValore);

    // ciclo sulle righe
    for (let i = startDataRow; i <= lastRow; i++) {
      try {
        const qta = sheet.getRange(i, colQta).getValue();
        const prezzo = sheet.getRange(i, colPrezzo).getValue();

        Logger.log("[HW] Riga %s | Qta=%s | Prezzo=%s", i, qta, prezzo);

        if (qta && prezzo) {
          const valore = qta * prezzo;
          sheet.getRange(i, colValore).setValue(valore);
          Logger.log("[HW] --> Valore scritto: %s", valore);
        } else {
          sheet.getRange(i, colValore).setValue("");
          Logger.log("[HW] --> Nessun valore (vuoto)");
        }
      } catch (rowErr) {
        Logger.log("[ERRORE HW - Riga %s] %s", i, rowErr.stack || rowErr);
      }
    }

    // ridimensiona colonna
    sheet.autoResizeColumn(colValore);
    Logger.log("[HW] Colonna %s ridimensionata automaticamente", colValore);

  } catch (err) {
    Logger.log("[ERRORE HW scriviValoreOrdine] %s", err.stack || err);
    throw err; // rilancia errore così interrompe
  }

  Logger.log("[HW] === FINE scriviValoreOrdineHW ===");
}

