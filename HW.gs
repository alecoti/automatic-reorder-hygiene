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

  // === STEP 1: Generazione disponibilità da donatori (HW) ===
  const disponibilita = generaDisponibilitaHW(data, startDataRow);

  // === STEP 2: Allocazione fabbisogni (HW) ===
  const {
    risultati,
    trasferimenti,
    bwRicalcolato,
    maxAllocazioni
  } = allocaFabbisogniHW(data, startDataRow, disponibilita);

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

/**
 * Alloca i fabbisogni per HW con logica ottimizzata:
 * - Considera solo le righe con categoria = "HW"
 * - Raggruppa per articolo
 * - Serve prima i riceventi con fabbisogno maggiore
 * - Usa i donor con disponibilità maggiore (preferendo Usati "_U" a Stock in caso di parità)
 */

function allocaFabbisogniHW(data, startDataRow, disponibilita) {
  let trasferimenti = [];
  let bwRicalcolato = {};
  let risultati = [];
  let maxAllocazioni = 0;

  const { disponibilitaU, disponibilitaS } = disponibilita;

  Logger.log("=== INIZIO ALLOCAZIONE HW ===");
  Logger.log("Disponibilità Usati (U): %s", JSON.stringify(disponibilitaU));
  Logger.log("Disponibilità Stock (S): %s", JSON.stringify(disponibilitaS));

  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];               // Articolo
    const ubic = data[i][1];                 // Ubicazione
    const tipo = data[i][12];                // Tipo (es. HW)
    const copertura = parseFloat(data[i][11]);
    let fabU = parseFloat(data[i][69]);      // Colonna BR
    let fabS = parseFloat(data[i][70]);      // Colonna BS
    let ord = parseFloat(data[i][75]);       // Colonna BX

    Logger.log("Riga %s -> codice:%s, ubic:%s, tipo:%s, copertura:%s, fabU:%s, fabS:%s, ord:%s",
      i + 1, codice, ubic, tipo, copertura, fabU, fabS, ord);

    // Skip se non valida
    if (!codice || (isNaN(fabU) && isNaN(fabS)) || isNaN(ord)) {
      Logger.log("Riga %s SKIPPATA (dati mancanti o non validi)", i + 1);
      risultati.push({});
      continue;
    }

    const key = codice + "|" + ubic;
    let allocazioni = [];
    let azione = "";
    let motivo = "";
    let acquisto = "";

    const fabTotale = (isNaN(fabU) ? 0 : fabU) + (isNaN(fabS) ? 0 : fabS);
    if (bwRicalcolato[key] === undefined) bwRicalcolato[key] = fabTotale;

    Logger.log("Riga %s -> fabTotale: %s", i + 1, fabTotale);

    if (tipo === "HW" && fabTotale > 0) {
      let fabBisogno = fabTotale;
      Logger.log("Riga %s -> fabBisogno iniziale: %s", i + 1, fabBisogno);

      if (copertura < 100) {
        azione = "ACQUISTARE";
        motivo = "Copertura <100";
        acquisto = ord;
        bwRicalcolato[key] = 0;
        Logger.log("Riga %s -> DECISIONE: ACQUISTARE (copertura <100)", i + 1);

      } else {
        Logger.log("Riga %s -> INIZIO allocazione da donor", i + 1);

        // Lista candidati (sia usati che stock), ordinata prima per qty, poi preferendo U
        let candidati = [];

        if (disponibilitaU[codice]) {
          for (let [mag, qty] of Object.entries(disponibilitaU[codice])) {
            if (mag !== ubic && qty > 0) {
              candidati.push({ mag, qty, tipo: "_U" });
            }
          }
        }
        if (disponibilitaS[codice]) {
          for (let [mag, qty] of Object.entries(disponibilitaS[codice])) {
            if (mag !== ubic && qty > 0) {
              candidati.push({ mag, qty, tipo: "" });
            }
          }
        }

        // Ordina: prima quantità più grande, a parità preferisci Usato
        candidati.sort((a, b) => {
          if (b.qty !== a.qty) return b.qty - a.qty;
          if (a.tipo === "_U" && b.tipo !== "_U") return -1;
          if (a.tipo !== "_U" && b.tipo === "_U") return 1;
          return 0;
        });

        Logger.log("Riga %s -> Candidati donor ordinati: %s", i + 1, JSON.stringify(candidati));

        // Assegna fino a coprire il fabBisogno
        for (let donor of candidati) {
          if (fabBisogno <= 0) break;
          const prelievo = Math.min(fabBisogno, donor.qty);
          if (prelievo <= 0) continue;

          Logger.log("Riga %s -> Assegno da donor %s (%s) qty:%s", i + 1, donor.mag, donor.tipo, prelievo);

          allocazioni.push([donor.mag, prelievo, donor.tipo]);

          if (donor.tipo === "_U") {
            disponibilitaU[codice][donor.mag] -= prelievo;
            trasferimenti.push([codice + "_U", donor.mag, ubic, prelievo]);
          } else {
            disponibilitaS[codice][donor.mag] -= prelievo;
            trasferimenti.push([codice, donor.mag, ubic, prelievo]);
          }

          const donorKey = codice + "|" + donor.mag;
          bwRicalcolato[donorKey] = (bwRicalcolato[donorKey] || 0) + prelievo;
          bwRicalcolato[key] -= prelievo;

          fabBisogno -= prelievo;
        }

        if (fabBisogno > 0) {
          azione = "ACQUISTARE";
          motivo = "Residuo non coperto";
          acquisto = fabBisogno;
          bwRicalcolato[key] = 0;
          Logger.log("Riga %s -> DECISIONE: ACQUISTARE residuo %s", i + 1, fabBisogno);
        } else {
          azione = "TRASFERIMENTO";
          Logger.log("Riga %s -> DECISIONE: TRASFERIMENTO completato", i + 1);
        }
      }
    }

    maxAllocazioni = Math.max(maxAllocazioni, allocazioni.length);

    const resultRow = {
      ubic,
      codice,
      key,
      bwRicalcolato: bwRicalcolato[key],
      allocazioni,
      azione,
      motivo,
      acquisto
    };

    risultati.push(resultRow);
    Logger.log("Riga %s -> RISULTATO: %s", i + 1, JSON.stringify(resultRow));
  }

  Logger.log("=== FINE ALLOCAZIONE HW ===");
  return { risultati, trasferimenti, bwRicalcolato, maxAllocazioni };
}



/**
 * Scrive sul foglio principale (solo HW).
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
  const lastCol = sheet.getLastColumn();
  const missing = (startCol + neededCols - 1) - lastCol;
  if (missing > 0) {
    sheet.insertColumnsAfter(lastCol, missing);
  }

  // Scrivo intestazioni
  sheet.getRange(startRow, startCol, 1, neededCols)
    .setValues([headers.concat(allocHeaders)]);

  let out = [];
  for (let i = 0; i < risultati.length; i++) {
    const r = risultati[i];
    if (!r.ubic) {
      out.push(new Array(neededCols).fill(""));
      continue;
    }

    const totDonato = (r.allocazioni || []).reduce((sum, a) => sum + a[1], 0);
    const totRicevuto = (r.allocazioni || []).reduce((sum, a) => sum + a[1], 0);
    const bwVal = (r.allocazioni.length > 0 || r.acquisto) ? r.bwRicalcolato : "";

    let rowData = [
      totDonato || "",
      totRicevuto || "",
      bwVal,
      r.azione || "",
      r.motivo || "",
      r.acquisto || "",   // colonna Quantità Acquisto
      ""                  // colonna Valore Ordine → ci scrivo formula dopo
    ];

    (r.allocazioni || []).forEach(a => rowData.push(a[0], a[1], a[2]));
    while (rowData.length < neededCols) rowData.push("");

    out.push(rowData);
  }

  const range = sheet.getRange(startDataRow, startCol, out.length, neededCols);
  range.setValues(out);

  range.setHorizontalAlignment("center");
  for (let col = startCol; col < startCol + neededCols; col++) {
    sheet.autoResizeColumn(col);
  }

  // === Inserisco la formula in "Valore Ordine" ===
  const lastRow = sheet.getLastRow();
  const colAzione = startCol + 3;         // colonna Azione
  const colQta = startCol + 5;            // Quantità Acquisto
  const colValore = startCol + 6;         // Valore Ordine
  const colPrezzo = 61;                   // Colonna BI (prezzo unitario)

  for (let r = startDataRow; r <= lastRow; r++) {
    const formula = `=IF($${colToLetter(colAzione)}${r}="ACQUISTARE", $${colToLetter(colQta)}${r}*$BI${r}, "")`;
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
  const lastRow = sheet.getLastRow();
  const colQta = 76; // BX
  const colPrezzo = 61; // BI
  const colValore = colQta + 1; // BY

  sheet.getRange(1, colValore).setValue("Valore Ordine");

  for (let i = startDataRow; i <= lastRow; i++) {
    const qta = sheet.getRange(i, colQta).getValue();
    const prezzo = sheet.getRange(i, colPrezzo).getValue();
    if (qta && prezzo) {
      sheet.getRange(i, colValore).setValue(qta * prezzo);
    } else {
      sheet.getRange(i, colValore).setValue("");
    }
  }

  sheet.autoResizeColumn(colValore);
}
