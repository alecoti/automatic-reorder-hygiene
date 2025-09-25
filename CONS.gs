/**
 * Funzione principale per CONS
 */
function allocazioneScorteCONS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startRow = 1;        // Riga intestazioni principali
  const startDataRow = 2;    // Prima riga dati
  const startCol = 81;       // Colonna CO

  // Pulizia eventuali colonne oltre startCol
  const lastCol = sheet.getLastColumn();
  if (lastCol >= startCol) {
    sheet.deleteColumns(startCol, lastCol - startCol + 1);
  }

  // === STEP 1: Generazione disponibilità da donatori (solo CONS) ===
  const disponibilita = generaDisponibilitaCONS(data, startDataRow);

  // === STEP 2: Allocazione fabbisogni (solo CONS) ===
  const {
    risultati,
    trasferimenti,
    bwRicalcolato,
    maxAllocazioni
  } = allocaFabbisogniCONS(data, startDataRow, disponibilita);

  // === STEP 3: Scrittura risultati (solo CONS) ===
  scriviRisultatiCONS(sheet, startRow, startDataRow, startCol, risultati, bwRicalcolato, maxAllocazioni);

  // === STEP 4: Scrittura trasferimenti (solo CONS) ===
  scriviTrasferimentiCONS(ss, trasferimenti);

  // === STEP 5: Report acquisti (solo CONS) ===
  scriviReportAcquistiCONS();
}

/**
 * Crea una mappa delle disponibilità donatori (solo CONS).
 * BG < 0 = esubero, quindi disponibile a donare.
 */
function generaDisponibilitaCONS(data, startDataRow) {
  let disponibilita = {};
  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];           // Colonna D (N° Articolo)
    const ubic = data[i][1];             // Colonna B (Ubicazione)
    const stock = parseFloat(data[i][58]); // Colonna BG

    if (!codice || isNaN(stock)) continue;

    if (stock < 0) {
      if (!disponibilita[codice]) disponibilita[codice] = {};
      disponibilita[codice][ubic] = (disponibilita[codice][ubic] || 0) + Math.abs(stock);
    }
  }
  return disponibilita;
}

/**
 * Alloca i fabbisogni per CONS.
 */
function allocaFabbisogniCONS(data, startDataRow, disponibilita) {
  let trasferimenti = [];
  let bwRicalcolato = {};
  let risultati = [];
  let maxAllocazioni = 0;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];          // Colonna D
    const ubic = data[i][1];            // Colonna B
    const tipo = data[i][12];           // Colonna M
    const copertura = parseFloat(data[i][11]); // Colonna L
    let fab = parseFloat(data[i][56]);  // Colonna BE → fabbisogno
    let ord = parseFloat(data[i][61]);  // Colonna BJ → quantità da acquistare

    if (!codice || isNaN(fab) || isNaN(ord)) {
      risultati.push({});
      continue;
    }

    const key = codice + "|" + ubic;
    let allocazioni = [];
    let azione = "";
    let motivo = "";
    let acquisto = "";

    if (bwRicalcolato[key] === undefined) bwRicalcolato[key] = fab;

    if (tipo === "CONS" && fab > 0) {
      let fabBisogno = fab;

      // Caso 1: copertura nazionale <100 → acquisto diretto
      if (copertura < 100) {
        azione = "ACQUISTARE";
        motivo = "Copertura <100";
        acquisto = ord; // uso BJ per acquisti
        bwRicalcolato[key] = 0;

      // Caso 2: ci sono donatori disponibili
      } else if (disponibilita[codice]) {
        let candidates = Object.entries(disponibilita[codice])
          .filter(([mag]) => mag !== ubic)
          .sort((a, b) => b[1] - a[1]);

        for (let [donatore, qty] of candidates) {
          if (fabBisogno <= 0) break;
          if (qty <= 0) continue;

          const prelievo = Math.min(Math.ceil(fabBisogno), qty); // arrotondo fabbisogno per eccesso
          if (prelievo <= 0) continue;

          allocazioni.push([donatore, prelievo]);
          disponibilita[codice][donatore] -= prelievo;

          const donorKey = codice + "|" + donatore;
          bwRicalcolato[donorKey] = (bwRicalcolato[donorKey] || 0) + prelievo;
          bwRicalcolato[key] = (bwRicalcolato[key] || fab) - prelievo;

          fabBisogno -= prelievo;
          trasferimenti.push([codice, donatore, ubic, prelievo]);
        }

        if (allocazioni.length > 0) {
          azione = "TRASFERIMENTO";
        }
        if (fabBisogno > 0) {
          azione = "ACQUISTARE";
          motivo = "Residuo non coperto";
          acquisto = Math.ceil(fabBisogno);
          bwRicalcolato[key] = 0;
        }

      // Caso 3: nessun donatore
      } else {
        azione = "ACQUISTARE";
        motivo = "Nessun donatore disponibile";
        acquisto = ord;
        bwRicalcolato[key] = 0;
      }
    }

    maxAllocazioni = Math.max(maxAllocazioni, allocazioni.length);

    risultati.push({
      ubic,
      codice,
      key,
      bwRicalcolato: bwRicalcolato[key],
      allocazioni,
      azione,
      motivo,
      acquisto
    });
  }

  return { risultati, trasferimenti, bwRicalcolato, maxAllocazioni };
}

/**
 * Scrive sul foglio principale (solo CONS).
 */
function scriviRisultatiCONS(sheet, startRow, startDataRow, startCol, risultati, bwRicalcolato, maxAllocazioni) {
  const headers = [
    "Totale Donato", "Totale Ricevuto", "BE Ricalcolato",
    "Azione", "Motivo", "Quantità Acquisto"
  ];
  let allocHeaders = [];
  for (let j = 0; j < maxAllocazioni; j++) {
    allocHeaders.push("Source " + (j + 1));
    allocHeaders.push("Qty " + (j + 1));
  }

  const neededCols = headers.length + allocHeaders.length;
  const lastCol = sheet.getLastColumn();
  const missing = (startCol + neededCols - 1) - lastCol;
  if (missing > 0) {
    sheet.insertColumnsAfter(lastCol, missing);
  }

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
      r.acquisto || ""
    ];

    (r.allocazioni || []).forEach(a => rowData.push(a[0], a[1]));
    while (rowData.length < neededCols) rowData.push("");

    out.push(rowData);
  }

  const range = sheet.getRange(startDataRow, startCol, out.length, neededCols);
  range.setValues(out);

  // 🔥 Formattazione: testo centrato + colonne auto-resize
  range.setHorizontalAlignment("center");
  for (let col = startCol; col < startCol + neededCols; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Scrive i trasferimenti (solo CONS).
 */
function scriviTrasferimentiCONS(ss, trasferimenti) {
  let tSheet = ss.getSheetByName("TrasferimentiCONS");
  if (!tSheet) {
    tSheet = ss.insertSheet("TrasferimentiCONS");
  } else {
    tSheet.clear();
  }

  tSheet.getRange(1, 1, 1, 4).setValues([["Articolo", "Da", "A", "Quantità"]]);

  if (trasferimenti.length > 0) {
    const range = tSheet.getRange(2, 1, trasferimenti.length, 4);
    range.setValues(trasferimenti);
    range.setHorizontalAlignment("center");
    for (let col = 1; col <= 4; col++) {
      tSheet.autoResizeColumn(col);
    }

    // Ordinamento per Da -> A -> Articolo
    const lastRow = tSheet.getLastRow();
    if (lastRow > 1) {
      tSheet.getRange(2, 1, lastRow - 1, 4).sort([
        { column: 2, ascending: true }, // Da
        { column: 3, ascending: true }, // A
        { column: 1, ascending: true }  // Articolo
      ]);
    }

    // Colori alternati per combinazioni (Da + A)
    const values = tSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let colors = [];
    let currentCombo = null;
    let useGreen = true;

    for (let i = 0; i < values.length; i++) {
      const combo = values[i][1] + "->" + values[i][2]; // Da + A
      if (combo !== currentCombo) {
        currentCombo = combo;
        useGreen = !useGreen;
      }
      let color = useGreen ? "#e6f4ea" : "#e6f0f9";
      colors.push([color, color, color, color]);
    }

    tSheet.getRange(2, 1, lastRow - 1, 4).setBackgrounds(colors);

    // === REPORT ===
    let totals = {};       // Totali donato/ricevuto
    let combinazioni = {}; // Totale pezzi e codici per Da->A

    values.forEach(([cod, da, a, qty]) => {
      // Totali donato/ricevuto
      if (!totals[da]) totals[da] = { donato: 0, ricevuto: 0 };
      if (!totals[a]) totals[a] = { donato: 0, ricevuto: 0 };
      totals[da].donato += qty;
      totals[a].ricevuto += qty;

      // Combinazioni Da->A
      const combo = da + "->" + a;
      if (!combinazioni[combo]) {
        combinazioni[combo] = { pezzi: 0, articoli: new Set() };
      }
      combinazioni[combo].pezzi += qty;
      combinazioni[combo].articoli.add(cod);
    });

    // Scrivi report su nuovo foglio
    let rSheet = ss.getSheetByName("ReportTrasferimentiCONS");
    if (!rSheet) {
      rSheet = ss.insertSheet("ReportTrasferimentiCONS");
    } else {
      rSheet.clear();
    }

    // Prima tabella: Totali donato/ricevuto
    let summary = [["Magazzino", "Totale Donato", "Totale Ricevuto"]];
    for (let mag in totals) {
      summary.push([mag, totals[mag].donato, totals[mag].ricevuto]);
    }
    rSheet.getRange(1, 1, summary.length, 3).setValues(summary);
    rSheet.getRange(1, 1, summary.length, 3).setHorizontalAlignment("center");
    for (let col = 1; col <= 3; col++) {
      rSheet.autoResizeColumn(col);
    }

    // Seconda tabella: Report combinazioni
    let comboReport = [["Da", "A", "Totale Pezzi", "Codici Unici"]];
    for (let combo in combinazioni) {
      const [da, a] = combo.split("->");
      comboReport.push([
        da,
        a,
        combinazioni[combo].pezzi,
        combinazioni[combo].articoli.size
      ]);
    }

    const startCombo = summary.length + 2;
    rSheet.getRange(startCombo, 1, comboReport.length, 4).setValues(comboReport);
    rSheet.getRange(startCombo, 1, comboReport.length, 4).setHorizontalAlignment("center");
    for (let col = 1; col <= 4; col++) {
      rSheet.autoResizeColumn(col);
    }
  }
}

/**
 * Report degli acquisti CONS
 * Genera un riepilogo per location con quantità, prezzo e totale
 */
function scriviReportAcquistiCONS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startDataRow = 2; 

  // Nuovo foglio
  let rSheet = ss.getSheetByName("ReportAcquistiCONS");
  if (!rSheet) {
    rSheet = ss.insertSheet("ReportAcquistiCONS");
  } else {
    rSheet.clear();
  }

  let acquistiPerLoc = {};
  let totaleGenerale = 0;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const codice = data[i][3];    // Articolo
    const ubic = data[i][1];      // Ubicazione
    const qta = parseFloat(data[i][61]); // Colonna BJ
    const prezzo = parseFloat(data[i][60]); // Colonna BI

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
    output.push(["", "", "", ""]); // Riga vuota di separazione
  }

  // Totale finale
  output.push(["", "", "Totale Generale", totaleGenerale]);

  // Scrivi sul foglio
  rSheet.getRange(1, 1, output.length, 4).setValues(output);
  rSheet.getRange(1, 1, output.length, 4).setHorizontalAlignment("center");

  // Auto-resize colonne
  for (let col = 1; col <= 4; col++) {
    rSheet.autoResizeColumn(col);
  }
}
