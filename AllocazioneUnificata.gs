/**
 * Funzioni di allocazione con logica combinata HW + CONS.
 */

/**
 * Punto di ingresso utilizzato dal bottone "Allocazione Scorte Non Parziale".
 * Elabora sia gli articoli HW che CONS e genera report consolidati.
 */
function allocazioneScorteNonParziale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startRow = 1;
  const startDataRow = 2;
  const startCol = 81; // Colonna CC

  // Pulisco eventuali residui dell'esecuzione precedente oltre la colonna CC
  const lastCol = sheet.getLastColumn();
  if (lastCol >= startCol) {
    sheet.deleteColumns(startCol, lastCol - startCol + 1);
  }

  // === HW ===
  const disponibilitaHW = generaDisponibilitaHW(data, startDataRow);
  const esitoHW = allocaFabbisogniHW(data, startDataRow, disponibilitaHW);

  // === CONS ===
  const disponibilitaCONS = generaDisponibilitaCONS(data, startDataRow);
  const esitoCONS = allocaFabbisogniCONS(data, startDataRow, disponibilitaCONS);

  // === Scrittura su foglio principale a partire dalla colonna CC ===
  const panoramica = preparaRisultatiPerFoglio(data, startDataRow, esitoHW, esitoCONS);
  scriviRisultatiAllocazioneUnificata(
    sheet,
    startRow,
    startDataRow,
    startCol,
    data,
    panoramica.risultati,
    panoramica.maxAllocazioni
  );

  // === Trasferimenti & report ===
  const trasferimentiComplessivi = combinaTrasferimenti(esitoHW.trasferimenti, esitoCONS.trasferimenti);
  scriviTrasferimentiComplessivi(ss, trasferimentiComplessivi);
  scriviReportTrasferimentiComplessivi(ss, trasferimentiComplessivi);
  scriviReportAcquistiComplessivi(ss, data, startDataRow);
}

/**
 * Alias mantenuto per il menu personalizzato esistente.
 */
function allocazioneScorteUnificata() {
  allocazioneScorteNonParziale();
}

/**
 * Prepara i risultati combinati pronti per la scrittura sul foglio principale.
 */
function preparaRisultatiPerFoglio(data, startDataRow, risultatiHW, risultatiCONS) {
  const risultati = new Array(data.length);
  let maxAllocazioni = 0;

  const offset = startDataRow - 1;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const categoria = data[i][12];
    let base = {};

    if (categoria === "HW") {
      const sorgente = risultatiHW.risultati[i] || {};
      base = {
        ubic: sorgente.ubic,
        codice: sorgente.codice,
        bwRicalcolato: sorgente.bwRicalcolato,
        allocazioni: (sorgente.allocazioni || []).map(a => [a[0], a[1], a[2] || ""]),
        azione: sorgente.azione,
        motivo: sorgente.motivo,
        acquisto: sorgente.acquisto
      };
    } else if (categoria === "CONS") {
      const consIndex = i - offset;
      const sorgente = risultatiCONS.risultati[consIndex] || {};
      base = {
        ubic: sorgente.ubic,
        codice: sorgente.codice,
        bwRicalcolato: sorgente.bwRicalcolato,
        allocazioni: (sorgente.allocazioni || []).map(a => [a[0], a[1], ""]),
        azione: sorgente.azione,
        motivo: sorgente.motivo,
        acquisto: sorgente.acquisto
      };
    }

    risultati[i] = base;
    if (base.allocazioni && base.allocazioni.length > maxAllocazioni) {
      maxAllocazioni = base.allocazioni.length;
    }
  }

  return { risultati, maxAllocazioni };
}

function scriviRisultatiAllocazioneUnificata(sheet, startRow, startDataRow, startCol, data, risultati, maxAllocazioni) {
  const headers = [
    "Totale Donato", "Totale Ricevuto", "Fabbisogno Ricalcolato",
    "Azione", "Motivo", "Quantità Acquisto", "Valore Ordine"
  ];

  const allocHeaders = [];
  for (let j = 0; j < maxAllocazioni; j++) {
    allocHeaders.push(`Source ${j + 1}`);
    allocHeaders.push(`Qty ${j + 1}`);
    allocHeaders.push(`Tipo ${j + 1}`);
  }

  const neededCols = headers.length + allocHeaders.length;
  const lastCol = sheet.getLastColumn();
  const missing = (startCol + neededCols - 1) - lastCol;
  if (missing > 0) {
    sheet.insertColumnsAfter(lastCol, missing);
  }

  sheet.getRange(startRow, startCol, 1, neededCols)
    .setValues([headers.concat(allocHeaders)]);

  const out = [];
  for (let i = startDataRow - 1; i < data.length; i++) {
    const categoria = data[i][12];
    const risultato = risultati[i] || {};

    if (categoria !== "HW" && categoria !== "CONS") {
      out.push(new Array(neededCols).fill(""));
      continue;
    }

    const allocazioni = (risultato.allocazioni || []).map(a => [a[0] || "", a[1] || "", a[2] || ""]);
    const totaleDonato = calcolaTotaleDonato(allocazioni);
    const totaleRicevuto = calcolaTotaleRicevuto(allocazioni);

    const acquistoNum = parseFloat(risultato.acquisto);
    const acquisto = (!isNaN(acquistoNum) && acquistoNum > 0) ? acquistoNum : "";
    const haAllocazioni = allocazioni.length > 0;
    const fabbisognoRicalcolato = (haAllocazioni || acquisto !== "")
      ? (risultato.bwRicalcolato ?? "")
      : "";

    const rowData = [
      totaleDonato || "",
      totaleRicevuto || "",
      fabbisognoRicalcolato,
      risultato.azione || "",
      risultato.motivo || "",
      acquisto,
      ""
    ];

    allocazioni.forEach(a => rowData.push(a[0], a[1], a[2]));
    while (rowData.length < neededCols) rowData.push("");

    out.push(rowData);
  }

  if (out.length > 0) {
    const range = sheet.getRange(startDataRow, startCol, out.length, neededCols);
    range.setValues(out);
    range.setHorizontalAlignment("center");
  }

  for (let col = startCol; col < startCol + neededCols; col++) {
    sheet.autoResizeColumn(col);
  }

  const lastRowOut = startDataRow + out.length - 1;
  if (out.length > 0) {
    const colAzione = startCol + 3;
    const colQta = startCol + 5;
    const colValore = startCol + 6;
    const colPrezzo = 61; // Colonna BI

    for (let r = startDataRow; r <= lastRowOut; r++) {
      const formula = `=IF($${colToLetter(colAzione)}${r}="ACQUISTARE", $${colToLetter(colQta)}${r}*$${colToLetter(colPrezzo)}${r}, "")`;
      sheet.getRange(r, colValore).setFormula(formula);
    }
  }
}

/**
 * Combina i trasferimenti HW e CONS in un'unica tabella.
 */
function combinaTrasferimenti(hwTrasferimenti, consTrasferimenti) {
  const output = [];

  (hwTrasferimenti || []).forEach(row => {
    if (!row || row.length < 4) return;
    const [codiceRaw, da, a, qty] = row;
    if (!da || !a) return;

    let codice = codiceRaw || "";
    let origine = "";
    if (/_U$/.test(codice)) {
      origine = "USATI";
      codice = codice.replace(/_U$/, "");
    } else if (/_S$/.test(codice)) {
      origine = "STOCK";
      codice = codice.replace(/_S$/, "");
    }

    output.push([
      "HW",
      codice,
      origine,
      da,
      a,
      Number(qty) || 0
    ]);
  });

  (consTrasferimenti || []).forEach(row => {
    if (!row || row.length < 4) return;
    const [codice, da, a, qty] = row;
    if (!da || !a) return;
    output.push([
      "CONS",
      codice || "",
      "",
      da,
      a,
      Number(qty) || 0
    ]);
  });

  return output;
}

/**
 * Scrive il foglio "TrasferimentiComplessivi" con colori alternati.
 */
function scriviTrasferimentiComplessivi(ss, trasferimenti) {
  const sheetName = "TrasferimentiComplessivi";
  let tSheet = ss.getSheetByName(sheetName);
  if (!tSheet) {
    tSheet = ss.insertSheet(sheetName);
  } else {
    tSheet.clear();
  }

  const headers = ["Categoria", "Articolo", "Origine Donatore", "Da", "A", "Quantità"];
  tSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (!trasferimenti || trasferimenti.length === 0) {
    return;
  }

  // Ordina per categoria, Da, A, Articolo
  trasferimenti.sort((a, b) => {
    if (a[0] !== b[0]) return a[0] > b[0] ? 1 : -1;
    if (a[3] !== b[3]) return a[3] > b[3] ? 1 : -1;
    if (a[4] !== b[4]) return a[4] > b[4] ? 1 : -1;
    return a[1] > b[1] ? 1 : (a[1] < b[1] ? -1 : 0);
  });

  tSheet.getRange(2, 1, trasferimenti.length, headers.length)
    .setValues(trasferimenti)
    .setHorizontalAlignment("center");

  const lastRow = tSheet.getLastRow();
  if (lastRow <= 1) return;

  // Colori alternati per combinazioni Categoria+Da->A
  const values = tSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  let colors = [];
  let currentCombo = null;
  let useGreen = true;

  values.forEach(row => {
    const combo = row[0] + "|" + row[3] + "->" + row[4];
    if (combo !== currentCombo) {
      currentCombo = combo;
      useGreen = !useGreen;
    }
    const color = useGreen ? "#e6f4ea" : "#e6f0f9";
    colors.push(new Array(headers.length).fill(color));
  });

  tSheet.getRange(2, 1, lastRow - 1, headers.length).setBackgrounds(colors);

  for (let col = 1; col <= headers.length; col++) {
    tSheet.autoResizeColumn(col);
  }
}

/**
 * Genera il report dei trasferimenti complessivi (totali e combinazioni).
 */
function scriviReportTrasferimentiComplessivi(ss, trasferimenti) {
  const sheetName = "ReportTrasferimentiComplessivi";
  let rSheet = ss.getSheetByName(sheetName);
  if (!rSheet) {
    rSheet = ss.insertSheet(sheetName);
  } else {
    rSheet.clear();
  }

  if (!trasferimenti || trasferimenti.length === 0) {
    rSheet.getRange(1, 1, 1, 1).setValue("Nessun trasferimento disponibile");
    return;
  }

  const totals = {};
  const combinazioni = {};

  trasferimenti.forEach(([categoria, codice, , da, a, qty]) => {
    if (!totals[da]) totals[da] = { donato: 0, ricevuto: 0 };
    if (!totals[a]) totals[a] = { donato: 0, ricevuto: 0 };
    totals[da].donato += qty;
    totals[a].ricevuto += qty;

    const comboKey = da + "->" + a;
    if (!combinazioni[comboKey]) {
      combinazioni[comboKey] = {
        pezzi: 0,
        articoli: new Set(),
        categorie: new Set()
      };
    }
    combinazioni[comboKey].pezzi += qty;
    combinazioni[comboKey].articoli.add(codice);
    combinazioni[comboKey].categorie.add(categoria);
  });

  const summary = [["Magazzino", "Totale Donato", "Totale Ricevuto"]];
  Object.keys(totals).sort().forEach(mag => {
    summary.push([mag, totals[mag].donato, totals[mag].ricevuto]);
  });

  rSheet.getRange(1, 1, summary.length, 3).setValues(summary);
  rSheet.getRange(1, 1, summary.length, 3).setHorizontalAlignment("center");

  const comboRows = [["Da", "A", "Totale Pezzi", "Codici Unici", "Categorie"]];
  Object.keys(combinazioni).sort().forEach(key => {
    const [da, a] = key.split("->");
    const info = combinazioni[key];
    comboRows.push([
      da,
      a,
      info.pezzi,
      info.articoli.size,
      Array.from(info.categorie).sort().join(", ")
    ]);
  });

  const startCombo = summary.length + 2;
  rSheet.getRange(startCombo, 1, comboRows.length, comboRows[0].length)
    .setValues(comboRows)
    .setHorizontalAlignment("center");

  const totalRows = startCombo + comboRows.length - 1;
  for (let col = 1; col <= comboRows[0].length; col++) {
    rSheet.autoResizeColumn(col);
  }
  rSheet.autoResizeColumn(4);
  rSheet.autoResizeColumn(5);

  rSheet.getRange(1, 1, totalRows, comboRows[0].length).setHorizontalAlignment("center");
}

/**
 * Genera il report acquisti complessivo per HW + CONS.
 */
function scriviReportAcquistiComplessivi(ss, data, startDataRow) {
  const sheetName = "ReportAcquistiComplessivi";
  let rSheet = ss.getSheetByName(sheetName);
  if (!rSheet) {
    rSheet = ss.insertSheet(sheetName);
  } else {
    rSheet.clear();
  }

  const acquistiPerLoc = {};
  let totaleGenerale = 0;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const categoria = data[i][12];
    const codice = data[i][3];
    const ubic = data[i][1];
    const prezzo = parseFloat(data[i][60]);

    let quantita = null;
    if (categoria === "HW") {
      quantita = parseFloat(data[i][75]); // BX
    } else if (categoria === "CONS") {
      quantita = parseFloat(data[i][61]); // BJ
    } else {
      continue;
    }

    if (!codice || isNaN(quantita) || quantita <= 0 || isNaN(prezzo)) continue;

    const totale = quantita * prezzo;
    totaleGenerale += totale;

    if (!acquistiPerLoc[ubic]) acquistiPerLoc[ubic] = [];
    acquistiPerLoc[ubic].push([categoria, codice, quantita, prezzo, totale]);
  }

  const headers = ["Categoria", "Articolo", "Quantità", "Prezzo Unitario", "Totale Riga"];
  const output = [];

  Object.keys(acquistiPerLoc).sort().forEach(loc => {
    output.push([`📦 Location: ${loc}`, "", "", "", ""]);
    output.push(headers);

    let totaleLoc = 0;
    acquistiPerLoc[loc].forEach(riga => {
      output.push(riga);
      totaleLoc += riga[4];
    });

    output.push(["", "", "", "Totale Location", totaleLoc]);
    output.push(["", "", "", "", ""]);
  });

  output.push(["", "", "", "Totale Generale", totaleGenerale]);

  rSheet.getRange(1, 1, output.length, headers.length)
    .setValues(output)
    .setHorizontalAlignment("center");

  for (let col = 1; col <= headers.length; col++) {
    rSheet.autoResizeColumn(col);
  }
}
