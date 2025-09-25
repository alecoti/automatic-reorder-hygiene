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
  const startDataRow = 2;

  // === HW ===
  const disponibilitaHW = generaDisponibilitaHW(data, startDataRow);
  const risultatiHW = allocaFabbisogniHW(data, startDataRow, disponibilitaHW);

  // === CONS ===
  const disponibilitaCONS = generaDisponibilitaCONS(data, startDataRow);
  const risultatiCONS = allocaFabbisogniCONS(data, startDataRow, disponibilitaCONS);

  // === Output complessivo ===
  const panoramica = preparaOutputComplessivo(data, startDataRow, risultatiHW, risultatiCONS);
  scriviAllocazioneComplessiva(ss, panoramica);

  // === Trasferimenti & report ===
  const trasferimentiComplessivi = combinaTrasferimenti(risultatiHW.trasferimenti, risultatiCONS.trasferimenti);
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
 * Prepara le righe di output per il foglio consolidato.
 * @returns {{ headers: string[], righe: any[][] }}
 */
function preparaOutputComplessivo(data, startDataRow, risultatiHW, risultatiCONS) {
  const headersBase = [
    "Categoria",
    "Ubicazione",
    "Articolo",
    "Totale Donato",
    "Totale Ricevuto",
    "Fabbisogno Ricalcolato",
    "Azione",
    "Motivo",
    "Quantità Acquisto",
    "Valore Ordine"
  ];

  let righe = [];
  let maxAllocazioni = 0;

  for (let i = startDataRow - 1; i < data.length; i++) {
    const categoria = data[i][12]; // Colonna M
    if (categoria !== "HW" && categoria !== "CONS") continue;

    const ubicazione = data[i][1];
    const articolo = data[i][3];
    const prezzo = parseFloat(data[i][60]); // Colonna BI

    const risultato = (categoria === "HW")
      ? (risultatiHW.risultati[i] || {})
      : (risultatiCONS.risultati[i] || {});

    const allocazioniRaw = risultato.allocazioni || [];
    const allocazioni = allocazioniRaw.map(entry => {
      if (categoria === "HW") {
        return [entry[0], entry[1], entry[2] || ""];
      }
      // CONS: aggiungo il riferimento categoria per coerenza con HW
      return [entry[0], entry[1], "CONS"];
    });

    maxAllocazioni = Math.max(maxAllocazioni, allocazioni.length);

    const totaleDonato = calcolaTotaleDonato(allocazioni);
    const totaleRicevuto = calcolaTotaleRicevuto(allocazioni);
    const acquistoNum = parseFloat(risultato.acquisto);
    const acquisto = (!isNaN(acquistoNum) && acquistoNum > 0) ? acquistoNum : "";
    const haAllocazioni = allocazioni.length > 0;
    const fabbisognoRicalcolato = (haAllocazioni || acquisto)
      ? (risultato.bwRicalcolato ?? "")
      : "";
    const valoreOrdine = (acquisto && !isNaN(prezzo)) ? acquistoNum * prezzo : "";

    let row = [
      categoria || "",
      ubicazione || "",
      articolo || "",
      totaleDonato || "",
      totaleRicevuto || "",
      fabbisognoRicalcolato,
      risultato.azione || "",
      risultato.motivo || "",
      acquisto,
      valoreOrdine || ""
    ];

    allocazioni.forEach(a => {
      row.push(a[0] || "", a[1] || "", a[2] || "");
    });

    righe.push(row);
  }

  // Completa le colonne delle allocazioni
  const allocHeaders = [];
  for (let j = 0; j < maxAllocazioni; j++) {
    allocHeaders.push(`Source ${j + 1}`, `Qty ${j + 1}`, `Tipo ${j + 1}`);
  }

  const headers = headersBase.concat(allocHeaders);
  const valori = righe.map(row => padRow(row, headers.length));

  return { headers, righe: valori };
}

/**
 * Scrive il foglio complessivo di allocazione.
 */
function scriviAllocazioneComplessiva(ss, panoramica) {
  const sheetName = "AllocazioneNonParziale";
  let outSheet = ss.getSheetByName(sheetName);
  if (!outSheet) {
    outSheet = ss.insertSheet(sheetName);
  } else {
    outSheet.clear();
  }

  const headers = panoramica.headers;
  outSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (panoramica.righe.length > 0) {
    outSheet.getRange(2, 1, panoramica.righe.length, headers.length)
      .setValues(panoramica.righe);
  }

  const lastRow = Math.max(1, panoramica.righe.length + 1);
  outSheet.getRange(1, 1, lastRow, headers.length).setHorizontalAlignment("center");
  for (let col = 1; col <= headers.length; col++) {
    outSheet.autoResizeColumn(col);
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
