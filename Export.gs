/**
 * System Zarządzania Szkoleniami - Silnik Eksportu ESG (Export.gs)
 */

function wygenerujRaportExcel(wybranyRok) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_DB_ID);
    const trainingRows = ss.getSheetByName('Szkolenia').getDataRange().getValues();
    const participantRows = ss.getSheetByName('Uczestnicy_Szkolen').getDataRange().getValues();
    
    trainingRows.shift(); 
    participantRows.shift();
    
    const aktywneSzkolenia = trainingRows.filter(r => r[0] !== "");
    const aktywniUczestnicy = participantRows.filter(p => p[0] !== "");
    
    const slownikTlumaczen = zaladujSlownikZTabeli(ss);
    
    const naglowki = [
      "Nr.",
      "SURNAME NAME",
      "EMPLOYEE IDENTYFICATION NUMBER",
      "DEPARTMENT",
      "COMPANY ID",
      "TYPE",
      "COURSE TITLE",
      "DURATION [HOURS]",
      "COURSE START DATA",
      "COURSE END DATA",
      "TOPIC",
      "MODALITY",
      "INSTRUCTOR",
      "PLANNING",
      "COMPLETATION CERTYFICATE NUMBER"
    ];
    
    const gotoweWierszeRaportu = [];
    
    aktywniUczestnicy.forEach(p => {
      const trainingId = getSafeTrainingId(p[1]);
      const szkolenie = aktywneSzkolenia.find(t => getSafeTrainingId(t[0]) === trainingId);
      
      if (szkolenie) {
        const indywidualneGodziny = (String(szkolenie[13]).toLowerCase() === 'true');
        const indywidualneDaty = (String(szkolenie[14]).toLowerCase() === 'true');

        let mainStart = szkolenie[7];
        if (mainStart instanceof Date) mainStart = Utilities.formatDate(mainStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (mainStart) mainStart = mainStart.toString();
        
        let mainEnd = szkolenie[8];
        if (mainEnd instanceof Date) mainEnd = Utilities.formatDate(mainEnd, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (mainEnd) mainEnd = mainEnd.toString();

        let partStart = p[9];
        if (partStart instanceof Date) partStart = Utilities.formatDate(partStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (partStart) partStart = partStart.toString();

        let partEnd = p[10];
        if (partEnd instanceof Date) partEnd = Utilities.formatDate(partEnd, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (partEnd) partEnd = partEnd.toString();

        const ostatecznaDataStart = indywidualneDaty && partStart ? partStart : mainStart;
        const ostatecznaDataKoniec = indywidualneDaty && partEnd ? partEnd : mainEnd;
        const ostatecznaDlugosc = indywidualneGodziny ? p[8] : szkolenie[9];

        // --- FILTROWANIE PO ROKU ---
        const startStr = ostatecznaDataStart ? ostatecznaDataStart.toString() : "";
        const koniecStr = ostatecznaDataKoniec ? ostatecznaDataKoniec.toString() : "";
        const idStr = szkolenie[0] ? szkolenie[0].toString() : "";

        if (startStr.startsWith(wybranyRok) || koniecStr.startsWith(wybranyRok) || idStr.includes("TR/" + wybranyRok + "/")) {

          const surowyDzial = p[6];
          const surowaFormula = szkolenie[4];
          const surowyTypSzkolenia = szkolenie[5];
          const surowaKategoria = szkolenie[6];
          const surowyInstruktor = szkolenie[2];
          const surowePlanowanie = szkolenie[3];

          gotoweWierszeRaportu.push([
            trainingId, // <--- Zamiast licznika, wrzucamy rzeczywiste ID (np. TR/2026/1)                                               
            "",                                                             
            "",                                                             
            tlow(slownikTlumaczen.dzial, surowyDzial),                       
            formatujIdPracownika(p[2]),                                     
            tlow(slownikTlumaczen.formula, surowaFormula),                 
            szkolenie[1],                                                   
            ostatecznaDlugosc,                                              
            ostatecznaDataStart,                                            
            ostatecznaDataKoniec,                                           
            tlow(slownikTlumaczen.kategoria, surowaKategoria),             
            tlow(slownikTlumaczen.typSzkolenia, surowyTypSzkolenia),         
            tlow(slownikTlumaczen.instruktor, surowyInstruktor),             
            tlow(slownikTlumaczen.planowanie, surowePlanowanie),             
            szkolenie[10]                                                   
          ]);
        }
      }
    });
    
    if (gotoweWierszeRaportu.length === 0) {
      return { status: "empty", message: `Brak szkoleń w bazie dla roku ${wybranyRok}.` };
    }
    
    const tempSS = SpreadsheetApp.create('Tymczasowy_Raport_ESG');
    const tempSheet = tempSS.getSheets()[0];
    
    tempSheet.appendRow(naglowki);
    
    tempSheet.getRange(2, 5, gotoweWierszeRaportu.length, 1).setNumberFormat('@');
    // Wymuszamy formatowanie tekstowe dla kolumny "Nr." żeby TR/... dobrze wchodziło
    tempSheet.getRange(2, 1, gotoweWierszeRaportu.length, 1).setNumberFormat('@'); 
    
    tempSheet.getRange(2, 1, gotoweWierszeRaportu.length, 15).setValues(gotoweWierszeRaportu);
    
    const headerRange = tempSheet.getRange(1, 1, 1, 15);
    headerRange.setFontWeight("bold").setBackground("#212529").setFontColor("#ffffff");
    
    SpreadsheetApp.flush();
    
    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    
    const xlsxBlob = response.getBlob();
    const base64Data = Utilities.base64Encode(xlsxBlob.getBytes());
    
    const tempFile = DriveApp.getFileById(tempSS.getId());
    tempFile.setTrashed(true);
    
    const dataWydania = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    return {
      status: "success",
      base64: base64Data,
      fileName: `ESG_Sustainability_Report_${wybranyRok}_${dataWydania}.xlsx`
    };
    
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function zaladujSlownikZTabeli(ss) {
  const sheet = ss.getSheetByName('Tlumacz');
  let mapy = { dzial: {}, formula: {}, kategoria: {}, typSzkolenia: {}, instruktor: {}, planowanie: {} };
  
  if (!sheet) return mapy; 
  
  const dane = sheet.getDataRange().getValues();
  
  for (let i = 1; i < dane.length; i++) {
    const r = dane[i];
    if (r[0])  mapy.dzial[r[0].toString().trim().toLowerCase()] = r[1].toString().trim();
    if (r[2])  mapy.formula[r[2].toString().trim().toLowerCase()] = r[3].toString().trim();
    if (r[4])  mapy.kategoria[r[4].toString().trim().toLowerCase()] = r[5].toString().trim();
    if (r[6])  mapy.typSzkolenia[r[6].toString().trim().toLowerCase()] = r[7].toString().trim();
    if (r[8])  mapy.instruktor[r[8].toString().trim().toLowerCase()] = r[9].toString().trim();
    if (r[10]) mapy.planowanie[r[10].toString().trim().toLowerCase()] = r[11].toString().trim();
  }
  return mapy;
}

function tlow(subSlownik, wartoscPL) {
  if (!wartoscPL) return "";
  const klucz = wartoscPL.toString().trim().toLowerCase();
  return subSlownik[klucz] || wartoscPL.toString().trim();
}
