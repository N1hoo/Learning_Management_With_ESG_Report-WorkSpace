/**
 * System Zarządzania Szkoleniami - Backend (Kod.gs)
 */

const ADMIN_EMAIL = 'Twój@mail.com'; 
const CSV_FILE_ID = 'CSV_ID'; 
const MAIN_DB_ID = 'DB_ID';

function doGet() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  let hasAccess = false;
  
  if (userEmail === ADMIN_EMAIL.toLowerCase()) {
    hasAccess = true;
  } else {
    try {
      const ss = SpreadsheetApp.openById(MAIN_DB_ID);
      const sheet = ss.getSheetByName('Uprawnieni');
      if (sheet) {
        const data = sheet.getRange("A:A").getValues().flat();
        const allowedEmails = data.map(e => e.toString().trim().toLowerCase()).filter(e => e !== "");
        if (allowedEmails.includes(userEmail)) {
          hasAccess = true;
        }
      }
    } catch (e) {
      Logger.log("Błąd weryfikacji uprawnień: " + e.message);
    }
  }
  
  if (hasAccess) {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('System Zarządzania Szkoleniami')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; text-align: center; padding-top: 100px; color: #333;">
        <h1 style="color: #dc3545; font-size: 48px;">Odmowa dostępu</h1>
        <p style="font-size: 18px;">Twój adres e-mail (<b>${userEmail}</b>) nie znajduje się w bazie autoryzowanych użytkowników tego systemu.</p>
        <p style="font-size: 16px; color: #666;">Jeśli uważasz, że to błąd, skontaktuj się z administratorem systemu LMS.</p>
      </div>
    `).setTitle('Brak uprawnień');
  }
}

function formatujIdPracownika(id) {
  if (!id) return "";
  let idStr = id.toString().trim();
  if (/^\d+$/.test(idStr)) {
    return idStr.padStart(6, '0');
  }
  return idStr;
}

function getSafeTrainingId(rawId) {
  if (!rawId) return "";
  
  if (rawId instanceof Date) {
    const gYear = rawId.getFullYear();
    const gMonth = rawId.getMonth() + 1; 
    return `TR/${gYear}/${gMonth}`;
  }
  
  let idStr = rawId.toString().trim();
  
  if (idStr.match(/^\d{4}\/\d+$/)) {
    return `TR/${idStr}`;
  }
  
  return idStr;
}

function getDictionaries() {
  const ss = SpreadsheetApp.openById(MAIN_DB_ID);
  const sheet = ss.getSheetByName('Slowniki');
  const data = sheet.getDataRange().getValues();
  
  let dicts = {
    instruktor: [],
    planowanie: [],
    formula: [],
    typ: [],
    kategoria: []
  };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) dicts.instruktor.push(data[i][0]);
    if (data[i][1]) dicts.planowanie.push(data[i][1]);
    if (data[i][2]) dicts.formula.push(data[i][2]);
    if (data[i][3]) dicts.typ.push(data[i][3]);
    if (data[i][4]) dicts.kategoria.push(data[i][4]);
  }
  return dicts;
}

function getEmployees() {
  const ss = SpreadsheetApp.openById(MAIN_DB_ID);
  const sheet = ss.getSheetByName('Pracownicy');
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  const filteredData = data.filter(r => r[0] !== "");
  
  return filteredData.map(r => ({
    id: formatujIdPracownika(r[0]), 
    nazwisko: r[1],
    imie: r[2],
    stanowisko: r[3],
    dzial: r[4],
    zmiana: r[5]
  }));
}

/**
 * Nowy schemat generowania ID sterowany wprost z panelu filtracji
 */
function generateTrainingId(wybranyRok) {
  let year = wybranyRok ? wybranyRok.toString().trim() : new Date().getFullYear().toString(); 
  
  const ss = SpreadsheetApp.openById(MAIN_DB_ID);
  const sheet = ss.getSheetByName('Szkolenia');
  const values = sheet.getDataRange().getValues();
  values.shift(); 
  
  let maxNum = 0;
  const prefix = `TR/${year}/`;
  
  values.forEach(r => {
    const idStr = getSafeTrainingId(r[0]);
    if (idStr.startsWith(prefix)) {
      const parts = idStr.split('/');
      if (parts.length === 3) { 
        const num = parseInt(parts[2]);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  });
  
  return `${prefix}${maxNum + 1}`;
}

function saveTraining(formData, participants, wybranyRok) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_DB_ID);
    const trainingSheet = ss.getSheetByName('Szkolenia');
    const participantSheet = ss.getSheetByName('Uczestnicy_Szkolen');
    const userEmail = Session.getActiveUser().getEmail();
    
    // Używamy roku pobranego z widoku HTML
    const trainingId = generateTrainingId(wybranyRok);
    const nextTrainingRow = trainingSheet.getLastRow() + 1;
    trainingSheet.getRange(nextTrainingRow, 1).setNumberFormat('@');
    
    trainingSheet.appendRow([
      trainingId,
      formData.tytul,
      formData.typInstruktora,
      formData.planowanie,
      formData.formula,
      formData.typSzkolenia,
      formData.kategoria,
      formData.dataStart || "",
      formData.dataKoniec || "",
      formData.dlugosc || "",
      formData.certyfikat,
      userEmail,
      "",
      formData.indywidualneGodziny, 
      formData.indywidualneDaty     
    ]);
    
    if (participants && participants.length > 0) {
      const isGodziny = formData.indywidualneGodziny;
      const isDaty = formData.indywidualneDaty;
      
      const bazaGodzin = formData.dlugosc || "";
      const bazaStart = formData.dataStart || "";
      const bazaKoniec = formData.dataKoniec || "";

      const rowsToAppend = participants.map(p => [
        Utilities.getUuid(), 
        trainingId, 
        formatujIdPracownika(p.id), 
        p.nazwisko,
        p.imie,
        p.stanowisko,
        p.dzial,
        p.zmiana,
        isGodziny ? (p.godziny || "") : bazaGodzin,
        isDaty ? (p.dataStart || "") : bazaStart,
        isDaty ? (p.dataKoniec || "") : bazaKoniec
      ]);
      
      const startRow = participantSheet.getLastRow() + 1;
      const numRows = rowsToAppend.length;
      
      participantSheet.getRange(startRow, 2, numRows, 2).setNumberFormat('@'); 
      participantSheet.getRange(startRow, 1, numRows, 11).setValues(rowsToAppend); 
    }
    
    return { status: "success", trainingId: trainingId };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function updateTraining(trainingId, formData, participants) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_DB_ID);
    const trainingSheet = ss.getSheetByName('Szkolenia');
    const participantSheet = ss.getSheetByName('Uczestnicy_Szkolen');
    
    const trainingRows = trainingSheet.getDataRange().getValues();
    let rowIndex = -1;
    const targetId = getSafeTrainingId(trainingId); 
    
    for (let i = 1; i < trainingRows.length; i++) {
      if (getSafeTrainingId(trainingRows[i][0]) === targetId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { status: "error", message: "Nie znaleziono szkolenia do edycji." };
    
    trainingSheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(targetId);
    trainingSheet.getRange(rowIndex, 2, 1, 10).setValues([[ 
      formData.tytul,
      formData.typInstruktora,
      formData.planowanie,
      formData.formula,
      formData.typSzkolenia,
      formData.kategoria,
      formData.dataStart || "",
      formData.dataKoniec || "",
      formData.dlugosc || "",
      formData.certyfikat
    ]]);
    
    trainingSheet.getRange(rowIndex, 14, 1, 2).setValues([[formData.indywidualneGodziny, formData.indywidualneDaty]]);
    
    const participantRows = participantSheet.getDataRange().getValues();
    for (let j = participantRows.length - 1; j >= 1; j--) {
      if (getSafeTrainingId(participantRows[j][1]) === targetId) {
        participantSheet.deleteRow(j + 1);
      }
    }
    
    if (participants && participants.length > 0) {
      const isGodziny = formData.indywidualneGodziny;
      const isDaty = formData.indywidualneDaty;
      
      const bazaGodzin = formData.dlugosc || "";
      const bazaStart = formData.dataStart || "";
      const bazaKoniec = formData.dataKoniec || "";

      const rowsToAppend = participants.map(p => [
        Utilities.getUuid(), 
        targetId,
        formatujIdPracownika(p.id), 
        p.nazwisko,
        p.imie,
        p.stanowisko,
        p.dzial,
        p.zmiana,
        isGodziny ? (p.godziny || "") : bazaGodzin,
        isDaty ? (p.dataStart || "") : bazaStart,
        isDaty ? (p.dataKoniec || "") : bazaKoniec
      ]);
      
      const startRow = participantSheet.getLastRow() + 1;
      const numRows = rowsToAppend.length;
      
      participantSheet.getRange(startRow, 2, numRows, 2).setNumberFormat('@'); 
      participantSheet.getRange(startRow, 1, numRows, 11).setValues(rowsToAppend); 
    }
    
    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function deleteTraining(trainingId) {
  const ss = SpreadsheetApp.openById(MAIN_DB_ID);
  const trainingSheet = ss.getSheetByName('Szkolenia');
  const participantSheet = ss.getSheetByName('Uczestnicy_Szkolen');
  const targetId = getSafeTrainingId(trainingId);
  
  const trainingRows = trainingSheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < trainingRows.length; i++) {
    if (getSafeTrainingId(trainingRows[i][0]) === targetId) {
      rowIndex = i + 1; 
      break;
    }
  }
  
  if (rowIndex === -1) return { status: "error", message: "Nie znaleziono wybranego szkolenia." };
  
  const parts = targetId.split('/');
  if (parts.length === 3) {
    const year = parts[1];
    const currentNum = parseInt(parts[2]);
    
    let isLast = true;
    for (let i = 1; i < trainingRows.length; i++) {
      const rowIdStr = getSafeTrainingId(trainingRows[i][0]);
      if (rowIdStr && rowIdStr.startsWith(`TR/${year}/`)) {
        const num = parseInt(rowIdStr.split('/')[2]);
        if (num > currentNum) {
          isLast = false;
          break;
        }
      }
    }
    
    if (!isLast) {
      return { status: "error", message: "Możesz usunąć tylko ostatnie wprowadzone szkolenie z tego roku." };
    }
  }
  
  trainingSheet.deleteRow(rowIndex);
  
  const participantRows = participantSheet.getDataRange().getValues();
  for (let j = participantRows.length - 1; j >= 1; j--) {
    if (getSafeTrainingId(participantRows[j][1]) === targetId) {
      participantSheet.deleteRow(j + 1);
    }
  }
  
  return { status: "success" };
}

function getMainTableData() {
  const ss = SpreadsheetApp.openById(MAIN_DB_ID);
  let trainingRows = ss.getSheetByName('Szkolenia').getDataRange().getValues();
  let participantRows = ss.getSheetByName('Uczestnicy_Szkolen').getDataRange().getValues();
  
  trainingRows.shift(); 
  participantRows.shift(); 
  
  trainingRows = trainingRows.filter(r => r[0] !== "");
  participantRows = participantRows.filter(p => p[0] !== "");
  
  return trainingRows.map(r => {
    let id = getSafeTrainingId(r[0]);
    
    const participants = participantRows.filter(p => getSafeTrainingId(p[1]) === id).map(p => {
      let pStart = p[9];
      if (pStart instanceof Date) pStart = Utilities.formatDate(pStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
      else if (pStart) pStart = pStart.toString();

      let pEnd = p[10];
      if (pEnd instanceof Date) pEnd = Utilities.formatDate(pEnd, Session.getScriptTimeZone(), "yyyy-MM-dd");
      else if (pEnd) pEnd = pEnd.toString();

      return {
        relId: p[0],
        id: formatujIdPracownika(p[2]), 
        nazwisko: p[3],
        imie: p[4],
        stanowisko: p[5],
        dzial: p[6],
        zmiana: p[7],
        godziny: p[8] || "",
        dataStart: pStart || "",
        dataKoniec: pEnd || ""
      };
    });
    
    let dStart = r[7];
    if (dStart instanceof Date) dStart = Utilities.formatDate(dStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
    else if (dStart) dStart = dStart.toString();
    
    let dKoniec = r[8];
    if (dKoniec instanceof Date) dKoniec = Utilities.formatDate(dKoniec, Session.getScriptTimeZone(), "yyyy-MM-dd");
    else if (dKoniec) dKoniec = dKoniec.toString();
    
    const indywidualneGodziny = (String(r[13]).toLowerCase() === 'true');
    const indywidualneDaty = (String(r[14]).toLowerCase() === 'true');
    
    return {
      id: id,
      tytul: r[1],
      typInstruktora: r[2],
      planowanie: r[3],
      formula: r[4],
      typSzkolenia: r[5],
      kategoria: r[6],
      dataStart: dStart,
      dataKoniec: dKoniec,
      dlugosc: r[9],
      certyfikat: r[10],
      autor: r[11],
      hasEditAccess: true, // Zostawiamy dla spójności frontu
      indywidualneGodziny: indywidualneGodziny,
      indywidualneDaty: indywidualneDaty,
      participants: participants
    };
  });
}

function zaimportujPracownikowZ_CSV() {
  try {
    const plik = DriveApp.getFileById(CSV_FILE_ID);
    const csvString = plik.getBlob().getDataAsString('UTF-8');
    
    let separator = ';';
    const pierwszaLinia = csvString.split('\n')[0];
    if (pierwszaLinia.indexOf('\t') > -1 && pierwszaLinia.indexOf(';') === -1) {
      separator = '\t';
    }
    
    const csvDane = Utilities.parseCsv(csvString, separator); 
    
    const ss = SpreadsheetApp.openById(MAIN_DB_ID);
    const sheet = ss.getSheetByName('Pracownicy');
    sheet.clearContents(); 
    
    sheet.appendRow(['ID pracownika', 'Nazwisko', 'Imię', 'Stanowisko', 'Dział', 'Zmiana']);
    
    const gotoweDane = [];
    const aktualnyRok = new Date().getFullYear(); 
    
    for (let i = 1; i < csvDane.length; i++) {
      const wiersz = csvDane[i];
      
      if (wiersz.length >= 9 && wiersz[0] !== "") { 
        const dataZwolnienia = wiersz[7] ? wiersz[7].trim() : "";
        let doImportu = false;
        
        if (dataZwolnienia === "") {
          doImportu = true; 
        } else {
          const match = dataZwolnienia.match(/\d{4}/);
          if (match && parseInt(match[0]) >= aktualnyRok) {
            doImportu = true; 
          }
        }
        
        if (doImportu) {
          let zmiana = wiersz[8] ? wiersz[8].trim() : "";
          if (zmiana.startsWith("Standard")) {
            zmiana = "Standard";
          }
          
          gotoweDane.push([
            formatujIdPracownika(wiersz[0]), 
            wiersz[2], 
            wiersz[1], 
            wiersz[3], 
            wiersz[4], 
            zmiana     
          ]);
        }
      }
    }
    
    gotoweDane.push([
      '999999', 
      'AGENCJA', 
      'PRACY', 
      'BRAKARKA', 
      'WYDZIAŁ PRODUKCJI HUTNICZEJ', 
      'Standard'
    ]);
    
    if (gotoweDane.length > 0) {
      sheet.getRange(2, 1, gotoweDane.length, 1).setNumberFormat('@'); 
      sheet.getRange(2, 1, gotoweDane.length, 6).setValues(gotoweDane);
    }
    
    return { status: "success", count: gotoweDane.length };
    
  } catch (e) {
    Logger.log("Błąd automatycznego importu CSV: " + e.message);
    return { status: "error", message: e.message };
  }
}
