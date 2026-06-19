/**
 * Skrypt obsługujący import danych z Excela do istniejących szkoleń
 */

function dopiszGodzinyDoSzkolenia(trainingId, zaimportowaniUczestnicy) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_DB_ID); // Zmienna pobrana z Kod.gs
    const szkoleniaSheet = ss.getSheetByName('Szkolenia');
    const uczestnicySheet = ss.getSheetByName('Uczestnicy_Szkolen');
    const pracownicySheet = ss.getSheetByName('Pracownicy');
    
    // 1. Sprawdzamy czy szkolenie w ogóle istnieje i pobieramy jego daty
    const tRows = szkoleniaSheet.getDataRange().getValues();
    let tExists = false;
    let tStart = "";
    let tEnd = "";
    
    for(let i = 1; i < tRows.length; i++) {
       if(getSafeTrainingId(tRows[i][0]) === trainingId) {
          tExists = true;
          tStart = tRows[i][7];
          tEnd = tRows[i][8];
          break;
       }
    }
    
    if(!tExists) throw new Error("Nie znaleziono szkolenia o ID: " + trainingId);

    // 2. Pobieramy bazę pracowników, żeby mieć ich szczegóły (Dział, Stanowisko itp.) do wpisania
    const pracRows = pracownicySheet.getDataRange().getValues();
    const pracMap = {};
    for(let i = 1; i < pracRows.length; i++) {
       if(pracRows[i][0]) {
         pracMap[formatujIdPracownika(pracRows[i][0])] = {
           nazwisko: pracRows[i][1],
           imie: pracRows[i][2],
           stanowisko: pracRows[i][3],
           dzial: pracRows[i][4],
           zmiana: pracRows[i][5]
         };
       }
    }

    // 3. Pobieramy aktualnych uczestników tego konkretnego szkolenia, żeby wiedzieć, komu zaktualizować godziny
    const uRows = uczestnicySheet.getDataRange().getValues();
    const obecniMap = {}; // Format: { "ID_Pracownika": rowIndex }
    
    for(let i = 1; i < uRows.length; i++) {
       if(getSafeTrainingId(uRows[i][1]) === trainingId) {
          obecniMap[formatujIdPracownika(uRows[i][2])] = i;
       }
    }

    const rowsToAppend = [];
    
    // 4. Przetwarzamy zaimportowanych uczestników z okienka przeglądarki
    zaimportowaniUczestnicy.forEach(imp => {
       let empId = formatujIdPracownika(imp.id);
       let dodawaneGodziny = parseFloat(imp.godziny) || 0;
       
       if (obecniMap[empId] !== undefined) {
           // PRACOWNIK JEST JUZ W TYM SZKOLENIU -> Aktualizujemy mu godziny
           let rIndex = obecniMap[empId];
           let obecneGodziny = parseFloat(uRows[rIndex][8]) || 0; 
           let noweGodziny = obecneGodziny + dodawaneGodziny;
           
           // rIndex to indeks w tablicy (gdzie nagłówek = 0). W arkuszu nagłówek to 1.
           // Zatem wiersz w arkuszu to rIndex + 1. Kolumna godzin to I (indeks 9).
           uczestnicySheet.getRange(rIndex + 1, 9).setValue(noweGodziny);
       } else {
           // PRACOWNIKA NIE MA W TYM SZKOLENIU -> Dodajemy nowy wiersz
           let pData = pracMap[empId];
           if(pData) {
              rowsToAppend.push([
                Utilities.getUuid(),
                trainingId,
                empId,
                pData.nazwisko,
                pData.imie,
                pData.stanowisko,
                pData.dzial,
                pData.zmiana,
                dodawaneGodziny,
                tStart || "", 
                tEnd || ""
              ]);
           }
       }
    });
    
    // 5. Zbiorcze dopisanie nowych osób na sam dół arkusza (jeśli jacyś byli)
    if(rowsToAppend.length > 0) {
       const startRow = uczestnicySheet.getLastRow() + 1;
       const numRows = rowsToAppend.length;
       uczestnicySheet.getRange(startRow, 2, numRows, 2).setNumberFormat('@'); 
       uczestnicySheet.getRange(startRow, 1, numRows, 11).setValues(rowsToAppend); 
    }

    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
