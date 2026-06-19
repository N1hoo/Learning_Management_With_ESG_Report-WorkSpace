# System Zarządzania Szkoleniami (LMS) dla Google Workspace

Lekki, szybki i w pełni bezserwerowy system do zarządzania szkoleniami pracowniczymi (LMS) zbudowany w architekturze **Google Apps Script (GAS)**, wykorzystujący Google Sheets jako bazę danych. 

System pozwala na kompleksowe planowanie szkoleń, śledzenie historii pracowników, inteligentny import list obecności (bezpośrednio od dostawców zewnętrznych) oraz generowanie natywnych raportów w formacie Microsoft Excel (.xlsx) na potrzeby audytów **ESG**.

## Główne funkcjonalności

*   **Natywny Eksport ESG (Excel):** Generowanie gotowych raportów `.xlsx` (wymuszanie formatu z pominięciem natywnych blokad Google PDF) z filtrowaniem po określonym roczniku.
*   **Inteligentny Import z Excela (.xlsx, .ods):** Moduł potrafiący czytać zewnętrzne listy obecności wewnątrz przeglądarki (Dzięki `SheetJS`). Używa autorskiego algorytmu typu *Fuzzy Matching* do automatycznego parowania imion i nazwisk z pliku (nawet z literówkami!) z wewnętrzną bazą kadr. Pozwala zaktualizować istniejące szkolenie lub wygenerować całkiem nowe.
*   **Wielojęzyczność (i18n):** Wbudowana, w pełni frontendowa obsługa 4 języków (Polski, Angielski, Włoski, Francuski). System automatycznie wykrywa język przeglądarki (np. nagłówek `Accept-Language`) i zapamiętuje ręczny wybór użytkownika w `localStorage`.
*   **Drukowanie List Obecności:** Generowanie w locie estetycznych, gotowych do wydruku list z danymi uczestników (sortowanymi alfabetycznie) i pustą kolumną na fizyczny podpis (pod kątem audytów).
*   **Indywidualne rozliczanie czasu i dat:** Elastyczność pozwalająca grupie mieć wspólny termin i czas, lub – przy szkoleniach językowych – każdemu pracownikowi przypisać własny przedział czasowy i zindywidualizowaną liczbę godzin.
*   **Wirtualna Karta Pracownika:** Dedykowany moduł pozwalający wyszukać osobę po nazwisku/ID i wyświetlić jej pełną, historyczną tabelę odbytych kursów wraz ze zsumowanym czasem wszystkich szkoleń.
*   **Kontrola dostępu ("Bramkarz"):** Prosty i skuteczny system autoryzacji sprawdzający adres e-mail zalogowanego użytkownika Google z białą listą (Whitelist) zapisaną w bazie.

## Stack Technologiczny

*   **Backend:** Google Apps Script (JavaScript/ES6)
*   **Frontend:** HTML5, CSS3, JavaScript
*   **Framework UI:** Bootstrap 5.3 + Bootstrap Icons
*   **Baza Danych:** Google Sheets (Zdenormalizowana, oddzielne arkusze relacyjne)
*   **Zewnętrzne Biblioteki:** [SheetJS](https://sheetjs.com/) (Do parsowania i importu Excel/ODS na froncie)

## Instrukcja wdrożenia

Z racji tego, że system oparty jest o infrastrukturę Google, nie wymaga zewnętrznego hostingu ani serwera.

1. Utwórz nowy arkusz w Google Sheets (to będzie Twoja baza danych).
2. W arkuszu przygotuj zakładki o nazwach: `Szkolenia`, `Uczestnicy_Szkolen`, `Pracownicy`, `Slowniki`, `Tlumacz`, `Uprawnieni`.
3. Z górnego menu arkusza wybierz **Rozszerzenia -> Apps Script**.
4. W edytorze kodu utwórz pliki zgodnie ze strukturą tego repozytorium:
   * `Kod.gs` (Główna logika i routing)
   * `Export.gs` (Silnik generowania raportów Excel i wywoływania API)
   * `Import.gs` (Backendowe dopisywanie zaimportowanych masowo danych)
   * `Index.html` (Główny plik Frontendowy - widoki, modale, i18n)
5. W pliku `Kod.gs` zaktualizuj stałą `MAIN_DB_ID` wklejając ID swojego arkusza Google.
6. Kliknij **Wdróż -> Nowe wdrożenie -> Aplikacja internetowa**.
   * Wykonaj jako: *Ja*
   * Kto ma dostęp: *Ktokolwiek z moim kontem Google / Organizacji*
7. Przejdź jednorazowy proces autoryzacji skryptu (Zezwolenie na łączenie się z zewnętrznymi usługami ze względu na API eksportu).

## Architektura Plików

*   **`Index.html`**: Monolityczny plik frontendowy (SPA - Single Page Application). Zawiera strukturę HTML, style CSS pod wydruk oraz skomplikowaną logikę JS obsługującą widoki (Tabela Główna / Karta Pracownika), translacje, modale i filtry wyszukiwania.
*   **`Kod.gs`**: Odbiera żądania HTTP, sprawdza autoryzację w `doGet()`, obsługuje zapytania do bazy arkusza (CRUD) oraz zarządza generatorami unikalnych ID dla szkoleń (w formacie `TR/RRRR/Lp`).
*   **`Export.gs`**: Obejście domyślnej konwersji Google do PDF. Pobiera dane, układa raport ESG do tymczasowego pliku, wysyła autoryzowane zapytanie (Endpoint API) z prośbą o wymuszenie `format=xlsx` i zwraca użytkownikowi ciąg Base64 generujący plik do pobrania.
*   **`Import.gs`**: Przyjmuje wstępnie obrobione przez przeglądarkę dane, mapuje ID pracowników na pozostałe ich dane strukturalne z Bazy Kadr i hurtowo wkleja (`setValues()`) na sam dół arkusza `Uczestnicy_Szkolen`.

---
*Stworzono w celu ratowania czasu działów HR.*
