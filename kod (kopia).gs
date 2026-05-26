function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var kod = data.kod ? data.kod.toString().trim() : "";
    var action = data.action;

    // --- LOGIKA POBIERANIA HISTORII ---

    // --- NOWA SEKCJA: POBIERANIE UNIKALNEJ LISTY PRACOWNIKÓW ---
    if (action === "get_employee_list") {
      var targetMonth = data.miesiac;
      var targetYear = data.rok.toString();
      var rows = sheet.getDataRange().getDisplayValues();
      
      var employeesSet = {};
      var results = [];
      
      for (var i = 1; i < rows.length; i++) {
        var rowName = rows[i][0].toString().trim();
        var rowYear = rows[i][1].toString();
        var rowMonth = rows[i][2].toString();
        
        // Jeśli zgadza się rok i miesiąc, a nazwisko nie jest puste
        if (rowYear === targetYear && rowMonth === targetMonth && rowName !== "") {
          // Używamy obiektu jako Set, aby uniknąć duplikatów (wielkość liter nie ma znaczenia przy sprawdzaniu)
          var key = rowName.toLowerCase();
          if (!employeesSet[key]) {
            employeesSet[key] = true;
            results.push(rowName); // Zachowujemy oryginalną pisownię
          }
        }
      }
      
      // Sortujemy listę osób alfabetycznie
      results.sort();
      
      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }
// ... fragment funkcji doPost(e) obsługujący "get_history" ...

    if (action === "get_history") {
      var targetMonth = data.miesiac; 
      var targetYear = data.rok.toString();
      var targetName = kod.toLowerCase();
      
      var rows = sheet.getDataRange().getDisplayValues();
      var results = [];
      
      for (var i = 1; i < rows.length; i++) {
        var rowName = rows[i][0].toString().toLowerCase();
        var rowYear = rows[i][1].toString();
        var rowMonth = rows[i][2].toString();
        
        if (rowName === targetName && rowYear === targetYear && rowMonth === targetMonth) {
          var fullStart = rows[i][3]; // "15.05.2026 15:46:02"
          var fullEnd = rows[i][4];
          
          // Wyciągamy dzień (pierwsza część daty przed kropką)
          var dzienNr = parseInt(fullStart.split(".")[0]);

          results.push({
            dzien: dzienNr,
            rozpoczecie: fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart,
            zakonczenie: fullEnd.includes(" ") ? fullEnd.split(" ")[1] : (fullEnd || "---"),
            czas: rows[i][5] || "---"
          });
        }
      }

      // SORTOWANIE: Od 1 do 31
      results.sort(function(a, b) {
        return a.dzien - b.dzien;
      });

      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }
// ... reszta kodu bez zmian ...
// --- NOWA SEKCJA: POBIERANIE 30 OSTATNICH WPISÓW ---
// --- POBIERANIE DANYCH DO OSTATNIE.HTML ---
    if (action === "get_recent") {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      
      var maxRows = Math.min(30, lastRow - 1);
      var rows = sheet.getRange(2, 1, maxRows, 6).getDisplayValues();
      var results = [];
      
      for (var i = 0; i < rows.length; i++) {
        var fullStart = rows[i][3]; // "17.05.2026 11:40:00"
        var fullEnd = rows[i][4];
        
        results.push({
          nazwisko: rows[i][0],
          data: fullStart.includes(" ") ? fullStart.split(" ")[0] : "---", // NOWOŚĆ: Pobieramy datę (np. 17.05.2026)
          rozpoczecie: fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart,
          zakonczenie: fullEnd.includes(" ") ? fullEnd.split(" ")[1] : (fullEnd || "---"),
          czas: rows[i][5] || "---"
        });
      }
      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }

// --- POBIERANIE CAŁEGO MIESIĄCA DO RAPORT.HTML ---
    if (action === "get_month_data") {
      var targetMonth = data.miesiac;
      var targetYear = data.rok.toString();
      var rows = sheet.getDataRange().getDisplayValues();
      var results = [];

      for (var i = 1; i < rows.length; i++) {
        var rowYear = rows[i][1].toString();
        var rowMonth = rows[i][2].toString();

        if (rowYear === targetYear && rowMonth === targetMonth) {
          var fullStart = rows[i][3]; // "17.05.2026 11:40:00"
          var fullEnd = rows[i][4];
          
          results.push({
            nazwisko: rows[i][0],
            data: fullStart.includes(" ") ? fullStart.split(" ")[0] : "---",
            rozpoczecie: fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart,
            zakonczenie: fullEnd.includes(" ") ? fullEnd.split(" ")[1] : (fullEnd || "---"),
            czas: rows[i][5] || "---"
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }

    // --- TWOJA ISTNIEJĄCA LOGIKA ZAPISU (BEZ ZMIAN) ---
    var datetimeStr = data.datetime; 
    var parts = datetimeStr.split(" ");
    var dateParts = parts[0].split(".");
    var timeParts = parts[1].split(":");
    var dateObj = new Date(dateParts[2], dateParts[1] - 1, dateParts[0], timeParts[0], timeParts[1], timeParts[2]);

    var rok = dateObj.getFullYear();
    var miesiace = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec", "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"];
    var nazwaMiesiaca = miesiace[dateObj.getMonth()];

    if (action === "rozpoczecie") {
      sheet.insertRowBefore(2);
      sheet.getRange(2, 2).setNumberFormat("@"); 
      sheet.getRange(2, 1, 1, 4).setValues([[kod, rok, nazwaMiesiaca, datetimeStr]]);
      return ContentService.createTextOutput(JSON.stringify({ "status": "success" })).setMimeType(ContentService.MimeType.JSON);
                           
    } else if (action === "zakonczenie") {
      var dataRange = sheet.getDataRange().getValues();
      var updated = false;

      for (var i = 1; i < dataRange.length; i++) {
        var rowKod = dataRange[i][0].toString().trim();
        var rowZakonczenie = dataRange[i][4];
        if (rowKod === kod && (!rowZakonczenie)) {
          var rowNum = i + 1; 
          sheet.getRange(rowNum, 5).setValue(datetimeStr);
          sheet.getRange(rowNum, 6).setFormula("=E" + rowNum + "-D" + rowNum);
          updated = true;
          break;
        }
      }

      if (updated) {
        return ContentService.createTextOutput(JSON.stringify({ "status": "success" })).setMimeType(ContentService.MimeType.JSON);
      } else {
        sheet.insertRowBefore(2);
        sheet.getRange(2, 1, 1, 6).setValues([[kod, rok, nazwaMiesiaca, "BRAK!", datetimeStr, "BŁĄD"]]);
        return ContentService.createTextOutput(JSON.stringify({ "status": "warning" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
                                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
