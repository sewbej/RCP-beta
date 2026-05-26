function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var kod = data.kod ? data.kod.toString().trim() : "";
    var action = data.action;
    
    // Pobranie identyfikatora urządzenia z przesłanych danych (jeśli brak w pakiecie, przypisuje pusty tekst)
    var deviceId = data.deviceId ? data.deviceId.toString().trim() : "";
// --- MODYFIKACJA AUTORYZACJI ---
  // Definiujemy tablicę akcji, które są DOSTĘPNE DLA KAŻDEGO bez autoryzacji
  var publicActions = ["get_history"]; 

  // Jeśli akcja NIE JEST publiczna, wtedy i TYLKO WTEDY sprawdzamy urządzenie
  if (publicActions.indexOf(action) === -1) {
    
    // =========================================================================
    // BEZWZGLĘDNE ZABEZPIECZENIE: Wszystkie akcje (odczyt i zapis) wymagają ID
    // =========================================================================
    if (deviceId !== "9DgU75ztF3Ck") {
      return ContentService.createTextOutput(JSON.stringify({ 
        "status": "error", 
        "message": "Nieautoryzowane urządzenie! Brak dostępu do systemu." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

  }

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
        
        if (rowYear === targetYear && rowMonth === targetMonth && rowName !== "") {
          var key = rowName.toLowerCase();
          if (!employeesSet[key]) {
            employeesSet[key] = true;
            results.push(rowName);
          }
        }
      }
      
      results.sort();
      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }

    // --- SEKCJA: POBIERANIE HISTORII INDYWIDUALNEJ ---
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
          var fullStart = rows[i][3];
          var fullEnd = rows[i][4];
          
          var dzienNr = parseInt(fullStart.split(".")[0]);

          var startStr = fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart;
          var startHHMM = startStr.includes(":") ? startStr.split(":").slice(0, 2).join(":") : startStr;

          var endStr = fullEnd.includes(" ") ? fullEnd.split(" ")[1] : fullEnd;
          var endHHMM = endStr.includes(":") ? endStr.split(":").slice(0, 2).join(":") : endStr;

          var rawTime = rows[i][5] || "---";
          var timeHHMM = (rawTime !== "---" && rawTime.includes(":")) ? rawTime.split(":").slice(0, 2).join(":") : rawTime;

          results.push({
            dzien: dzienNr,
            rozpoczecie: startHHMM,
            zakonczenie: (endHHMM && endHHMM !== "") ? endHHMM : "---",
            czas: timeHHMM
          });
        }
      }

      results.sort(function(a, b) {
        return a.dzien - b.dzien;
      });

      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }

    // --- POBIERANIE DANYCH DO OSTATNIE.HTML ---
    if (action === "get_recent") {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      
      var maxRows = Math.min(30, lastRow - 1);
      var rows = sheet.getRange(2, 1, maxRows, 6).getDisplayValues();
      var results = [];
      
      for (var i = 0; i < rows.length; i++) {
        var fullStart = rows[i][3];
        var fullEnd = rows[i][4];

        var startStr = fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart;
        var startHHMM = startStr.includes(":") ? startStr.split(":").slice(0, 2).join(":") : startStr;

        var endStr = fullEnd.includes(" ") ? fullEnd.split(" ")[1] : fullEnd;
        var endHHMM = endStr.includes(":") ? endStr.split(":").slice(0, 2).join(":") : endStr;

        var rawTime = rows[i][5] || "---";
        var timeHHMM = (rawTime !== "---" && rawTime.includes(":")) ? rawTime.split(":").slice(0, 2).join(":") : rawTime;

        results.push({
          nazwisko: rows[i][0],
          data: fullStart.includes(" ") ? fullStart.split(" ")[0] : "---", 
          rozpoczecie: startHHMM,
          zakonczenie: (endHHMM && endHHMM !== "") ? endHHMM : "---",
          czas: timeHHMM
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
          var fullStart = rows[i][3];
          var fullEnd = rows[i][4];

          var startStr = fullStart.includes(" ") ? fullStart.split(" ")[1] : fullStart;
          var startHHMM = startStr.includes(":") ? startStr.split(":").slice(0, 2).join(":") : startStr;

          var endStr = fullEnd.includes(" ") ? fullEnd.split(" ")[1] : fullEnd;
          var endHHMM = endStr.includes(":") ? endStr.split(":").slice(0, 2).join(":") : endStr;

          var rawTime = rows[i][5] || "---";
          var timeHHMM = (rawTime !== "---" && rawTime.includes(":")) ? rawTime.split(":").slice(0, 2).join(":") : rawTime;

          results.push({
            nazwisko: rows[i][0],
            data: fullStart.includes(" ") ? fullStart.split(" ")[0] : "---", 
            rozpoczecie: startHHMM,
            zakonczenie: (endHHMM && endHHMM !== "") ? endHHMM : "---",
            czas: timeHHMM
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
    }

    // =========================================================================
    // LOGIKA ZAPISU (ROZPOCZĘCIE / ZAKOŃCZENIE)
    // =========================================================================
    var datetimeStr = data.datetime; 
    var parts = datetimeStr.split(" ");
    var dateParts = parts[0].split(".");
    var timeParts = parts[1].split(":");
    
    var sekundy = timeParts[2] ? parseInt(timeParts[2]) : 0;
    var dateObj = new Date(dateParts[2], dateParts[1] - 1, dateParts[0], timeParts[0], timeParts[1], sekundy);

    if (isNaN(dateObj.getTime())) {
      throw new Error("Błędny format daty: " + datetimeStr);
    }

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
        sheet.getRange(2, 1, 1, 6).setValues([[kod, rok, nazwaMiesiaca, "BRAK!", datetimeStr, ""]]);
        sheet.getRange(2, 6).setFormula("=E2-D2");
        
        return ContentService.createTextOutput(JSON.stringify({ "status": "warning" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
                                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
