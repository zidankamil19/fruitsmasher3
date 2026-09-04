function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    if (data.action === "saveScore") {
      sheet.appendRow([data.username, data.whatsapp, data.score, new Date()]);
      
      var output = JSON.stringify({ result: "success" });
      return ContentService.createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    var errOutput = JSON.stringify({ result: "error", message: err.toString() });
    return ContentService.createTextOutput(errOutput)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
    // Urutkan berdasarkan Skor (Kolom C / Index 2) terbanyak
    rows.sort(function(a, b) { 
      var scoreA = Number(a[2]) || 0;
      var scoreB = Number(b[2]) || 0;
      return scoreB - scoreA; 
    });
    
    // Ambil Top 5 Shinigami
    var leaderboard = rows.slice(0, 5).map(function(row) {
      return { 
        username: String(row[0] || "Anonim"), 
        score: Number(row[2]) || 0 
      };
    });
    
    return ContentService.createTextOutput(JSON.stringify(leaderboard))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
