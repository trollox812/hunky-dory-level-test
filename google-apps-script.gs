function doPost(e) {
  try {
    const sheet = SpreadsheetApp
      .openById("1DQKketDHjco_qsyc4jTFfALkXLomlij_srDC5gf0JyA")
      .getSheetByName("Sheet1");

    const data = JSON.parse(e.postData.contents);

    const headers = [
      "Timestamp",
      "Student Name",
      "Age",
      "Class",
      "Best Level",
      "Last Passed",
      "Level Failed",
      "A0 Score",
      "A1 Score",
      "A2 Score",
      "B1 Score",
      "B2 Score",
      "Cumulative %",
      "Written Response"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    let scores = {
      A0: "",
      A1: "",
      A2: "",
      B1: "",
      B2: ""
    };

    if (Array.isArray(data.roundScores)) {
      data.roundScores.forEach(r => {
        const code = r.levelCode;
        const total = Number(r.total || 0);
        const correct = Number(r.correct || 0);

        if (scores.hasOwnProperty(code) && total > 0) {
          scores[code] = Math.round((correct / total) * 100) + "%";
        }
      });
    }

    sheet.appendRow([
      new Date(),
      data.studentName || "",
      data.age || "",
      data.class || "",
      data.bestLevel || "",
      data.lastPassed || "",
      data.levelFailed || "",
      scores.A0,
      scores.A1,
      scores.A2,
      scores.B1,
      scores.B2,
      data.cumulativePercent || "",
      data.writtenResponse || data.writingResponseText || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
