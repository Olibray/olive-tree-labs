// Deploy this as a Web App:
//   Extensions → Apps Script → paste this code → Deploy → New deployment
//   Type: Web app · Execute as: Me · Who has access: Anyone
//   Copy the deployment URL into baliaric-beats.html (SHEET_URL constant)

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Honeypot — silently succeed
    if (data._honey) {
      return json({ ok: true });
    }

    const ss = SpreadsheetApp.openById('1SMpJkjykNJoEDuZBNbAGHtuuXd9sAquvApU7idTJNHk');

    if (data.type === 'pledge') {
      const sheet = getOrCreateSheet(ss, 'Pledges',
        ['Timestamp', 'Name', 'Email', 'People', 'From', 'Artists']);
      sheet.appendRow([new Date(), data.name, data.email, data.people, data.from, data.artists]);

    } else if (data.type === 'artist') {
      const sheet = getOrCreateSheet(ss, 'Artist Requests',
        ['Timestamp', 'Artist Name', 'Instagram', 'Genre', 'Note']);
      sheet.appendRow([new Date(), data.artist_name, data.artist_instagram, data.artist_genre, data.artist_note]);
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
