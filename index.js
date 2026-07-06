const express = require('express');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

// פונקציית עזר להמרת מספר עמודה (1-based) לאותיות (A, B, C...)
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

app.all('/update-sheet', async (req, res) => {
  try {
    // מאפשר לקרוא את הפרמטרים גם מתוך ה-URL (קריאת GET) וגם מתוך ה-Body (קריאת POST)
    const {
      spreadsheetId,
      sheetName,
      searchColumnNum,
      updateColumnNum,
      searchValue,
      updateValue
    } = Object.keys(req.query).length ? req.query : req.body;

    // הגדרת החיבור באמצעות משתני סביבה (Environment Variables) ב-Render
    const auth = new google.auth.JWT(
      process.env.GOOGLE_CLIENT_EMAIL,
      null,
      process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });

    // 1. שליפת כל הנתונים מהגיליון כדי לחפש את השורה
    const readResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:ZZ`,
    });

    const rows = readResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).send('לא נמצאו נתונים בגיליון');
    }

    // חיפוש השורה המתאימה (אינדקס העמודה הוא מספר העמודה פחות 1)
    const searchIndex = parseInt(searchColumnNum) - 1;
    let rowIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][searchIndex] == searchValue) {
        rowIndex = i + 1; // גוגל שיטס מתחיל משורה 1
        break;
      }
    }

    if (rowIndex === -1) {
      return res.status(404).send('הערך לחיפוש לא נמצא בעמודה המבוקשת');
    }

    // המרת מספר עמודת העדכון לאות
    const updateLetter = columnToLetter(parseInt(updateColumnNum));
    const targetRange = `${sheetName}!${updateLetter}${rowIndex}`;

    // 2. עדכון הערך בתא שנמצא
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: targetRange,
      valueInputOption: 'RAW', // RAW מבטיח שהערך יישמר כטקסט פשוט ולא יפורש כמספר או תאריך
      requestBody: {
        values: [[updateValue]],
      },
    });

    // 3. אופציונלי: החלת עיצוב פורמט טקסט מפורש (Plain Text) על התא
    // הערה: שימוש ב-valueInputOption: 'RAW' לרוב מספיק, אך קוד זה מקבע את הפורמט בגיליון
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: 0, // אם יש מספר גיליונות, מומלץ להביא את ה-sheetId המדויק
                startRowIndex: rowIndex - 1,
                endRowIndex: rowIndex,
                startColumnIndex: parseInt(updateColumnNum) - 1,
                endColumnIndex: parseInt(updateColumnNum)
              },
              cell: {
                userEnteredFormat: {
                  numberFormat: {
                    type: 'TEXT'
                  }
                }
              },
              fields: 'userEnteredFormat.numberFormat'
            }
          }
        ]
      }
    });

    res.send({ status: 'success', message: `שורה ${rowIndex} עודכנה בהצלחה בעמודה ${updateLetter}` });

  } catch (error) {
    console.error(error);
    res.status(500).send('שגיאה בעיבוד הבקשה: ' + error.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
