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
  // מגדיר מראש את פורמט התשובה כטקסט פשוט בעברית עבור ימות המשיח
  res.type('text/plain; charset=utf-8');

  try {
    // קליטת הפרמטרים בדיוק לפי השמות של הסקריפט המקורי שלך
    const {
      sheetId,
      sheetName,
      searchColumn,
      updateColumn,
      searchValue,
      updateValue
    } = { ...req.body, ...req.query };

    // גיבוי למקרה שהמשתמש שלח spreadsheetId במקום sheetId
    const spreadsheetId = sheetId || req.query.spreadsheetId || req.body.spreadsheetId;

    if (!spreadsheetId || !sheetName || !searchColumn || !updateColumn) {
      return res.send("id_list_message=t-שגיאה בפרמטרים של המערכת,&");
    }

    // הגדרת החיבור באמצעות משתני סביבה ב-Render
    const auth = new google.auth.JWT(
      process.env.GOOGLE_CLIENT_EMAIL,
      null,
      process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });

    // 1. שליפת כל הנתונים מהגיליון כדי לחפש את הערך
    const readResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:ZZ`,
    });

    const rows = readResponse.data.values || [];
    const lastRow = rows.length;

    // הגנה למקרה שלא נשלח בכלל ערך לחיפוש (searchValue)
    if (searchValue === undefined || searchValue === "") {
      return res.send("read=t-נא הקש את מספר הזיהוי שלך וסולמית=searchValue,,7,7,,NO,yes,yes,,,&");
    }

    const searchIndex = parseInt(searchColumn) - 1;
    const updateColumnNum = parseInt(updateColumn);
    let foundIndex = -1;

    // חיפוש הערך בטבלה
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][searchIndex] !== undefined && String(rows[i][searchIndex]) === String(searchValue)) {
        foundIndex = i + 1; // גוגל שיטס מתחיל משורה 1
        break;
      }
    }

    // --- מצב א': נמצאה התאמה ל-searchValue ---
    if (foundIndex > -1) {
      
      if (updateValue === undefined || updateValue === "") {
        // שליפת הערך הנוכחי שנמצא בתא העדכון
        const currentValue = rows[foundIndex - 1][updateColumnNum - 1];
        
        if (currentValue !== "" && currentValue !== null && currentValue !== undefined) {
          return res.send("read=t-רשום עכשיו " + currentValue + ",להחלפה הקש את הערך הרצוי וסולמית=updateValue,no,10,9,,NO,yes,yes,,,&");
        } else {
          return res.send("read=t-אין ערך,לרישום הקש את הערך הרצוי וסולמית=updateValue,,10,9,,NO,yes,yes,,,&");
        }
        
      } else {
        // עדכון ערך קיים בשורה שנמצאה
        const oldValue = rows[foundIndex - 1][updateColumnNum - 1];
        const updateLetter = columnToLetter(updateColumnNum);
        const targetRange = `${sheetName}!${updateLetter}${foundIndex}`;

        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: targetRange,
          valueInputOption: 'RAW',
          requestBody: { values: [[updateValue]] },
        });

        const successMessage = (oldValue === "" || oldValue == null || oldValue === undefined) 
          ? "המספר נרשם בהצלחה" 
          : "המספר הוחלף בהצלחה";

        return res.send("id_list_message=t-" + successMessage + ",&");
      }
      
    } else {
      // --- מצב ב': לא נמצאה התאמה (משתמש חדש) ---
      
      // אם הוא לא נמצא, אבל הוא כבר הקיש את הנתון לעדכון (בסבב השני של הטלפון) -> נרשום אותו בשורה חדשה
      if (updateValue !== undefined && updateValue !== "") {
        const nextRow = (lastRow === 0) ? 1 : lastRow + 1;
        
        // יצירת מערך שורה חדש ריק בגודל המתאים
        const maxColumn = Math.max(parseInt(searchColumn), updateColumnNum);
        const newRowData = new Array(maxColumn).fill("");
        newRowData[parseInt(searchColumn) - 1] = updateValue; // ערך לחיפוש
        newRowData[updateColumnNum - 1] = updateValue; // ערך לעדכון
        
        const targetRange = `${sheetName}!A${nextRow}`;

        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: targetRange,
          valueInputOption: 'RAW',
          requestBody: { values: [[updateValue]] }, // זמנית מעדכן נקודתית דרך הטווח של עמודת העדכון
        });
        
        // נעדכן את שתי העמודות בצורה בטוחה
        const searchLetter = columnToLetter(parseInt(searchColumn));
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!${searchLetter}${nextRow}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[searchValue]] },
        });

        const updateLetter = columnToLetter(updateColumnNum);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!${updateLetter}${nextRow}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[updateValue]] },
        });

        return res.send("id_list_message=t-המספר נרשם בהצלחה,&");
      }
      
      // אם הוא לא נמצא ועדיין לא ביקשנו ממנו את הערך לעדכון -> נבקש אותו כעת
      return res.send("read=t-לא מצאנו אותך,לרישום מחדש נא הקש את הערך הרצוי וסולמית או נתק את השיחה והתקשר מחדש להכנסת מספר קיים=updateValue,no,10,9,,NO,yes,yes,,,&");
    }

  } catch (error) {
    console.error(error);
    return res.send("id_list_message=t-שגיאה זמנית במערכת,&");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
