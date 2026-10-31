const SHEET_NAME = "Expenses";
const PEOPLE_SHEET_NAME = "Members";

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, SHEET_NAME, ["id", "date", "description", "amount", "paidBy", "peopleJson", "createdAt"]);
    const membersSheet = getOrCreateSheet_(ss, PEOPLE_SHEET_NAME, ["member1", "member2", "member3", "member4", "member5"]);

    if (action === "list") {
      const rows = sheet.getDataRange().getValues();
      const expenses = rows.slice(1).filter(r => r[0] !== "").map(r => ({
        id: String(r[0]),
        date: formatDate_(r[1]),
        description: String(r[2]),
        amount: Number(r[3]),
        paidBy: String(r[4]),
        people: safeJson_(r[5], [])
      }));
      const memberRows = membersSheet.getDataRange().getValues();
      const people = memberRows.length > 1 ? memberRows[1].slice(0,5).map(String) : [];
      return json_({ok:true, expenses, people: people.length === 5 ? people : null});
    }

    if (action === "add") {
      const x = body.expense || {};
      if (!x.id || !x.description || !(Number(x.amount) > 0) || !x.paidBy) throw new Error("Missing or invalid expense details.");
      const people = Array.isArray(body.people) && body.people.length === 5 ? body.people : (x.people || []);
      sheet.appendRow([String(x.id), String(x.date || ""), String(x.description), Number(x.amount), String(x.paidBy), JSON.stringify(people), new Date()]);
      if (people.length === 5) {
        membersSheet.getRange(2,1,1,5).setValues([people.map(String)]);
      }
      return json_({ok:true, expenses: readExpenses_(sheet)});
    }

    if (action === "delete") {
      const id = String(body.id || "");
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]) === id) {
          sheet.deleteRow(i + 1);
          return json_({ok:true});
        }
      }
      throw new Error("Expense not found.");
    }
    throw new Error("Unknown action.");
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}

function doGet() {
  return json_({ok:true, message:"Trip Expense Manager API is running. Use POST requests from index.html."});
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function readExpenses_(sheet) {
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => r[0] !== "").map(r => ({
    id:String(r[0]), date:formatDate_(r[1]), description:String(r[2]),
    amount:Number(r[3]), paidBy:String(r[4]), people:safeJson_(r[5], [])
  }));
}

function safeJson_(value, fallback) {
  try { return JSON.parse(String(value || "")); } catch (_) { return fallback; }
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(value || "");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
