const EXPENSE_SHEET = "TripExpenses";
const PAYMENT_SHEET = "Payments";
const SETTINGS_SHEET = "Settings";


function setupSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* EXPENSE SHEET */

  let expenseSheet =
    ss.getSheetByName(EXPENSE_SHEET);

  if (!expenseSheet) {

    expenseSheet =
      ss.insertSheet(EXPENSE_SHEET);

  }

  if (expenseSheet.getLastRow() === 0) {

    expenseSheet.appendRow([
      "ID",
      "Date",
      "Description",
      "Amount",
      "Paid By",
      "Person 1",
      "Person 2",
      "Person 3",
      "Person 4",
      "Person 5"
    ]);

  }


  /* PAYMENT SHEET */

  let paymentSheet =
    ss.getSheetByName(PAYMENT_SHEET);

  if (!paymentSheet) {

    paymentSheet =
      ss.insertSheet(PAYMENT_SHEET);

  }

  if (paymentSheet.getLastRow() === 0) {

    paymentSheet.appendRow([
      "ID",
      "Date",
      "From",
      "To",
      "Amount"
    ]);

  }


  /* SETTINGS SHEET */

  let settingsSheet =
    ss.getSheetByName(SETTINGS_SHEET);

  if (!settingsSheet) {

    settingsSheet =
      ss.insertSheet(SETTINGS_SHEET);

  }


  if (settingsSheet.getLastRow() === 0) {

    settingsSheet.appendRow([
      "Setting",
      "Value"
    ]);

    settingsSheet.appendRow([
      "Person 1",
      "Person 1"
    ]);

    settingsSheet.appendRow([
      "Person 2",
      "Person 2"
    ]);

    settingsSheet.appendRow([
      "Person 3",
      "Person 3"
    ]);

    settingsSheet.appendRow([
      "Person 4",
      "Person 4"
    ]);

    settingsSheet.appendRow([
      "Person 5",
      "Person 5"
    ]);

  }


  return "Sheets created successfully.";

}


/* GET */

function doGet(e) {

  return jsonResponse({
    ok:true,
    message:"Trip Expense Manager API is running."
  });

}


/* POST */

function doPost(e) {

  try {

    const data =
      JSON.parse(
        e.postData.contents
      );

    const action =
      data.action;


    if (action === "list") {

      return jsonResponse(
        getAllData()
      );

    }


    if (action === "add") {

      return jsonResponse(
        addExpense(
          data.expense,
          data.people
        )
      );

    }


    if (action === "delete") {

      return jsonResponse(
        deleteExpense(
          data.id
        )
      );

    }


    if (action === "addPayment") {

      return jsonResponse(
        addPayment(
          data.payment
        )
      );

    }


    if (action === "deletePayment") {

      return jsonResponse(
        deletePayment(
          data.id
        )
      );

    }


    return jsonResponse({

      ok:false,

      error:
        "Unknown action: " + action

    });

  }
  catch(error) {

    return jsonResponse({

      ok:false,

      error:error.message

    });

  }

}


/* GET ALL DATA */

function getAllData() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const expenseSheet =
    ss.getSheetByName(
      EXPENSE_SHEET
    );


  const paymentSheet =
    ss.getSheetByName(
      PAYMENT_SHEET
    );


  const expenses = [];

  const payments = [];


  /* EXPENSES */

  if (
    expenseSheet &&
    expenseSheet.getLastRow() > 1
  ) {

    const rows =
      expenseSheet
      .getRange(
        2,
        1,
        expenseSheet.getLastRow() - 1,
        10
      )
      .getValues();


    rows.forEach(row => {

      if (!row[0]) return;


      expenses.push({

        id:String(row[0]),

        date:formatDate(row[1]),

        description:String(
          row[2] || ""
        ),

        amount:Number(
          row[3] || 0
        ),

        paidBy:String(
          row[4] || ""
        ),

        people:[
          row[5],
          row[6],
          row[7],
          row[8],
          row[9]
        ].filter(
          x => String(x).trim() !== ""
        )

      });

    });

  }


  /* PAYMENTS */

  if (
    paymentSheet &&
    paymentSheet.getLastRow() > 1
  ) {

    const rows =
      paymentSheet
      .getRange(
        2,
        1,
        paymentSheet.getLastRow() - 1,
        5
      )
      .getValues();


    rows.forEach(row => {

      if (!row[0]) return;


      payments.push({

        id:String(row[0]),

        date:formatDate(row[1]),

        from:String(
          row[2] || ""
        ),

        to:String(
          row[3] || ""
        ),

        amount:Number(
          row[4] || 0
        )

      });

    });

  }


  return {

    ok:true,

    expenses,

    payments,

    people:getPeople()

  };

}


/* GET PEOPLE */

function getPeople() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );


  if (!sheet) {

    return [
      "Person 1",
      "Person 2",
      "Person 3",
      "Person 4",
      "Person 5"
    ];

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [
      "Person 1",
      "Person 2",
      "Person 3",
      "Person 4",
      "Person 5"
    ];

  }


  const rows =
    sheet
    .getRange(
      2,
      2,
      Math.min(
        5,
        lastRow - 1
      ),
      1
    )
    .getValues();


  const people =
    rows.map(
      row => String(row[0])
    );


  while(
    people.length < 5
  ){

    people.push(
      "Person " +
      (people.length + 1)
    );

  }


  return people;

}


/* SAVE PEOPLE */

function savePeople(
  people
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );


  if (!sheet) {

    setupSheet();

    sheet =
      ss.getSheetByName(
        SETTINGS_SHEET
      );

  }


  for(
    let i = 0;
    i < 5;
    i++
  ){

    sheet
    .getRange(
      i + 2,
      2
    )
    .setValue(
      people[i] ||
      "Person " + (i + 1)
    );

  }

}


/* ADD EXPENSE */

function addExpense(
  expense,
  people
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      EXPENSE_SHEET
    );


  if (!sheet) {

    setupSheet();

    sheet =
      ss.getSheetByName(
        EXPENSE_SHEET
      );

  }


  sheet.appendRow([

    expense.id,

    expense.date,

    expense.description,

    Number(expense.amount),

    expense.paidBy,

    people[0] || "",

    people[1] || "",

    people[2] || "",

    people[3] || "",

    people[4] || ""

  ]);


  savePeople(people);


  return getAllData();

}


/* DELETE EXPENSE */

function deleteExpense(
  id
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      EXPENSE_SHEET
    );


  if (!sheet) {

    throw new Error(
      "TripExpenses sheet not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2)
    return getAllData();


  const ids =
    sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues();


  for(
    let i = ids.length - 1;
    i >= 0;
    i--
  ){

    if(
      String(ids[i][0])
      === String(id)
    ){

      sheet.deleteRow(i + 2);

      break;

    }

  }


  return getAllData();

}


/* ADD PAYMENT */

function addPayment(
  payment
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      PAYMENT_SHEET
    );


  if (!sheet) {

    setupSheet();

    sheet =
      ss.getSheetByName(
        PAYMENT_SHEET
      );

  }


  sheet.appendRow([

    payment.id,

    payment.date,

    payment.from,

    payment.to,

    Number(payment.amount)

  ]);


  return getAllData();

}


/* DELETE PAYMENT */

function deletePayment(
  id
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      PAYMENT_SHEET
    );


  if (!sheet) {

    throw new Error(
      "Payments sheet not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2)
    return getAllData();


  const ids =
    sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues();


  for(
    let i = ids.length - 1;
    i >= 0;
    i--
  ){

    if(
      String(ids[i][0])
      === String(id)
    ){

      sheet.deleteRow(i + 2);

      break;

    }

  }


  return getAllData();

}


/* DATE FORMAT */

function formatDate(
  value
) {

  if (!value)
    return "";


  if(
    Object.prototype
    .toString
    .call(value)
    === "[object Date]"
  ){

    return Utilities
      .formatDate(
        value,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

  }


  return String(value);

}


/* JSON RESPONSE */

function jsonResponse(
  data
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
