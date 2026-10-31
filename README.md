Features included
- Manage 5 people's names.
- Add expense descriptions.
- Enter expense amounts in ₹.
- Select who paid.
- Automatically split expenses equally among 5 people.
- Calculate each person's balance.
- View complete expense history.
- Delete expenses.
- Save expenses to Google Sheets.
- Retrieve previously saved expenses from Google Sheets.

- Google Sheets connection setup
1. Create a new Google Sheet.
2. Open Extensions → Apps Script.
3. Copy the contents of Code.gs into Apps Script.
4. Click Deploy → New deployment → Web app.
5. Set Execute as: Me.
6. Set access appropriately (Anyone, if you want a publicly accessible endpoint).
7. Deploy and copy your Web App URL.
8. Open index.html and paste the URL into the Google Sheets connection field.
9. Click Connect & load data.
Your webpage will now save and retrieve expenses from Google Sheets.
One question for the next improvement
Would you like the app to also calculate who should pay whom to settle all trip expenses?

For example: Person 1 pays ₹500 to Person 3, and everyone becomes settled.
