# Tennis Schedule Manager — Google Apps Script edition

This version uses a private Google spreadsheet for the shared roster and schedule. Apps Script validates every status change: a player can update their own row, while the player whose ID is in `Settings → ManagerPlayerID` can manage every row and the schedule. Only the manager can set or clear `B` (Booked).

## Install

1. Create a project at [script.google.com](https://script.google.com/).
2. Add a script file named `Code.gs` and an HTML file named `Index`. Copy the contents of the matching files in this folder. If you use `appsscript.json`, turn on “Show appsscript.json manifest file” in Project Settings and copy it too.
3. In the Apps Script editor, run `setupProject`. Approve the requested spreadsheet access. The function creates a private data spreadsheet and returns its URL.
4. Open the generated spreadsheet. In the `Players` tab, paste player rows under the header using `PlayerID`, `Name`, and `Email`. Keep PlayerID cells as plain text if IDs can start with zeroes.
5. In `Settings`, set the `ManagerPlayerID` value to the manager’s exact PlayerID. Keep this spreadsheet private; players use the web app, not the data sheet.
6. Run `authorizeEmailSending` in the Apps Script editor and approve the mail permission.
7. Choose **Deploy → New deployment → Web app**. Set **Execute as** to your account and access to **Anyone** (including anonymous visitors). The one-time email link is the app’s sign-in gate; the roster and permissions are checked by the server.

## Use

- Players enter the name and email from the roster to request a one-time sign-in link. Links expire after 15 minutes and work once; sessions last up to six hours.
- The manager can upload or update a roster CSV, resend a player’s link, import/export schedules, change play days, add/remove players, and roll the current Monday–Sunday week forward.
- Player roster CSV headers must be `PlayerID,Name,Email`. The included template has the required header.
- Schedule CSV import accepts `PlayerID` (recommended) or `Player` as the first column, followed by `YYYY-MM-DD` date columns. Old `P` availability values are converted to `A`.
- Apps Script sends sign-in messages from the deploying Google account and is subject to that account’s Google mail limits.

## Security notes

- The spreadsheet should remain private. The web app runs as its owner so players do not need direct access to the spreadsheet.
- The web app must be reachable without Google sign-in so a recipient can open the emailed link. Every app action then checks a short-lived, one-time email-link session on the server.
- Session and link tokens are stored as hashes in Apps Script cache. Cache entries can expire early, in which case the player can request another link.

Apps Script HTML pages call server functions asynchronously through `google.script.run`; see [HTML Service communication](https://developers.google.com/apps-script/guides/html/communication). Deploy the project as a web app using the owner identity and selected access level; see [Apps Script web apps](https://developers.google.com/apps-script/guides/web).
