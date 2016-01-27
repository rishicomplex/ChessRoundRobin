## ChessRoundRobin
A Google Apps Script that extends a spreadsheet to generate pairings and maintain scores for a round robin chess tournament.

### How to use
Create a new spreadsheet on Google Sheets. Click on the _'Script editor'_ option on the _'Tools'_ menu. Copy in the code in `tournament.gs`, and reload the spreadsheet.

List the names of the players of your tournament in the first column (cells A1, A2 etc). Ensure that the number of players is even. If there are an odd number of players, add a dummy player. All games with the dummy player can be considered byes. Now, select the _'Generate tournament'_ option in the _'Tournaments'_ menu.

The pairings for all rounds will be generated. Once a game is completed, enter the result in the cell between the players' names. The score table will update automatically.
