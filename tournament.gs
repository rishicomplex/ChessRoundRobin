
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate tournament', functionName: 'generateTournament_'}
  ];
  spreadsheet.addMenu('Tournament', menuItems);
  //generateTournament_();
}

var numberOfPlayers;
var players = [];
var lastRow;

function generateTournament_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet1 = spreadsheet.getSheetByName("Sheet1");
  sheet1.activate();
  
  var selectedRow = Browser.inputBox('Generate tournament',
                                    'Please enter the number of players',
                                     Browser.Buttons.OK_CANCEL);
  if (selectedRow == 'cancel') {
    return;
  }
  var rowNumber = Number(selectedRow);
  
  if (isNaN(rowNumber) ||
      rowNumber > sheet1.getLastRow()) {
    Browser.msgBox('Error',
                   Utilities.formatString('Row "%s" is not valid.', selectedRow),
                   Browser.Buttons.OK);
    return;
  }
  
  numberOfPlayers = rowNumber;
  
  var row = sheet1.getRange(1, 1, rowNumber, 1);
  var rowValues = row.getValues();
  
  for(var i = 0; i < rowValues.length; i++) {
    players.push(rowValues[i][0]); 
  }
  
  addScoreTable(sheet1);
  
  lastRow = players.length + 6;
  
  addInstructions(sheet1);
  
  addRounds(sheet1);
  
}

function addInstructions(sheet) {
  newRow = [["After your game, enter your result in the cell between player names. The score table will automatically update."], ["Result can be one of:   1-0,   0.5-0.5,   0-1"]];
  sheet.getRange(lastRow, 1, 2, 1).setValues(newRow);
  
  sheet.getRange(lastRow, 1, 1, 8).merge();
  sheet.getRange(lastRow + 1, 1, 1, 8).merge();
  
  lastRow += 4;
  
}

function addRound(no, player1, rotPlayers, sheet) {
  Logger.log("Add round called" + no);
  var newRows = [["Round " + no]];
  sheet.getRange(lastRow, 1, 1, 1).setValues(newRows);
  sheet.getRange(lastRow, 1, 1, 1).setFontWeight('bold');
  lastRow++;
  newRows = [];
  for(i = 0; i < (rotPlayers.length + 1) / 2; i++) {
    newRows.push([]); 
  }
  var p1_color, p2_color;
  if(no % 2 == 1) {
    p1_color = 0;
    p2_color = 2
  }
  else {
    p1_color = 2;
    p2_color = 0;
  }
  newRows[0][p1_color] = player1;
  newRows[0][1] = '-';
  no--;
  for(var i = 1; i < ((rotPlayers.length + 1) / 2); i++) {
    newRows[i][0] = rotPlayers[no];
    no = nextIndex(no, rotPlayers.length);
  }
  for(var i = ((rotPlayers.length + 1) / 2) - 1; i >= 0; i--) {
    newRows[i][1] = '-';
    if(i == 0) 
      newRows[i][p2_color] = rotPlayers[no];
    else 
      newRows[i][2] = rotPlayers[no];
    no = nextIndex(no, rotPlayers.length);
  }
  
  for(var i = 0; i <= (rotPlayers.length / 2); i++) {
    addFormula(newRows[i][0], newRows[i][2], lastRow + i, sheet);
  }
  
  sheet.getRange(lastRow, 1, (rotPlayers.length + 1) / 2, 3).setValues(newRows);
  sheet.getRange(lastRow, 2, (rotPlayers.length + 1) / 2, 1).setHorizontalAlignment('center');
  lastRow += ((rotPlayers.length + 1)/ 2) + 1;
  
                  
}

function nextIndex(ind, size) {
  if(ind < size - 1) return ind + 1;
  else return 0;
}

function addRounds(sheet) {
  
  var newRow = [["White", "", "Black"]];
  sheet.getRange(lastRow, 1, 1, 3).setValues(newRow);
  sheet.getRange(lastRow, 1, 1, 3).setFontWeight('bold');
  
  lastRow += 2;
  
  var player1 = players[players.length - 1];
  var rotPlayers = players.slice(0, players.length - 1);
  
  for(var i = 0; i < players.length - 1; i++) {
    //Logger.log("Loop iter");
    addRound(i + 1, player1, rotPlayers, sheet);
  }
 
  
  
}

function getIndex(player) {
  for(var i = 0; i < players.length; i++) {
    if(players[i] == player) return i;
  }
}


function addFormula(player1, player2, row, sheet) {
  var opponent;
  
  var cell = sheet.getRange(row, 2, 1, 1).getCell(1, 1).getA1Notation();
  Logger.log(cell);
  var formula1 = '=LEFTSCORE(' + cell + ')';
  var formula2 = '=RIGHTSCORE(' + cell + ')';
  var ind1 = getIndex(player1);
  var ind2 = getIndex(player2);
  var row1 = 1 + ind1 + 1;
  var column1 = 1 + ind2 + 3;
  var row2 = 1 + ind2 + 1;
  var column2 = 1 + ind1 + 3;
  sheet.getRange(row1, column1, 1, 1).setFormula(formula1);
  sheet.getRange(row2, column2, 1, 1).setFormula(formula2);
}

function addScoreTable(sheet) {
  
  var newRows = [];
  newRows.push(['']);
  for(var i = 0; i < players.length; i++) {
    newRows[0].push(players[i]);
  }
  for(var i = 0; i < players.length; i++) {
    newRows.push([players[i]]);
    for(var j = 0; j < players.length; j++) {
      if(j == i) 
        newRows[i + 1].push('-');
      else
        newRows[i + 1].push('');
    }
  }
  for(var i = 0; i < players.length; i++) {
    newRows[i + 1][players.length + 1] = "0";
  }
  newRows[0][players.length + 1] = "Total";
  sheet.getRange(1, 3, players.length + 1, players.length + 2).setValues(newRows);
  sheet.getRange(1, 4 + players.length, 1, 1).setFontWeight('bold');
  var formula;
  for(var i = 0; i < players.length; i++) {
    formula = '=CHESSSUM(';
    formula += sheet.getRange(i + 2, 4, 1, players.length).getA1Notation();
    formula += ')';
    sheet.getRange(i + 2, 4 + players.length, 1, 1).setFormula(formula);
  }
  
}

function leftScore(score) {
  if(score.indexOf('.') != -1) {
    return 0.5;
  }
  return score[0];
}

function rightScore(score) {
  if(score.indexOf('.') != -1) {
    return 0.5;
  }
  return score[score.length - 1];
}

function chessSum(range) {
  var sum = 0;
  for(var i = 0; i < range[0].length; i++) {
    if(range[0][i] != '-')
      sum += Number(range[0][i]);
  }
  return sum;
}
