//Global Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dumpSheet = ss.getActiveSheet();
var dumpSheetName = dumpSheet.getName();
var dumpSheetValues = dumpSheet.getRange(3,1,dumpSheet.getLastRow(), 5).getValues();
var finalSheet;
var finalSheetValues = [];

function main() {
  //native Google Sheet row number is index + 3
  for(i=0; i<dumpSheetValues.length; i++) {
    var tempComment = dumpSheetValues[i][1];
    var tempAnswer = String()
    if(tempComment == "author") {
      //Highlights the author's response to a comment in light-green
      dumpSheet.getRange(i+5,2).setBackground('#93C47D');
      breakme: for(k=0; k<50; k++) {
        var authorParagraph = String(dumpSheetValues[i+k][1])
        Logger.log(`row: ${i+3} paragraph iteration: ${i+k}`)
        if(authorParagraph.includes('Reply') == false) {
          if(authorParagraph.includes('author') == false) {
            if(authorParagraph.includes('Willy Woo') == false)
            tempAnswer = tempAnswer.concat(`${dumpSheetValues[i+k][1]} \n`);
          }
        } else {
          break breakme;
        }
      }
      storeFinalValues(dumpSheetValues[i-3][1], tempAnswer, i+5);
    }
  }
  displayResults();
}

function storeFinalValues(question, answer, answerLine) {
  return finalSheetValues.push([question, answer, answerLine]);
}

function displayResults() {
  if(ss.getSheetByName(`${dumpSheetName} - Final`)){
    ss.deleteSheet(ss.getSheetByName(`${dumpSheetName} - Final`));
  }
  var finalSheetName = ss.insertSheet(`${dumpSheetName} - Final`).getName();
  finalSheet = ss.getSheetByName(finalSheetName);
  finalSheet.getRange(1, 1, finalSheetValues.length, 3).setValues(finalSheetValues);
  finalSheet.setColumnWidths(1, 1, 380);
  finalSheet.setColumnWidths(2, 1, 1236);
  finalSheet.setRowHeights(1, finalSheetValues.length, 46);
  finalSheet.getRange(1, 1, finalSheetValues.length, 3).setVerticalAlignment('middle');
}

// Creates custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Identify Author (On Current Sheet)', 'main')
    .addToUi();
}
