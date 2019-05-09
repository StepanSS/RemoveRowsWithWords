
var sheetSett = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
var sheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet1");
var sheetRes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Result");

//=================Function - Remove All Rows with Selected Words
function main(){
  sheetRes.clearContents();
  getData();
  
}
//=================Function - Remove All EXCEPT Rows with Selected Words
function main2(){
  sheetRes.clearContents();
  getData2();
}

function getWords() {
  var last_row = sheetSett.getLastRow();
  var wordsArray = sheetSett.getRange(2, 1, last_row-1).getValues();
  
  var wordsToRegEx;
  // --------------------------------------create REGEX Pattern From All Words
  for (var i = 0; i<wordsArray.length; i++){
    //Logger.log(wordsArray[i][0]);
    //wordsArray[i][0] = "(^"+wordsArray[i][0]+"\\s)|(\\s"+wordsArray[i][0]+"\\s)|(\\s"+wordsArray[i][0]+"\\b)";
  }
  wordsToRegEx = wordsArray.join("|");//concat all words in one string
  Logger.log(wordsToRegEx);
  return wordsToRegEx;
}
//===Function for main() case - (remove all rows with selected words)
function getData(){
  var last_row = sheetData.getLastRow();
  var last_column = sheetData.getLastColumn();
  var wordsArray = sheetData.getRange(1, 1, last_row, last_column).getValues();
  Logger.log(wordsArray.length);
  
  var counter = 0;
  var newData = [];
  var wordsToRem = getWords();
  
  //var regExPrice = "(^"+testRegEx+"\\s)|(\\s"+testRegEx+"\\s)";
  var myRegexp = new RegExp(wordsToRem,'i');
  Logger.log(myRegexp);
  
  for (var j =0; j<wordsArray.length; j++){
      
      var myArray;      
      var trigger =false;
      
      for(var i=0; i< wordsArray[j].length; i++){
          //if
//        if(i!=1 || i!=3 || i!=4 || i!=5 || i!=6 || i!=8 || i!=11 || i!=13 || i!=14 || i!=15 ){
          //var result = wordsArray[j][i];
          myArray = myRegexp.exec(wordsArray[j][i]);
          if (myArray != null){
            //Logger.log(myArray);
            trigger = true;
            counter++;
            break;
          }
//        }
      }
      if(trigger){
            //Logger.log("Found in Row:"+j);
      }else{
        // Create NEW Array
            //Logger.log("Not found in Row:"+j);
            //sheetRes.appendRow(wordsArray[j]);
            newData.push(wordsArray[j]);
      }
  } 
  //Logger.log(newData);
  sheetRes.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  Browser.msgBox("Found rows with selected words: "+counter);
  
}

//===Function for main2() case - (remove all rows except rows which has selected words)
function getData2(){
  var last_row = sheetData.getLastRow();
  var last_column = sheetData.getLastColumn();
  var wordsArray = sheetData.getRange(1, 1, last_row, last_column).getValues();
  Logger.log(wordsArray.length);
  
  var counter = 0;
  var newData = [];
  var wordsToRem = getWords();
  
  //var regExPrice = "(^"+testRegEx+"\\s)|(\\s"+testRegEx+"\\s)";
  var myRegexp = new RegExp(wordsToRem,'i');
  Logger.log(myRegexp);
  
  for (var j =0; j<wordsArray.length; j++){
      
      var myArray;      
      var trigger =false;
      
      for(var i=0; i< wordsArray[j].length; i++){
          //if
//        if(i!=1 || i!=3 || i!=4 || i!=5 || i!=6 || i!=8 || i!=11 || i!=13 || i!=14 || i!=15 ){
          //var result = wordsArray[j][i];
          myArray = myRegexp.exec(wordsArray[j][i]);
          if (myArray != null){
            //Logger.log(myArray);
            trigger = true;
            
          }
//        }
      }
      if(trigger){
            //Logger.log("Found in Row:"+j);
            counter++;
            newData.push(wordsArray[j]);
//            Logger.log(wordsArray[j]); 
//            Logger.log("Counter: "+counter);
      }else{
        // Create NEW Array
            //Logger.log("Not found in Row:"+j);
            
      }
  } 
  //Logger.log(newData);
  sheetRes.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  Browser.msgBox("Found rows with selected words: "+counter);
  
}