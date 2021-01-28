function onEdit(e) {
  var oldValue;
  var edit;
  var newValue;
  var oldValueAsArray;
  var editAsArray;
  var newValueAsArray;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var activeCell = sheet.getCurrentCell();
  var col = activeCell.getColumn();
  var row = activeCell.getRow();

//customize sheet and column where multiple selection drop downs should be
  if (sheetName == "sheet1") && col == 3) {
    edit = e.range;
    newValue = e.value;
    Logger.log(edit);

    oldValue = e.oldValue;
    Logger.log(oldValue);

	  if(!e.value) {
      
      activeCell.setValue(e.range.getValue() != null ? edit.getValue() : "");

    } else if (!e.oldValue) {

      activeCell.setValue(newValue);
      
	 } else {

      oldValueAsArray = oldValue.split(',').map(function(name) {
        return name.trim();
      }); 

      Logger.log(oldValueAsArray);

      editAsArray = newValue.split(',').map(function(name) {
        return name.trim();
      });
        
      Logger.log(editAsArray);

      editAsArray = valueCleanup(editAsArray);
      Logger.log(editAsArray);

      //check if any of the values in the edited string, now an array, were part of the previous value, now an array
      var check = editAsArray.every((name) => {
        return oldValueAsArray.indexOf(name) != -1;
      })

      //if not one element in the edit array is found in the previous cell value
      if (check != true) {
        
        var combined = oldValueAsArray.concat(editAsArray);

        var combset = new Set(combined);

        newValueAsArray = valueCleanup(Array.from(combset));

        newValue = newValueAsArray.join(', ');

        activeCell.setValue(newValue);
          
      } else {
        // if one or more names is found in the previous value the code executes this
        newValue = editAsArray.join(', ');
        activeCell.setValue(newValue);
        }
      }

    updateDataVal(activeCell.getA1Notation());
    Logger.log(activeCell);
  }
  
  
function updateDataVal (str) {
  var cell = SpreadsheetApp.getActiveSheet().getRange(str);
  var value = cell.getDisplayValue();
  var col = cell.getColumn();
  var newArray = addToNewList(value,col);


  var list = Array.from(newArray);
  list.unshift(value);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(true).build();
  cell.setDataValidation(rule);
}

//update selection list with choices
function addToNewList (val,col) {
  let value = new String (val);
  let column = Number(col);
  
  var referenceArray;
  
  if (column == 3) {
  //base list of selections added as starting reference
  var listRange = SpreadsheetApp.getActiveSpreadsheet().getRange("Selection List!$A$1:$A$15");
  var selections = listRange.getValues().join(',').split(',');
  valueCleanup(selections);
  referenceArray = selections;
  }
  
  var setA = new Set(referenceArray);
  var setB; 
  var diff;
    
  if (value.indexOf(',') != -1) {

    var valueArray = value.split(',').map(function(name) {
      return name.trim();
    });

    var newArray = valueArray;
    Logger.log(newArray);
    
    setB = new Set(newArray);
   
   if (setB.has("")) {
      setB.delete("");
    }
    
    diff = difference(setA, setB);

    Logger.log(Array.from(setB));
    Logger.log(Array.from(diff));

    return Array.from(diff);

    } else if (value != null || value != "") {

      var setC = new Set (setA);
      setC.delete(value.valueOf());

      return Array.from(setC);

    } else {

      return Array.from(setA);
    }
}

//remove extra spaces
function valueCleanup (arr) {
  var cleanArray = [];
  var output = arr
    .map(val => val.trim())
   .filter(val => val !== '')
  Logger.log(output)

  cleanArray = output.sort();
  Logger.log(cleanArray);

  return cleanArray;
}

