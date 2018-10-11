/**
 *    This program is free software: you can redistribute it and/or modify
 *    it under the terms of the GNU General Public License as published by
 *    the Free Software Foundation, either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   This program is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU General Public License for more details.
 *
 *   You should have received a copy of the GNU General Public License
 *   along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

/**
 * Creates and builds a card that has the drop-downs of people and the button to generate email.
 * @param {event object} containing information about the open Gmail message ID.
 * @return {CardBuilder} The card with used for user interation.
 */
function buildAddOn(e) { 
  // Get the cell values
  var peopleObjects = getIntroCellsFromSpreadsheet(getSpreadsheet());
  
  if (!peopleObjects){
    // return an error card
  }
  
  //Create a drop down with all the people in it
  var dropDownPerson1 = CardService.newSelectionInput()    
  .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Select First Person")
    .setFieldName("person_1");
  for (var i = 0; i<peopleObjects.length; i++){
    dropDownPerson1.addItem(peopleObjects[i]['name'], i, false);
  };

  //Create another drop down with all the people in it
  var dropDownPerson2 = CardService.newSelectionInput()    
  .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Select Second Person")
    .setFieldName("person_2");
  for (var i = 0; i<peopleObjects.length; i++){
     dropDownPerson2.addItem(peopleObjects[i]['name'], i, false);
  };

  // Create a button that will compose the email
  var action = CardService.newAction().setFunctionName('composeEmailCallback').setParameters({key1: "value1"});
  var button =CardService.newTextButton()
    .setText('Compose Email')
    .setComposeAction(action, CardService.ComposedEmailType.STANDALONE_DRAFT);
  
  // Create a card with a single card section and the dropdowns and button.
  var introCard = CardService.newCardBuilder()
  .setHeader(CardService.newCardHeader()
             .setTitle('Generate an intro between these 2 people'))
  .addSection(CardService.newCardSection()
              .addWidget(dropDownPerson1)
              .addWidget(dropDownPerson2)
              .addWidget(button))
  .build();   
  
  // Return the card we built
  return [introCard];
}

/**
 * Returns The generated email draft builder.
 * @param {object} e The event info.
 * @return {ComposeActionResponseBuilder} The builder with the email configured
 */
function composeEmailCallback(e) {
  // Get the spreadsheet
  var spreadsheet = getSpreadsheet();
  
  // Get the people objects
  var peopleObjects = getIntroCellsFromSpreadsheet(spreadsheet);
  
  // Get the first string and email
  var selection1 = parseInt(e.formInputs.person_1);
  var personData1 = peopleObjects[selection1];
  var email1 = personData1['email'];
  
  // Get the second string and email
  var selection2 = parseInt(e.formInputs.person_2);
  var personData2 = peopleObjects[selection2];
  var email2 = personData2['email'];
  
  // Get the template
  var templateSheet = spreadsheet.getSheets()[1];
  var template = templateSheet.getRange('A1').getValue();
  //var template = 'Hi ${"Name1"} ,\n\nCopied here is ${"Name2"} who is ${"Text2"}.\n\n ${"Name2"}, ${"Name1"} is ${"Text1"}.\n\n I hope you two can connect.\n\nCheers,\nSean Byrnes\nCEO/Outlier\nhttp://outlier.ai'; 
  
  // Get the email string
  var emailString = fillInTemplateFromObject(template, personData1, personData2);
   
  // Create a draft of the email
  var draft = GmailApp.createDraft(email1 +',' + email2 , 'Intro', emailString); 
  return CardService.newComposeActionResponseBuilder()
      .setGmailDraft(draft)
      .build();
}

/**
 * Replaces markers in a template string with values defined in a JavaScript data object.
 * @param {string} template Contains markers, for instance ${"Column name"}
 * @param {object} data values to that will replace markers with the 1 suffix.
 *   For instance data.columnName will replace marker ${"Column name"}
 * @param {object} data values to that will replace markers with the 2 suffix.
 *   For instance data.columnName will replace marker ${"Column name"}
 * @return {string} A string without markers. If no data is found to replace a marker,
 *   it is simply removed.
 */
function fillInTemplateFromObject(template, data1, data2) {
  var email = template;
  
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var normalizedVariable = normalizeHeader(templateVars[i]);
    var dataNumber = normalizedVariable.slice(-1);
    var variable = normalizedVariable.slice(0, -1);
    // If no number is added assume to use the first one.
    var variableData;
    if (dataNumber === '2'){
      variableData = data2[variable];
    }
    else{
      variableData = data1[variable];
    }
    email = email.replace(templateVars[i], variableData || '');
  }

  return email;
}

function getSpreadsheet(){
  var allFilesInFolder,fileNameToGet,fldr;//Declare all variable at once
  
  // Get the sheets file and open it
  fileNameToGet = 'Intros';
  allFilesInFolder = DriveApp.getFilesByName(fileNameToGet);
  if (allFilesInFolder.hasNext() === false) {
    //If no file is found, the user gave a non-existent file name
    return false;
  };
  
  // Get the first one
  var file = allFilesInFolder.next();
  // Open the intros spreadsheet
  return SpreadsheetApp.open(file);
}

function getIntroCellsFromSpreadsheet(sheet){
  
  var sheet = sheet.getSheets()[0];
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  //Get the cells that hold the names
  var range = sheet.getRange(2, 1, lastRow-1, lastColumn);
  
  // Create one JavaScript object per row of data.
  return getRowsData(sheet, range);
}

/**
 * Normalizes a string, by removing all alphanumeric characters and using mixed case
 * to separate words. The output will always start with a lower case letter.
 * This function is designed to produce JavaScript object property names.
 * @param {string} header The header to normalize.
 * @return {string} The normalized header.
 * @example "First Name" -> "firstName"
 * @example "Market Cap (millions) -> "marketCapMillions
 * @example "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 */
function normalizeHeader(header) {
  var key = '';
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == ' ' && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

/**
 * Returns true if the character char is alphabetical, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a number.
 */
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

/**
 * Returns true if the cell where cellData was read from is empty.
 * @param {string} cellData Cell data
 * @return {boolean} True if the cell is empty.
 */
function isCellEmpty(cellData) {
  return typeof(cellData) == 'string' && cellData == '';
}

/**
 * Returns true if the character char is a digit, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a digit.
 */
function isDigit(char) {
  return char >= '0' && char <= '9';
}

/**
 * Iterates row by row in the input range and returns an array of objects.
 * Each object contains all the data for a given row, indexed by its normalized column name.
 * @param {Sheet} sheet The sheet object that contains the data to be processed
 * @param {Range} range The exact range of cells where the data is stored
 * @param {number} columnHeadersRowIndex Specifies the row number where the column names are stored.
 *   This argument is optional and it defaults to the row immediately above range;
 * @return {object[]} An array of objects.
 */
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

/**
 * For every row of data in data, generates an object that contains the data. Names of
 * object fields are defined in keys.
 * @param {object} data JavaScript 2d array
 * @param {object} keys Array of Strings that define the property names for the objects to create
 * @return {object[]} A list of objects.
 */
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

/**
 * Returns an array of normalized Strings.
 * @param {string[]} headers Array of strings to normalize
 * @return {string[]} An array of normalized strings.
 */
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}
