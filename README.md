# App Script for Getting Unique Values and Counting Their Number in Google Sheets

This App Script for Google Sheets allows you to get unique values in a specified range of cells and count their number.

## Usage

1. Open the Google Sheet in which you want to use this script.

2. From the menu, select "Tools" -> "Script Editor".

3. Paste the provided code into the script editor.

4. Save the script and give it a name.

5. Close the script editor.

6. Now you have the `getUniqueValuesAndCount()` function.

7. Run this function and it will create a new range with unique values and their number.

```javascript
function getUniqueValuesAndCount() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   var range = sheet.getRange("A1:D10"); // Replace "A2:D10" with your range of cells

   var values = range.getValues();
   var uniqueValues = {};

   for (var i = 0; i < values.length; i++) {
     for (var j = 0; j < values[i].length; j++) {
       var cellValue = values[i][j];
       if (cellValue !== "") {
         if (uniqueValues[cellValue]) {
           uniqueValues[cellValue]++;
         } else {
           uniqueValues[cellValue] = 1;
         }
       }
     }
   }

   var outputRange = sheet.getRange("F2:G" + (Object.keys(uniqueValues).length + 1));
   var outputValues = [];

   for (var key in uniqueValues) {
     outputValues.push([key, uniqueValues[key]]);
   }

   outputRange.setValues(outputValues);
}
```
