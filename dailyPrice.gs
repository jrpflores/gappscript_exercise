class DailyPriceService {
    constructor() {
        this.source = SpreadsheetApp.getActiveSpreadsheet();
        this.dailyPricesSheet = this.source.getSheetByName("Daily Prices") || this.createDailyPricesSheet();
    }
    createDailyPricesSheet() {
        let sheet = this.source.insertSheet("Daily Prices");
        sheet.getRange(1, 1, 1, 1).setValues([["Product"]]);
        return sheet;
    }
    getDailyPriceHeaderDate() {
        let lastColumn = this.dailyPricesSheet.getLastColumn();
        let lastColumnValue = this.dailyPricesSheet.getRange(1, lastColumn).getValue();
        let defaultDate = "2025-03-01";
        let dateValue = lastColumnValue === "Product" || !isValidDate(lastColumnValue) 
            ? defaultDate 
            : getNextDay(lastColumnValue);

        this.dailyPricesSheet.insertColumnAfter(lastColumn)
            .getRange(1, lastColumn + 1)
            .setValue(dateValue)
            .setNumberFormat("yyyy-mm-dd");

        return dateValue;
    }
    getProductPosition(productName) {
        const dataRange = this.dailyPricesSheet.getDataRange();
        const lastRow = dataRange.getLastRow();
        const data = this.dailyPricesSheet.getRange("A2:A" + lastRow).getValues().flat();
        const index = data.findIndex( value => value === productName);
        return index !== -1? index + 1 : -1; //if -1 it means product is not yet in the list
    }
    insertNewDailyPriceData(name, price, date) {
        const dataRange = this.dailyPricesSheet.getDataRange();
        const lastRow = dataRange.getLastRow();
        const lastCol = dataRange.getLastColumn();
        const lastColLetter = columnToLetter(lastCol);
        const headers = this.dailyPricesSheet
                              .getRange(`A1:${lastColLetter}1`)
                              .getValues()
                              .flat()
                              .map(header => {
                                // Check if the value is a date
                                if (header instanceof Date) {
                                  return Utilities.formatDate(header, Session.getScriptTimeZone(), "yyyy-MM-dd");
                                }
                                return header;
                              });
        const headerColumnIndex = headers.findIndex( value => value === date)+1;
        const data = this.dailyPricesSheet.getRange("A2:A" + lastRow).getValues().flat();
        var rowIndex = data.findIndex(value => value === name); // Find exact match
        if (rowIndex !== -1) {
            var actualRow = rowIndex + 2; //2 cause range starts at a2
            this.dailyPricesSheet.getRange(actualRow, headerColumnIndex).setValue(price*1);
            return true;
        } 
        var newRow = lastRow + 1;
        this.dailyPricesSheet.getRange(newRow, 1).setValue(name); // Insert new  product
        this.dailyPricesSheet.getRange(newRow, headerColumnIndex).setValue(price*1); // Insert new insert price
        return true;
    }
}