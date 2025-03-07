class Averages {
    constructor() {
      this.source = SpreadsheetApp.getActiveSpreadsheet();
      this.weeklyPricesSheet = this.source.getSheetByName("Weekly Average Prices");
      this.averagesSheet = this.source.getSheetByName("Averages");
      this.avgProducts = [];
    }
    reCalculate() {
      const lastRow = this.averagesSheet.getLastRow();
      const data = this.weeklyPricesSheet.getDataRange().getValues();
      this.avgProducts = this.averagesSheet.getRange("A2:A" + lastRow).getValues().flat();
      data.slice(1).forEach((cell, index, arr) => {
            const productName = cell[0];
            const rowNumber = this.insertProduct(productName);
            this.weekly(cell, rowNumber);
            this.fourWeekRange(cell, rowNumber);
        });
     
    }
    weekly(cell, rowNumber) {
        //get the latest week in the weekly sheet
        const lastColumn = this.weeklyPricesSheet.getLastColumn();
        const lastColumnLetter = columnToLetter(lastColumn);
        if(cell[lastColumn-1] !== undefined) {
            this.averagesSheet.getRange(rowNumber, 2).setValue(cell[lastColumn-1]*1);
        }
    }
    fourWeekRange(cell, rowNumber){
        //calculate the last 4 weeks in the weekly average sheet
        let slicedCell = cell.slice(1).slice(-4); // Remove product name and only get the last 4 data
        if (slicedCell.length === 0) return 0; // Avoid division by zero
        let sum = slicedCell.reduce((acc, num) => acc + num, 0);
        let avg = sum / slicedCell.length;
        this.averagesSheet.getRange(rowNumber, 3).setValue(avg);
    }
    insertProduct(name) {
        var rowIndex = this.avgProducts.findIndex(value => value === name); // Find exact match
        if (rowIndex !== -1) {
            var actualRow = rowIndex + 2;
            return actualRow;
        } 
        const lastRow = this.averagesSheet.getLastRow();
        var newRow = lastRow + 1;
        this.averagesSheet.getRange(newRow, 1).setValue(name); // Insert new  product
        return newRow;
    }
}