class WeeklyPrices {
    constructor() {
      this.source = SpreadsheetApp.getActiveSpreadsheet();
      this.dailyPricesSheet = this.source.getSheetByName("Daily Prices") || this.createDailyPricesSheet();
      this.weeklyPricesSheet = this.source.getSheetByName("Weekly Average Prices") || this.source.insertSheet("Weekly Average Prices");
    }
    create() {
      //clear data
      this.weeklyPricesSheet.clear();

      const data = this.dailyPricesSheet.getDataRange()
                                        .getValues();
      const lastCol = this.dailyPricesSheet.getLastColumn();
      let finalData = [];
      let headers = [data[0][0]];
      let lastEndDate = null;
      data[0].slice(1).forEach((cell, col, arr) => {
        if (col % 7 === 0 || col === data[0].length - 1) { // Start of a new week or last column
          let startDate = new Date(cell);
          let endIndex = Math.min(col + 6, arr.length - 1);
          let endDate = new Date(arr[endIndex]);

          let formattedStart = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
          let formattedEnd = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM/dd/yyyy");

          let rangeLabel = `${formattedStart} - ${formattedEnd}`;

          if (lastEndDate !== formattedEnd) { // Avoid duplicate date ranges
            headers.push(rangeLabel);
            lastEndDate = formattedEnd;
          }
        }
      });
      finalData.push(headers);
      //get data
      data.slice(1).forEach(row => {
        let rowData = [row[0]]; // Product name
        row.slice(1).reduce((acc, value, index) => {
          acc.push(value === "" || isNaN(value) ? 0 : value);
          if (index % 7 === 6 || index === row.length - 2) { // Every 7 days or last column
            let avg = acc.reduce((sum, val) => sum + val, 0) / acc.length;
            rowData.push(avg.toFixed(2));
            acc.length = 0; // Reset accumulator for next week
          }
          return acc;
        }, []);
        finalData.push(rowData);
      });
      this.weeklyPricesSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);

    }
}
