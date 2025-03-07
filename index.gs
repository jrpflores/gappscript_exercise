//Create a Google Script function that the user can run on this spreadsheet to update the "Averages" sheet with new data from "New 24 Hour Data" sheet.
//TO DO:
//. 1. insert data in daily price ✅
//. 2. insert weekly with average price ✅
//  3. update Averages in weekly and 4week column ✅
//. 4. recreate chart for weekly and 4week range ✅

function main() {
  const response = exercise()
  //run chart if response is successful
  if(response) {
    const chartProcessorService = new ChartProcessor();
    chartProcessorService.render();
  }
  
}
function exercise() {
    var source = SpreadsheetApp.getActiveSpreadsheet();
    //getting the sheets to use
    var twentyFourHourPriceSheet = source.getSheetByName("New 24 Hour Data");
    var averagesSheet = source.getSheetByName("Averages");
    if(!twentyFourHourPriceSheet || !averagesSheet) {
       Logger.log("Cant find New 24 Hour Data or Averages sheet");
      return false;
    }
    //get data from new 24 hour data
    var data = twentyFourHourPriceSheet.getDataRange().getValues();
    //halt operation when there is no data in 24 hour data sheet
    // < 2 to exclude header
    if(data.length < 2) { 
      Logger.log("Nothing to process. No data");
      return false;
    }
    const dailyPriceService = new DailyPriceService();
    const weeklyPriceService = new WeeklyPrices();
    const dailyPricesDate = dailyPriceService.getDailyPriceHeaderDate()
    data.slice(1).forEach(([productName, productPrice]) => {
        if (productName.trim() != "") {
            dailyPriceService.insertNewDailyPriceData(productName, productPrice, dailyPricesDate);
        }
    });
    weeklyPriceService.create();
    //update Averages sheet - USE the latest weekly data in weekly
    const averageService = new Averages();
    averageService.reCalculate();
    return true;
}
