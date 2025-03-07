class ChartProcessor {
    constructor() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      this.chartsSheet = ss.getSheetByName("Charts") || ss.insertSheet("Charts");
      this.dailyPricesSheet = ss.getSheetByName("Daily Prices");
      this.weeklyPricesSheet = ss.getSheetByName("Weekly Average Prices");
    }
    clear() {
      const charts = this.chartsSheet.getCharts();
      charts.forEach(chart => this.chartsSheet.removeChart(chart))
    }
    render() {
        this.clear();
        this.weekly();
        this.fourweekRange();

    }
    weekly() {
        const lastColumn = this.dailyPricesSheet.getLastColumn();
        const lastColLetter = columnToLetter(lastColumn);
        const lastRow = this.dailyPricesSheet.getLastRow();
        const data = this.dailyPricesSheet.getRange(`B1:${lastColLetter}${lastRow}`);

        // Get product names (excluding header row)
          const productNames = this.dailyPricesSheet.getRange(`A2:A${lastRow}`).getValues().flat();

          // Build series config dynamically
          let seriesConfig = {};
          productNames.forEach((product, index) => {
              seriesConfig[index] = { labelInLegend: product };
          });
        var chart = this.chartsSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(data)
            .setOption("title", "Daily Price Trend of Products")
            .setOption("hAxis", { title: "Date" })
            .setOption("vAxis", { title: "Price" })
            .setOption("series", seriesConfig) 
            .setOption("legend", { position: "right" }) // Products in the legend
            .setOption("useFirstColumnAsDomain", true)
            .setOption("useFirstRowAsHeaders", true)
            .setTransposeRowsAndColumns(true) // Ensure products are series
            .setPosition(2, 2, 0, 0)
            .build();

          this.chartsSheet.insertChart(chart);
    }
    fourweekRange() {
        const lastColumn = this.weeklyPricesSheet.getLastColumn();
        const lastColLetter = columnToLetter(lastColumn);
        const lastRow = this.weeklyPricesSheet.getLastRow();
        const data = this.weeklyPricesSheet.getRange(`B1:${lastColLetter}${lastRow}`);

        // Get product names (excluding header row)
          const productNames = this.weeklyPricesSheet.getRange(`A2:A${lastRow}`).getValues().flat();

          // Build series config dynamically
          let seriesConfig = {};
          productNames.forEach((product, index) => {
              seriesConfig[index] = { labelInLegend: product };
          });
        var chart = this.chartsSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(data)
            .setOption("title", "Weekly Price Trend of Products")
            .setOption("hAxis", { title: "Weeks"})
            .setOption("vAxis", { title: "Price" })
            .setOption("series", seriesConfig) 
            .setOption("legend", { position: "right" }) // Products in the legend
            .setOption("useFirstColumnAsDomain", true)
            .setOption("useFirstRowAsHeaders", true)
            .setTransposeRowsAndColumns(true) // Ensure products are series
            .setPosition(20, 2, 0, 0)
            .build();

          this.chartsSheet.insertChart(chart);
    }
}


