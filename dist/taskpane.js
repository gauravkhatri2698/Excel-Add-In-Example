/******/ (function() { // webpackBootstrap
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    document.getElementById("create-table").onclick = createTable;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function createTable() {
  Excel.run(function (context) {
    // TODO1: Queue table creation logic here.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true
    /*hasHeaders*/
    );
    expensesTable.name = "ExpensesTable"; // TODO2: Queue commands to populate the table with data.

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
    expensesTable.rows.add(null
    /*add at the end*/
    , [["1/1/2017", "The Phone Company", "Communications", "120"], ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"], ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"], ["1/10/2017", "Coho Vineyard", "Restaurant", "33"], ["1/11/2017", "Bellows College", "Education", "350.1"], ["1/15/2017", "Trey Research", "Other", "135"], ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]]); // TODO3: Queue commands to format the table.

    expensesTable.columns.getItemA(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error); // eslint-disable-next-line no-undef

    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
/******/ })()
;
//# sourceMappingURL=taskpane.js.map