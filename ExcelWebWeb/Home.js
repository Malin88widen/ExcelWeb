'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            //$('#create-table').click(createTable);

            $('#set-first-statement').click(setFirstStatement);
            $('#set-articlenr').click(setArticelnumber); 
            $('#set-middle-statement').click(setMiddleStatement);
            $('#set-articel').click(setArticel);
            $('#set-last-statement').click(setLastStatement); 
           
            //S - M
            $('#set-Sstatement').click(setSStatement); 
            $('#set-Sproduct').click(setSProduct);
            $('#set-Mstatement').click(setMStatement);
            $('#set-Mproduct').click(setMProduct);
            // S1 - M1
            $('#set-S1statement').click(setS1Statement);
            $('#set-S1product').click(setS1Product);
            $('#set-M1statement').click(setM1Statement);
            $('#set-M1product').click(setM1Product);
            // S2 - M2
            $('#set-S2statement').click(setS2Statement);
            $('#set-S2product').click(setS2Product);
            $('#set-M2statement').click(setM2Statement);
            $('#set-M2product').click(setM2Product);
            // S3 - M3
            $('#set-S3statement').click(setS3Statement);
            $('#set-S3product').click(setS3Product);
            $('#set-M3statement').click(setM3Statement);
            $('#set-M3product').click(setM3Product);
            // S4 - M4
            $('#set-S4statement').click(setS4Statement);
            $('#set-S4product').click(setS4Product);
            $('#set-M4statement').click(setM4Statement);
            $('#set-M4product').click(setM4Product);

            $('#set-last-statement_2').click(setLastStatement); 

        });

    });

 


    function setFirstStatement() {
        Excel.run(function (context) {
            //Excel.createWorkbook();

            //The code if you want to choose where the statement is going to be in the excel
            //var actSh = context.workbook.worksheets.getActiveWorksheet(); 
            //var rng = actSh.getRange('A2');

            var rng = context.workbook.getSelectedRange();
            rng.values = "SET @chvXMLDoc = @chvXMLDoc + '<a I=\"";
           

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setArticelnumber() {
        Excel.run(function (context) {
            //Excel.createWorkbook();

            var range = context.workbook.getSelectedRange(); //The code to let the user choose where to put the statement in the excel
            range.values = "Artikelnnummer:";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setMiddleStatement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\"D =\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setArticel() {
        Excel.run(function (context) {
            //Excel.createWorkbook();

            var rng = context.workbook.getSelectedRange();
            rng.values = "Artikel:";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // S-M 

    function setSStatement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" S=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setSProduct() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "S-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setMStatement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" M=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setMProduct() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "M-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    //S1- M1
    
    function setS1Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" S1=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setS1Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "S1-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM1Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" M1=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM1Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "M1-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // S2 - M2

    function setS2Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" S2=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setS2Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "S2-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM2Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" M2=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM2Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "M2-Product";
      

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // S3 - M3

    function setS3Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" S3=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setS3Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "S3-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM3Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" M3=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM3Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "M3-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // S4 - M4

    function setS4Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" S4=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setS4Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "S4-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM4Statement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" M4=\"";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setM4Product() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "M4-Product";

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // Slutet "\" />'"

    function setLastStatement() {
        Excel.run(function (context) {

            var rng = context.workbook.getSelectedRange();
            rng.values = "\" />'";
         

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    //TEST

    //function createTable() {
    //    Excel.run(function (context) {

    //        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    //        var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    //        expensesTable.name = "ExpensesTable";

    //        expensesTable.getHeaderRowRange().values =
    //            [["Date", "Merchant", "Category", "Amount"]];

    //        expensesTable.rows.add(null /*add at the end*/, [
    //            ["1/1/2017", "The Phone Company", "Communications", "120"],
    //            ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
    //            ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
    //            ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
    //            ["1/11/2017", "Bellows College", "Education", "350.1"],
    //            ["1/15/2017", "Trey Research", "Other", "135"],
    //            ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    //        ]);

    //        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    //        expensesTable.getRange().format.autofitColumns();
    //        expensesTable.getRange().format.autofitRows();

            

    //        return context.sync();
    //    })
    //        .catch(function (error) {
    //            console.log("Error: " + error);
    //            if (error instanceof OfficeExtension.Error) {
    //                console.log("Debug info: " + JSON.stringify(error.debugInfo));
    //            }
    //        });
    //}


})();

