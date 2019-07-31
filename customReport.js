const {
  Connection,
  Statement,
  SQL_ATTR_DBC_SYS_NAMING,
  SQL_TRUE,
  IN,
  NUMERIC,
  CHAR
} = require('idb-pconnector');
var Excel = require("exceljs");
var nodemailer = require("nodemailer");
var fs = require("fs");
var util = require("util");
var moment = require("moment");
var tools = require('./tools'); //references tools.js, put frequently used functions like finding date ranges ect here and call tools.function()

const connection = new Connection({
  url: '*LOCAL'
});

generateReport();

async function generateReport() {
  await connection.setConnAttr(SQL_ATTR_DBC_SYS_NAMING, SQL_TRUE);
  await executeStatement("CALL QSYS2.QCMDEXC('CHGLIBL (HCSLIB)')");

  //Fetch a list of customers - we create one document per company&customer
  var customerList = await executeStatement("SELECT wco, wcust, wemail FROM table1 left join table2 on MCMSTCLS = wmscls group by wco, wcust, wemail");
  var currentCustomer;

  var excelRow, workbook, worksheet;

  console.log(customerList);
  //Loop through our customer list 
  for (var i = 0; i < customerList.length; i++) {
    currentCustomer = customerList[i];
    tools.normaliseObject(currentCustomer);

    workbook = new Excel.Workbook(); //create the workbook
    await workbook.xlsx.readFile("./ExcelTemplates/outPutFile.xlsx"); //read template file into workbook
    worksheet = workbook.getWorksheet(1);
    tools.excelPrinterFriendly(worksheet);

    //Fetch header information for the document
    var customerInfo = await executeStatement("SELECT  * FROM table3 WHERE CO = " + currentCustomer.WCO + " AND CUST = " + currentCustomer.WCUST);

    tools.normaliseObject(customerInfo[0]);

    //Set the headers
    tools.setCell(worksheet, "C2", customerInfo[0].CUST);
    tools.setCell(worksheet, "B3", customerInfo[0].CNAME);
    tools.setCell(worksheet, "B4", customerInfo[0].CADDR1);
    tools.setCell(worksheet, "B5", customerInfo[0].CADDR2);
    tools.setCell(worksheet, "B6", customerInfo[0].CCITY + ', ' + customerInfo[0].CSTATE);
    tools.setCell(worksheet, "B7", "Salesrep: " + customerInfo[0].CSLM);
    tools.setCell(worksheet, "K2", "Date: " + moment().format("MM/DD/YYYY"));

    if (currentCustomer.WCO == someNumber) {
      tools.setCell(worksheet, "D2", "Header Text 1");
      tools.setCell(worksheet, "E3", "Header Info 1");
    } else {
      tools.setCell(worksheet, "D2", "Header Text 2");
      tools.setCell(worksheet, "E3", "Header Text 2");
      tools.setCell(worksheet, "E4", "Header Text 2");
    }

    excelRow = 10;

    //Next, fetch the order guide based on the company&customer data to populate the document
    var orderGuideData = await executeStatement("SELECT * FROM table1 left join table2 on MCMSTCLS = wmscls where WCO = " + currentCustomer.WCO + " AND WCUST = " + currentCustomer.WCUST + " ORDER BY WCO, WCUST, WMSCLS , WPROD");
    for (var y = 0; y < orderGuideData.length; y++) {
      tools.normaliseObject(orderGuideData[y]);
      tools.normaliseObject(orderGuideData[y+1]);

      if(y == 0)
      {
        worksheet.mergeCells("C" + excelRow + ":F" + excelRow);

        tools.setCell(worksheet, "A" + excelRow, " ", undefined, 'thinBorder');
        worksheet.getCell("A" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

        tools.setCell(worksheet, "B" + excelRow, String(orderGuideData[y].WMSCLS), undefined, 'thinBorder');
        worksheet.getCell("B" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
        worksheet.getCell("B" + excelRow).numFmt = '@';

        tools.setCell(worksheet, "C" + excelRow, orderGuideData[y].MCDESC, undefined, 'thinBorder');
        worksheet.getCell("C" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

        tools.setCell(worksheet, "G" + excelRow, " ", undefined, 'thinBorder');
        tools.setCell(worksheet, "H" + excelRow, " ", undefined, 'thinBorder');
        tools.setCell(worksheet, "I" + excelRow, " ", undefined, 'thinBorder');
        tools.setCell(worksheet, "J" + excelRow, " ", undefined, 'thinBorder');
        tools.setCell(worksheet, "K" + excelRow, " ", undefined, 'thinBorder');
        tools.setCell(worksheet, "L" + excelRow, " ", undefined, 'thinBorder');

        excelRow++;
      }

      if (y + 1 < orderGuideData.length) {
        if (orderGuideData[y + 1].WPROD != orderGuideData[y].WPROD) {
          worksheet.mergeCells("C" + excelRow + ":F" + excelRow);

          var productNumber = orderGuideData[y].WPROD;
          productNumber = productNumber.toString();
          productNumber = productNumber.padStart(6, '0')

          tools.setCell(worksheet, "A" + excelRow, productNumber, undefined, 'thinBorder');
          worksheet.getCell("A" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

          tools.setCell(worksheet, "B" + excelRow, String(orderGuideData[y].WCPROD), undefined, 'thinBorder');
          worksheet.getCell("B" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
          worksheet.getCell("B" + excelRow).numFmt = '@';

          tools.setCell(worksheet, "C" + excelRow, orderGuideData[y].WDESC, undefined, 'thinBorder');
          worksheet.getCell("C" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

          tools.setCell(worksheet, "G" + excelRow, orderGuideData[y].WCPACK, undefined, 'thinBorder');
          worksheet.getCell("G" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
          worksheet.getCell("G" + excelRow).numFmt = '@';

          tools.setCell(worksheet, "H" + excelRow, orderGuideData[y].WCOST, undefined, 'thinBorder');
          worksheet.getCell("H" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
          worksheet.getCell("H" + excelRow).numFmt = '0.00';
          tools.setCell(worksheet, "I" + excelRow, " ", undefined, 'thinBorder');
          tools.setCell(worksheet, "J" + excelRow, " ", undefined, 'thinBorder');
          tools.setCell(worksheet, "K" + excelRow, " ", undefined, 'thinBorder');
          tools.setCell(worksheet, "L" + excelRow, " ", undefined, 'thinBorder');
  
          excelRow++;

          if (orderGuideData[y + 1].WMSCLS != orderGuideData[y].WMSCLS) {
            worksheet.mergeCells("C" + excelRow + ":F" + excelRow);

            tools.setCell(worksheet, "A" + excelRow, " ", undefined, 'thinBorder');    
            worksheet.getCell("A" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

            tools.setCell(worksheet, "B" + excelRow, String(orderGuideData[y].WMSCLS), undefined, 'thinBorder');
            worksheet.getCell("B" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
            worksheet.getCell("B" + excelRow).numFmt = '@';

            tools.setCell(worksheet, "C" + excelRow, orderGuideData[y].MCDESC, undefined, 'thinBorder');
            worksheet.getCell("C" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

            tools.setCell(worksheet, "G" + excelRow, " ", undefined, 'thinBorder');
            tools.setCell(worksheet, "H" + excelRow, " ", undefined, 'thinBorder');    
            tools.setCell(worksheet, "I" + excelRow, " ", undefined, 'thinBorder');
            tools.setCell(worksheet, "J" + excelRow, " ", undefined, 'thinBorder');
            tools.setCell(worksheet, "K" + excelRow, " ", undefined, 'thinBorder');
            tools.setCell(worksheet, "L" + excelRow, " ", undefined, 'thinBorder');
    
            excelRow++;
          }

        }

      }
      else if (y + 1 == orderGuideData.length)
      { 
        worksheet.mergeCells("C" + excelRow + ":F" + excelRow);

        var productNumber = orderGuideData[y].WPROD;
        productNumber = productNumber.toString();
        productNumber = productNumber.padStart(6, '0')
        tools.setCell(worksheet, "A" + excelRow, productNumber, undefined, 'thinBorder');
        worksheet.getCell("A" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

        tools.setCell(worksheet, "B" + excelRow, String(orderGuideData[y].WCPROD), undefined, 'thinBorder');
        worksheet.getCell("B" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
        worksheet.getCell("B" + excelRow).numFmt = '@';

        tools.setCell(worksheet, "C" + excelRow, orderGuideData[y].WDESC, undefined, 'thinBorder');
        worksheet.getCell("C" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column

        tools.setCell(worksheet, "G" + excelRow, orderGuideData[y].WCPACK, undefined, 'thinBorder');
        worksheet.getCell("G" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
        worksheet.getCell("G" + excelRow).numFmt = '@';

        tools.setCell(worksheet, "H" + excelRow, orderGuideData[y].WCOST, undefined, 'thinBorder');
        worksheet.getCell("H" + excelRow).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};  //trying to center column
        worksheet.getCell("H" + excelRow).numFmt = '0.00';

        tools.setCell(worksheet, "I" + excelRow, " ", undefined, "thinBorder");
        tools.setCell(worksheet, "J" + excelRow, " ", undefined, "thinBorder");
        tools.setCell(worksheet, "K" + excelRow, " ", undefined, "thinBorder");
        tools.setCell(worksheet, "L" + excelRow, " ", undefined, "thinBorder");

        excelRow++;

      }

    }

    worksheet.pageSetup.printArea = 'A1:L' + excelRow +'';
    console.log(excelRow);
    console.log('Writing file...');
    await workbook.xlsx.writeFile("./ExcelOutput/outputFile" + currentCustomer.WCO + currentCustomer.WCUST + ".xlsx");

    console.log('Deleting from table: ' + currentCustomer);
    // "DELETE FROM mastwrxn WHERE WCO = " & iHCo & " AND WCUST = " & iHCust
    await executeStatement("DELETE FROM table1 WHERE WCO = " + currentCustomer.WCO + " AND WCUST = " + currentCustomer.WCUST + " with nc");
    
    //Send email
    var transporter = nodemailer.createTransport({
      logger: true,
      port: 587,
      host: "outlook.office365.com",
      auth: {
        user: "user",
        pass: "password"
      },
      tls: {
        ciphers: "SSLv3"
      }
    });

    console.log("Sending Report to " + currentCustomer.WEMAIL + " ...");

    var data = {
      from: sendingAddress,
      subject: subjectLineText + currentCustomer.WCO + "-" + currentCustomer.WCUST,
      text: mainBodyText + currentCustomer.WCO + "-" + currentCustomer.WCUST,
      attachments: [{
        filename: outputFileName + currentCustomer.WCO + "-" + currentCustomer.WCUST + ".xlsx",
        path: "./Path/toFile/Name" + currentCustomer.WCO + currentCustomer.WCUST + ".xlsx"
      }],
      to: sendToArray 
    }

    console.log("Report Sent Successfully");

    try {
      var response = await transporter.sendMail(data); 
      console.log(response);

      fs.unlink('./Path/toFile/Name' + currentCustomer.WCO + currentCustomer.WCUST + '.xlsx', function (err) {
        if (err) throw err;
        // if no error, file has been deleted successfully
        console.log('File deleted Successfully!');
      });
    } catch (e) {
      console.log(e)
    }

  }// END FOR LOOP
}// END GENERATEREPORT()


async function executeStatement(sqlStatement) {
  const statement = new Statement(connection);

  const results = await statement.exec(sqlStatement);

  await statement.close();

  return results;
}
