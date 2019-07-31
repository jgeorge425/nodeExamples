const {
    Connection,
    Statement,
    SQL_ATTR_DBC_SYS_NAMING,
    SQL_TRUE
} = require('idb-pconnector');
const fs = require('fs');
var tools = require('./tools'); //references tools.js, put frequently used functions like finding date ranges ect here and call tools.function()
var Excel = require("exceljs");
var nodemailer = require("nodemailer");
var moment = require("moment");

const connection = new Connection({
    url: '*LOCAL'
});

generateReport();

async function generateReport() {
    await connection.setConnAttr(SQL_ATTR_DBC_SYS_NAMING, SQL_TRUE);
    await executeStatement("CALL QSYS2.QCMDEXC('CHGLIBL (HCSLIB)')");

    var sendToArray //Used to store the email addresses that will recieve output email
    var timeStamp = moment();


    var workbook = new Excel.Workbook(); //create the workbook
    var sheet = workbook.addWorksheet("Order Summary");

    sheet.columns = [{
            header: 'Invoice#',
            key: 'InvoiceNumber',
            width: 20
        },
        {
            header: 'Invoice Date (MM/DD/YYYY)',
            key: 'InvoiceDate',
            width: 30
        },
        {
            header: 'Customer #',
            key: 'CustomerNumber',
            width: 18
        },
        {
            header: 'Sold-To Name',
            key: 'SoldToName',
            width: 40
        },
        {
            header: 'Ship-To Name',
            key: 'ShipToName',
            width: 40
        },
        {
            header: 'PO#',
            key: 'PONumber',
            width: 20
        },
        {
            header: 'Invoice Amt',
            key: 'InvoiceAmt',
            width: 14
        }
    ];

    sheet.getRow(1).font = {
        bold: true
    };

    var excelRow = 2;
    var excelModified = false;
    var previousBusinessDay;

    var items = await executeStatement("select table1.NHINV# as nhiv, table1.NHIDAT, table1.NHCUST, table1.NHBTNM, table1.NHSHNM, table1.NHCUPO, table1.NHITOT from table1 join table2  on table2.excust = table1.nhcust and table2.exloc = table1.nhloc where table2.exid = 'CompanyId'");

    for (var i = 0; i < items.length; i++) {
        var currentItems = items[i];

        currentItemDay = moment(currentItems.NHIDAT);

        tools.setCell(sheet, "A" + excelRow, Number(currentItems.NHIV)); //Inv. Number
        tools.setCell(sheet, "B" + excelRow, currentItemDay.format("MM/DD/YYYY")); //Inv. Date
        tools.setCell(sheet, "C" + excelRow, Number(currentItems.NHCUST)); //Customer #
        tools.setCell(sheet, "D" + excelRow, currentItems.NHBTNM); //Sold-To Name
        tools.setCell(sheet, "E" + excelRow, currentItems.NHSHNM); //Ship-to Name
        tools.setCell(sheet, "F" + excelRow, currentItems.NHCUPO); //PO#
        tools.setCell(sheet, "G" + excelRow, Number(currentItems.NHITOT), '$#,##0.00_);($#,##0.00)'); //Inv. Total (Total Sales Amt. + Freight, etc.): NHITOT       Total Sales Amount:NHSALES

        excelModified = true;
        excelRow++;
        previousBusinessDay = moment(currentItems.NHIDAT);

    } //END shipto for loop

    if (excelModified) {



        await workbook.xlsx.writeFile(outputFileName + timeStamp.format("HH-mm-ss-S") + ".xlsx");

        var transporter = nodemailer.createTransport({
            logger: true,
            port: 587,
            host: "outlook.office365.com",
            auth: {
                user: "userName",
                pass: "password"
            },
            tls: {
                ciphers: "SSLv3"
            }
        });

        var data = {
            from: sentFromAddress,
            subject: subjectText + previousBusinessDay.format("MM/DD/YYYY"),
            text: mainText + previousBusinessDay.format("MM/DD/YYYY"),
            attachments: [{
                filename: outputFileName + previousBusinessDay.format("MMDDYYYY") + ".xlsx",
                path: "./Path/toFile/fileName" + timeStamp.format("HH-mm-ss-S") + ".xlsx"
            }],
            to: sendToArray
        }

        try {
            var response = await transporter.sendMail(data);
            console.log(response);
            console.log("Report Sent Successfully");

            fs.unlink(outputFileName + timeStamp.format("HH-mm-ss-S") + ".xlsx", function (err) {
                if (err) throw err;
                // if no error, file has been deleted successfully
                console.log('File deleted Successfully!');
            });
        } catch (e) {
            console.log(e)
        }
    } //END if excel was modified

} //END generateReport()



async function executeStatement(sqlStatement) {
    const statement = new Statement(connection);

    const results = await statement.exec(sqlStatement);
    statement.close();

    return results
}