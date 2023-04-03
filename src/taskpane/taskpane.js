/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
const axios = require('axios')
// import axios from '../../axios'
let range_change = {}
let format_x

Office.onReady(async (info) => {

  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("save").onclick = () => tryCatch(saveEvent);
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets;
      console.log('aa')
      worksheet.onChanged.add(handleChange);
      worksheet.onActivated.add(handleChange);
      // let range = worksheet.getRange('A2')
      // range.load('format/fill/color , format/borders')
      await context.sync();
      // format_x = range.format
    });

    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets;
      sheet.onAdded.add(onWorksheetAdd);
      await context.sync();
    });
  }
});


let onWorksheetAdd = async (event) => {
  console.log(event)
  await Excel.run(async (context) => {
    await context.sync();
    var ws_id = event.worksheetId
    const worksheet = context.workbook.worksheets.getItem(ws_id);
    worksheet.load("name")
    worksheet.onChanged.add(handleChange);
    worksheet.onActivated.add(handleChange);
    console.log(worksheet)
    await context.sync();
  });
}

let saveEvent = async () => {
  await Excel.run(async (context) => {
    // const sheet = context.workbook.worksheets.getActiveWorksheet();
    // console.log(sheet)

    // await context.sync();
    // console.log(sheet['id'])
    for (let j in range_change) {
      for (let i in range_change[j]) {
        let sheet = context.workbook.worksheets.getItem(j);
        console.log(sheet)
        console.log(i)
        console.log(format_x)
        range = sheet.getRange(i);
        range.format.font.color = "black"

        await context.sync();

      }
    }

  });
  range_change = {}
}

const getData = async () => {
  try {
    return await axios.get('http://localhost:8090/mssql', {
      params: {
        query: `select * from dbo.react_app`,
        token: 123456
      },
      headers: { 'Content-Type': 'application/json' }
    })

  } catch (error) {
    console.error(error)
  }
}

async function handleChange(event) {
  let Data = await getData()
  console.log(Data['data']['data'])
  await Excel.run(async (context) => {
    await context.sync();
    var ws_id = event.worksheetId
    const worksheet = context.workbook.worksheets.getItem(ws_id);
    worksheet.load("name")
    console.log(worksheet)
    await context.sync();
    console.log(event)
    console.log("The activated worksheet ID is: " + ws_id);
    console.log("The activated worksheet Name is: " + worksheet.name);

    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Creating the SettableCellProperties objects to use for the range.
    // In your add-in, these should be created once, outside the function.


    const range = sheet.getRange(event['address']);
    range.load('format')
    // range.load('format/font')
    await context.sync();
    console.log(range)

    console.log('bbb')

    await context.sync();
    console.log(event['address'])
    if (event['address']) {
      range.format.font.color = "#ff0000"
      if (range_change[ws_id]) {
        range_change[ws_id][event['address']] = 1
      } else {
        range_change[ws_id] = {}
        range_change[ws_id][event['address']] = 1
      }
    }


    // await context.sync();
    // if (event['address'].includes(":")) {
    //   console.log('aaa')
    //   range_change[ws_id][event['address']] = 1
    //   // await context.sync();

    //   // range.format.fill.color = "#ff0000"
    //   range.format.font.color = "#ff0000"
    //   await context.sync();
    // } else if (event['details']['valueBefore'] != event['details']['valueAfter']) {
    //   console.log('testa')
    //   if (!range_change[ws_id]) {
    //     range_change[ws_id] = {}
    //   }
    //   range_change[ws_id][event['address']] = 1
    //   await context.sync();
    //   // range.format.fill.color = "#ff0000"
    //   range.format.font.color = "#ff0000"
    //   await context.sync();
    // }

    console.log(range_change)

    // You can use empty JSON objects to avoid changing a cell's properties.

  });
}



async function createTable() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.load('name')
    await context.sync();
    console.log(currentWorksheet)
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = currentWorksheet.name + "_ExpensesTable";

    expensesTable.getHeaderRowRange().values =
      [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    expensesTable.rows.deleteRows()
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();


    await context.sync();

    console.log("Added a worksheet-level data-changed event handler.");
  });

}
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}