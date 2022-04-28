/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const POPUP_ADDRESS = "https://localhost:3000/confirm-popup.html";
let dialog = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("log-preview-btn").onclick = logPreviewConfirmDialog;
    document.getElementById("log-time-btn").onclick = logTimeConfirmDialog;

    document.getElementById("pay-run-date").valueAsDate = new Date();

    createDropDown();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function createDropDown() {
  await Excel.run(async (context) => {
    const driverSheet = context.workbook.worksheets.getItem("Sheet1");
    const driverTable = driverSheet.tables.getItem("PreviewTable");
    const headerRng = driverTable.getHeaderRowRange();
    // const DriverDataRng = driverTable.getDataBodyRange();

    const optionElement = document.getElementById("pd-headers");

    headerRng.load("values, cellCount");
    await context.sync();

    let headers = [];

    for (let i = 6; i < headerRng.cellCount - 1; i++) {
      let cell = headerRng.getCell(0, i);

      cell.load("text, values");
      headers.push(cell);
    }

    await context.sync();

    for (let i = 0; i < headers.length; i++) {
      let headerText = headers[i].text.toString();
      // console.log(headerText);

      let option = document.createElement("option");
      option.value = headerText;
      option.text = headerText;
      optionElement.add(option);
    }
  });
}

function logPreviewConfirmDialog() {
  Office.context.ui.displayDialogAsync(
    POPUP_ADDRESS,
    { height: 30, width: 30, displayInIframe: true },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();

        if (arg.message == "YES") {
          logPreview();
        } else if (arg.message == "NO") {
          console.log("No: " + arg.message);
        }
      });
    }
  );
}

async function logPreview() {
  await Excel.run(async (context) => {
    const driverSheet = context.workbook.worksheets.getItem("Sheet1");
    const driverTable = driverSheet.tables.getItem("PreviewTable");
    //const headerRng = driverTable.getHeaderRowRange();
    const DriverDataRng = driverTable.getDataBodyRange();

    const pgField = document.getElementById("paygroups");
    let paygroupsText = pgField.value;

    const pdHeaders = document.getElementById("pd-headers");
    let targetHeader = pdHeaders.value;

    DriverDataRng.load("values");
    await context.sync();

    const runDateElement = document.getElementById("pay-run-date");

    // let dt = new Date("2022-4-12");
    let dt = new Date(runDateElement.value + "T00:00:00");
    let y = new Intl.DateTimeFormat("en", { year: "numeric" }).format(dt);
    let m = new Intl.DateTimeFormat("en", { month: "2-digit" }).format(dt);
    let d = new Intl.DateTimeFormat("en", { day: "2-digit" }).format(dt);

    // let paygroup = "GB8";
    // let id = paygroup + m + d + y;

    let IDRng = driverTable.columns.getItem("ID").getDataBodyRange().load("values");
    let targetHeaderRng = driverTable.columns.getItem(targetHeader).getDataBodyRange();

    await context.sync();

    let paygroups = paygroupsText.split(/\r?\n/);

    paygroups.forEach((paygroup) => {
      let id = paygroup + m + d + y;
      let matchIndex = IDRng.values.findIndex((value) => value[0] == id);

      targetHeaderRng.getCell(matchIndex, 0).values = null;
      Log(paygroup + " added for preview @ " + targetHeader, "good");
    });

    await context.sync();
  });
}

function logTimeConfirmDialog() {
  Office.context.ui.displayDialogAsync(
    POPUP_ADDRESS,
    { height: 30, width: 30, displayInIframe: true },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();

        if (arg.message == "YES") {
          logTime();
        } else if (arg.message == "NO") {
          console.log("No: " + arg.message);
        }
      });
    }
  );
}

async function logTime() {
  await Excel.run(async (context) => {
    const driverSheet = context.workbook.worksheets.getItem("Sheet1");
    const driverTable = driverSheet.tables.getItem("PreviewTable");
    const DriverDataRng = driverTable.getDataBodyRange();

    const pgField = document.getElementById("paygroups");
    let paygroupsText = pgField.value;

    const pdHeaders = document.getElementById("pd-headers");
    let targetHeader = pdHeaders.value;

    console.log(targetHeader);

    DriverDataRng.load("values");
    await context.sync();

    const runDateElement = document.getElementById("pay-run-date");

    let dt = new Date(runDateElement.value + "T00:00:00");
    let y = new Intl.DateTimeFormat("en", { year: "numeric" }).format(dt);
    let m = new Intl.DateTimeFormat("en", { month: "2-digit" }).format(dt);
    let d = new Intl.DateTimeFormat("en", { day: "2-digit" }).format(dt);

    let IDRng = driverTable.columns.getItem("ID").getDataBodyRange().load("values");
    let targetHeaderRng = driverTable.columns.getItem(targetHeader).getDataBodyRange();

    await context.sync();

    let paygroupLines = paygroupsText.split(/\r?\n/);

    paygroupLines.forEach((line) => {
      const lineArr = line.trim().split(/\s+/);

      if (lineArr.length > 1) {
        let paygroup = lineArr[0];
        let paygroupTime = lineArr[1];

        let id = paygroup + m + d + y;
        let matchIndex = IDRng.values.findIndex((value) => value[0] == id);

        targetHeaderRng.getCell(matchIndex, 0).values = paygroupTime;

        Log(paygroup + " added with value " + paygroupTime, "good");
      }
    });

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function Log(msg, status) {
  const logElement = document.getElementById("log");

  let currentdate = new Date();
  let datetime =
    currentdate.getMonth() +
    1 +
    "/" +
    currentdate.getDate() +
    "/" +
    currentdate.getFullYear() +
    " " +
    currentdate.getHours() +
    ":" +
    currentdate.getMinutes() +
    ":" +
    currentdate.getSeconds();

  var span = document.createElement("span");
  span.textContent = msg + " : " + datetime;
  span.setAttribute("class", status);
  logElement.prepend(span);
}
