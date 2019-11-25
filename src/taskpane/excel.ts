/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Excel, Office */

import { getGraphData } from "./../helpers/graphHelper";

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function() {
      $("#getGraphDataButton").click(getGraphData);
    });
  }
});

export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo: string[] = [];
    userProfileInfo.push(result["displayName"]);
    userProfileInfo.push(result["jobTitle"]);
    userProfileInfo.push(result["mail"]);
    userProfileInfo.push(result["mobilePhone"]);
    userProfileInfo.push(result["officeLocation"]);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
