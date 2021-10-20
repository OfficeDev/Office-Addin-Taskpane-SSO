/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../assets/icon-16.png";
import icon32 from "../../assets/icon-32.png";
import icon64 from "../../assets/icon-64.png";
import icon80 from "../../assets/icon-80.png";
import icon128 from "../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global $, document, Office */

import { getGraphData } from "./../helpers/ssoauthhelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    $(document).ready(function () {
      $("#getGraphDataButton").click(getGraphData);
    });
  }
});

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
