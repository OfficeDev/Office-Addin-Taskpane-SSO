/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, OfficeExtension */

import * as excel from "./excel";
import * as powerpoint from "./powerpoint";
import * as word from "./word";

export function writeDataToOfficeDocument(result: string[]): Promise<any> {
  return new OfficeExtension.Promise(function(resolve, reject) {
    try {
      switch (Office.context.host) {
        case Office.HostType.Excel:
          excel.writeDataToOfficeDocument(result);
          break;
        case Office.HostType.PowerPoint:
          powerpoint.writeDataToOfficeDocument(result);
          break;
        case Office.HostType.Word:
          word.writeDataToOfficeDocument(result);
          break;
        default:
          throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
      }
      resolve();
    } catch (error) {
      reject(Error("Unable to write data to document. " + error.toString()));
    }
  });
}
