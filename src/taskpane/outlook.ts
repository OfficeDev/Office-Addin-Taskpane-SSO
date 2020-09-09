/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
/* global $, document, Office */

import { getGraphData } from "./../helpers/ssoauthhelper";

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
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
    Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
