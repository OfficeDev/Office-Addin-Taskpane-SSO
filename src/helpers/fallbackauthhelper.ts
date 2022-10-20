/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, location, Office */

import { getUserData } from "./msgraph-helper";
import { showMessage } from "./message-helper";
import { publicClientApp, loginRequest } from "./fallbackauthdialog";

let loginDialog: Office.Dialog = null;
let homeAccountId = null;
let callbackFunction = null;

export async function dialogFallback(callback) {
  // Attempt to acquire token silently if user is already signed in.
  if (homeAccountId !== null) {
    const result = await publicClientApp.acquireTokenSilent(loginRequest);
    if (result !== null && result.accessToken !== null) {
      const response = await getUserData(result.accessToken);
      callbackFunction(response);
    }
  } else {
    callbackFunction = callback;

    // We fall back to Dialog API for any error.
    const url = "/fallbackauthdialog.html";
    showLoginPopup(url);
  }
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  console.log("Message received in processMessage: " + JSON.stringify(arg));
  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    // We now have a valid access token.
    loginDialog.close();

    // Configure MSAL to use the signed-in account as the active account for future requests.
    homeAccountId = messageFromDialog.accountId; // Track the account id for future silent token requests.
    const homeAccount = publicClientApp.getAccountByHomeId(messageFromDialog.accountId);
    publicClientApp.setActiveAccount(homeAccount);

    const response = await getUserData(messageFromDialog.result);
    callbackFunction(response);
  } else if (messageFromDialog.error === undefined && messageFromDialog.result.errorCode === undefined) {
    // Need to pick the user to use to auth
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
    if (messageFromDialog.error) {
      showMessage(JSON.stringify(messageFromDialog.error.toString()));
    } else if (messageFromDialog.result) {
      showMessage(JSON.stringify(messageFromDialog.result.errorMessage.toString()));
    }
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}
