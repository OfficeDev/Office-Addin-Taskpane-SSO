/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthdialog";
import { callGetUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global OfficeRuntime */

let retryGetMiddletierToken = 0;

export async function getUserData(callback): Promise<void> {
  try {
    let middletierToken: string = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response: any = await callGetUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken: string = await OfficeRuntime.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = callGetUserData(mfaMiddletierToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response: any, callback: any): void {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    getUserData(callback);
  } else {
    dialogFallback(callback);
  }
}
