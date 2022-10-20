/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthhelper";
import { getUserData } from "./msgraph-helper";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global OfficeRuntime */

let retryGetBootstrapToken = 0;

export async function getGraphData(callback): Promise<void> {
  try {
    let bootstrapToken: string = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response: any = await getUserData(bootstrapToken);
    if (!response) {
      return Promise.reject();
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaBootstrapToken: string = await OfficeRuntime.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = getUserData(mfaBootstrapToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
      Promise.resolve();
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
      Promise.reject();
    }
  }
}

function handleAADErrors(response: any, callback: any): void {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetBootstrapToken <= 0) {
    retryGetBootstrapToken++;
    getGraphData(callback);
  } else {
    dialogFallback(callback);
  }
}
