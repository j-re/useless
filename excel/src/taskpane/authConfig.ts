// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, document*/

/// <reference types="office-js" />

import {
  BrowserAuthError,
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./msalconfig";
import { createLocalUrl } from "./util";
import { getTokenRequest } from "./msalcommon";

export type AuthDialogResult = {
  accessToken?: string;
  error?: string;
};

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

// Constants
const DIALOG_DIMENSIONS = { height: 60, width: 30 } as const;
const DIALOG_CLOSED_ERROR_CODE = 12006;
const POPUP_WINDOW_ERROR_CODE = "popup_window_error";
const NESTED_APP_AUTH_REQUIREMENT = { name: "NestedAppAuth", version: "1.1" } as const;
const FORCE_FALLBACK_DIALOG_KEY = "forceFallbackDialog";

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<string> | null = null;

  private isNestedAppAuthSupported(): boolean {
    return Office.context.requirements.isSetSupported(
      NESTED_APP_AUTH_REQUIREMENT.name, 
      NESTED_APP_AUTH_REQUIREMENT.version
    );
  }

  private shouldForceFallbackDialog(): boolean {
    return window.localStorage.getItem(FORCE_FALLBACK_DIALOG_KEY) === "true";
  }

  // Initialize MSAL public client application.
  async initialize(): Promise<void> {
    try {
      // Make sure office.js is initialized.
      await Office.onReady();

      // Initialize a nested public client application.
      this.pca = await createNestablePublicClientApplication(msalConfig);
    } catch (error) {
      console.error("Failed to initialize AccountManager:", error);
      throw new Error(`Initialization failed: ${error}`);
    }
  }

  // Get login hint for Word, Excel, or PowerPoint on the web from the auth context.
  private async getLoginHint(): Promise<string | undefined> {
    try {
      if (typeof Office !== "undefined" && Office.context) {
            const authContext = await Office.auth.getAuthContext();
            if (authContext?.userPrincipalName) return authContext.userPrincipalName;
        }
    } catch (error) {
      console.warn("Could not get login hint:", error);
    }
    return undefined;
  }

  async clearCache(): Promise<void> {
    if (this.pca) {
      // const accounts = this.pca.getAllAccounts();
      // for (const account of accounts) {
      //   await this.pca.logoutPopup({ account });
      // }
      // // Clear MSAL's internal cache
      await this.pca.clearCache();
      sessionStorage.clear();
      localStorage.clear();
      console.log("Token cache cleared successfully.");
    } else {
      console.warn("AccountManager is not initialized. Cannot clear cache.");
    }
  }

  async acquireToken(scopes: string[], allowInteractive = true): Promise<string> {
    // Check if the user is already signed in via fallback dialog API.
    if (this._dialogApiResult) {
      return this._dialogApiResult;
    }
    
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }
    const loginHint = await this.getLoginHint();
    console.log(loginHint);

    if (this.shouldForceFallbackDialog()) {
      if (!allowInteractive) {
        throw new Error("Fallback dialog is forced, but interactive authentication is disabled for this call.");
      }
      console.log("Fallback dialog is forced; skipping NAA silent flow.");
      return this.getTokenWithDialogApi(scopes);
    }
    
    try {
      console.log("Trying to acquire token silently...");
      const tokenRequest = getTokenRequest(scopes, false, undefined, loginHint);

      const account =
        this.pca!.getActiveAccount() ??
        (loginHint ? this.pca!.getAccount({ username: loginHint }) : null) ??
        this.pca!.getAllAccounts()[0];

      let authResult;
      if (account) {
        authResult = await this.pca!.acquireTokenSilent({
          ...tokenRequest,
          account,
          forceRefresh: true,
        });
      } else {
        // No account yet: bootstrap with ssoSilent (no forceRefresh support here)
        authResult = await this.pca!.ssoSilent(tokenRequest);
      }
      this.pca.setActiveAccount(authResult.account);
      console.log("Acquired token silently.");
      console.log("Access token claims", authResult);
      return authResult.accessToken;
    } catch (silentError) {
      if (silentError instanceof InteractionRequiredAuthError) {
        if (!allowInteractive) {
          throw new Error(`Silent token acquisition requires user interaction: ${String(silentError)}`);
        }
        return this.acquireTokenInteractively(scopes, loginHint);
      } else {
        if (!allowInteractive) {
          throw new Error(`Unable to acquire token silently: ${String(silentError)}`);
        }
        // For running on a localhost server, use the following line of code
        // to work around CORS errors with localhost.
        // Comment this code when deploying to production.
        return this.acquireTokenInteractively(scopes, loginHint);

        // For production uncomment the following code.
        // throw new Error(`Unable to acquire access token: ${silentError}`);
        
      }
    }
  }

  private async acquireTokenInteractively(scopes: string[], loginHint: string | undefined): Promise<string> {
    try {
      console.log("Trying to acquire token interactively...");
      
      const authResult = await this.pca!.acquireTokenPopup(
        getTokenRequest(scopes, false, undefined, loginHint)
      );
      this.pca!.setActiveAccount(authResult.account);
      console.log("Acquired token interactively.");
      return authResult.accessToken;
    } catch (popupError) {
      return this.handleInteractiveTokenError(popupError);
    }
  }

  private async handleInteractiveTokenError(popupError: unknown): Promise<string> {
    // Optional fallback if about:blank popup should not be shown
    if (popupError instanceof BrowserAuthError && popupError.errorCode === POPUP_WINDOW_ERROR_CODE) {
      const accessToken = await this.getTokenWithDialogApi(scopes);
      return accessToken;
    } else {
      // Acquire token interactive failure.
      console.error(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(scopes?: string[]): Promise<string> {
    const scopeQuery = scopes && scopes.length > 0 ? `&scopes=${encodeURIComponent(scopes.join(" "))}` : "";
    this._dialogApiResult = new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?source=taskpane${scopeQuery}`), 
        DIALOG_DIMENSIONS, 
        (result: any) => {
          result.value.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
            if ((arg as DialogEventError).error === DIALOG_CLOSED_ERROR_CODE) {
              this._dialogApiResult = null;
              reject("Dialog closed");
            }
          });
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
            const parsedMessage = JSON.parse((arg as DialogEventMessage).message);
            result.value.close();
            if (parsedMessage.error) {
              this._dialogApiResult = null;
              reject(parsedMessage.error);
            } else {
              resolve(parsedMessage.accessToken);
            }
          });
        }
      );
    });
    return this._dialogApiResult;
  }

  /**
   * Clean up resources and event listeners
   */
  cleanup(): void {
    this._dialogApiResult = null;
  }
}
