// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { LogLevel } from "@azure/msal-browser";
import { createLocalUrl } from "./util";

/* global console */

export const clientId = "3e7b63eb-0830-45c6-95ef-d20ab6bba49c"; // Replace with your actual Application ID
export const msalConfig = {
  auth: {
    clientId,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Verbose,
      loggerCallback: (level: LogLevel, message: string) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
      piiLoggingEnabled: true,
    },
  },
};

// Default scopes to use in the fallback dialog.
export const defaultScopes = ["openid", "profile", "user.read", "files.read"];
