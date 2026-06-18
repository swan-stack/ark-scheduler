import { Configuration, LogLevel } from "@azure/msal-browser";

const clientId = import.meta.env.VITE_MSAL_CLIENT_ID || "REPLACE_WITH_ENTRA_APP_CLIENT_ID";
const tenantId = import.meta.env.VITE_MSAL_TENANT_ID || "common";

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      loggerCallback: (_level, message, containsPii) => {
        if (!containsPii && import.meta.env.DEV) console.log(message);
      },
      logLevel: LogLevel.Warning,
    },
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "User.ReadBasic.All",
    "Calendars.Read",
    "Calendars.Read.Shared",
  ],
};
