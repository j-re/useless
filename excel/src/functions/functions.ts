/* global CustomFunctions */
import { AccountManager } from "../taskpane/authConfig";
const accountManager = new AccountManager();
const accountManagerInitialization = Office.onReady().then(() => accountManager.initialize());

type AccessTokenClaims = {
  aud?: string;
  scp?: string;
  roles?: string[];
  tid?: string;
  oid?: string;
  appid?: string;
};

function decodeAccessTokenClaims(accessToken: string): AccessTokenClaims | null {
  try {
    const tokenParts = accessToken.split(".");
    if (tokenParts.length < 2) {
      return null;
    }

    const payload = tokenParts[1];
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), "=");
    const decodedJson = atob(padded);
    return JSON.parse(decodedJson) as AccessTokenClaims;
  } catch (error) {
    console.warn("Unable to decode access token claims.", error);
    return null;
  }
}

function logTokenClaims(accessToken: string): void {
  const claims = decodeAccessTokenClaims(accessToken);
  if (!claims) {
    console.warn("No token claims were decoded.");
    return;
  }

  console.log("Access token claims", {
    aud: claims.aud,
    scp: claims.scp,
    roles: claims.roles,
    tid: claims.tid,
    oid: claims.oid,
    appid: claims.appid,
  });
}

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @returns Sum of the two numbers.
 */
export async function add(first: number, second: number): Promise<number> {
  try {
    
    await accountManagerInitialization;
    const accessToken = await accountManager.acquireToken(
      ["api://651cdece-d8c6-4840-8c58-79acf2a8bb77/access_as_user"]
    );
    logTokenClaims(accessToken);
    const requestUrl = `http://localhost:3900/add?first=${encodeURIComponent(first)}&second=${encodeURIComponent(second)}`;
        
    const headers: HeadersInit = {
      Accept: "application/json",
    };

    if (accessToken) {
      headers.Authorization = `Bearer ${accessToken}`;
    }

    const response = await fetch(requestUrl, {
      method: "GET",
      headers,
    });

    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}`);
    }

    const data = (await response.json()) as { result?: number };
    if (typeof data.result === "number") {
      return data.result;
    }
  } catch (error) {
    console.log(error);
  }

  return first + second;
}

CustomFunctions.associate("ADD", add);
