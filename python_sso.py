import os
import sys
from typing import Any

import msal
import requests


def get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def parse_scopes(raw_scopes: str) -> list[str]:
    # Supports comma- or space-separated scopes
    scopes = [s.strip() for s in raw_scopes.replace(",", " ").split() if s.strip()]
    if not scopes:
        raise ValueError("ENTRA_SCOPES must contain at least one scope")
    return scopes


def acquire_access_token_interactive(tenant_id: str, client_id: str, scopes: list[str]) -> tuple[str, str | None]:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
    )

    result: dict[str, Any] | None = None

    # Try cached token first (if available)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])

    # Fall back to interactive user sign-in (authorization code flow + PKCE)
    if not result or "access_token" not in result:
        result = app.acquire_token_interactive(
            scopes=scopes,
            prompt="select_account",
            redirect_uri="http://localhost",
        )

    access_token = result.get("access_token") if result else None
    id_token = result.get("id_token") if result else None

    if not access_token:
        error = (result or {}).get("error", "unknown_error")
        description = (result or {}).get("error_description", "No error description returned.")
        raise RuntimeError(f"Token acquisition failed: {error} - {description}")

    return access_token, id_token


def call_endpoint(url: str, bearer_token: str) -> requests.Response:
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Accept": "application/json",
    }
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response


def main() -> None:
    try:
        tenant_id = "2f5d408f-3037-4aa2-b0f6-c2dadfb05497" #get_required_env("ENTRA_TENANT_ID")
        client_id = "651cdece-d8c6-4840-8c58-79acf2a8bb77" #get_required_env("ENTRA_CLIENT_ID")
        scopes = ["api://651cdece-d8c6-4840-8c58-79acf2a8bb77/access_as_user"] #parse_scopes(get_required_env("ENTRA_SCOPES"))
        endpoint = "https://api-app.politewave-fab973fc.australiaeast.azurecontainerapps.io" #get_required_env("API_ENDPOINT_URL")

        access_token, id_token = acquire_access_token_interactive(tenant_id, client_id, scopes)

        print("Access Token:")
        print(access_token)
        print("ID Token:")
        print(id_token if id_token else "No id_token returned")

        response = call_endpoint(endpoint, access_token)

        print(f"Status: {response.status_code}")
        print(response.text)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
