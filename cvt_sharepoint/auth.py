from __future__ import annotations

import json
import os
import shutil
import subprocess
import webbrowser
from pathlib import Path
from typing import Any

import msal

# Delegated Graph scopes. Short names like Files.ReadWrite are the same
# permissions; MSAL expects the Graph resource prefix.
GRAPH_SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "offline_access",
]

DEFAULT_CACHE_PATH = Path(".sharepoint_token_cache.json")
DEFAULT_AUTHORITY_TENANT = "organizations"


class SharePointAuthError(RuntimeError):
    pass


def cache_path() -> Path:
    return Path(os.environ.get("SHAREPOINT_TOKEN_CACHE", DEFAULT_CACHE_PATH))


def client_id_optional() -> str:
    return (os.environ.get("AZURE_CLIENT_ID") or "").strip()


def client_id() -> str:
    value = client_id_optional()
    if not value:
        raise SharePointAuthError(_no_graph_auth_help())
    return value


def _no_graph_auth_help() -> str:
    return (
        "No way to call Microsoft Graph without an app identity.\n"
        "Nscale has no OneDrive license and you cannot register Entra.\n"
        "Use SharePoint in the browser instead:\n"
        "  ./cvt sharepoint open\n"
        "  Download the xlsx, then:\n"
        "  ./cvt sharepoint use --file ~/Downloads/<file>.xlsx\n"
        "After edits, upload/replace that file in the same SharePoint library."
    )


def tenant_id() -> str:
    return (os.environ.get("AZURE_TENANT_ID") or DEFAULT_AUTHORITY_TENANT).strip()


def _load_cache(path: Path) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if path.is_file():
        cache.deserialize(path.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache, path: Path) -> None:
    if not cache.has_state_changed:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(cache.serialize())


def _app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id(),
        authority=f"https://login.microsoftonline.com/{tenant_id()}",
        token_cache=cache,
    )


def acquire_token_azure_cli(*, login_if_needed: bool = False) -> dict[str, Any] | None:
    """Use Microsoft's Azure CLI app. You do not register an Entra app."""
    az = shutil.which("az")
    if not az:
        return None

    def _read_token() -> dict[str, Any] | None:
        proc = subprocess.run(
            [
                az,
                "account",
                "get-access-token",
                "--resource",
                "https://graph.microsoft.com",
                "-o",
                "json",
            ],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            return None
        try:
            payload = json.loads(proc.stdout)
        except json.JSONDecodeError:
            return None
        token = payload.get("accessToken")
        if not token:
            return None
        return {
            "access_token": token,
            "expires_on": payload.get("expiresOn"),
            "source": "azure-cli",
        }

    result = _read_token()
    if result or not login_if_needed:
        return result
    print(
        "No Entra app ID. Opening Azure CLI login in your browser.\n"
        "Complete JumpCloud / company SSO there. This uses Microsoft's Azure CLI\n"
        "application, not an app you register.\n"
    )
    login = subprocess.run([az, "login"], check=False)
    if login.returncode != 0:
        return None
    return _read_token()


def _raise_if_failed(result: dict[str, Any] | None) -> dict[str, Any]:
    if not result:
        raise SharePointAuthError("Microsoft returned no token result")
    if "access_token" in result:
        return result
    error = result.get("error") or "unknown_error"
    description = result.get("error_description") or result.get("error_uri") or ""
    raise SharePointAuthError(f"OAuth failed: {error}\n{description}".rstrip())


def acquire_token_silent() -> dict[str, Any] | None:
    path = cache_path()
    cache = _load_cache(path)
    app = _app(cache)
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    _save_cache(cache, path)
    if result and "access_token" in result:
        return result
    return None


def acquire_token_interactive() -> dict[str, Any]:
    """Open the default browser. JumpCloud/SSO happens there.

    After you finish sign-in, Microsoft redirects to http://localhost with an
    authorization code. MSAL exchanges that code for tokens. The script never
    reads cookies or scrapes the login page.
    """
    path = cache_path()
    cache = _load_cache(path)
    app = _app(cache)
    print(
        "Opening your default browser for Microsoft sign-in.\n"
        "Complete JumpCloud / company SSO in the browser, then return here.\n"
        "This script does not scrape the page; Microsoft sends the token to a\n"
        "local redirect after authentication succeeds.\n"
    )
    result = app.acquire_token_interactive(
        scopes=GRAPH_SCOPES,
        prompt="select_account",
        timeout=300,
        port=8400,
    )
    _save_cache(cache, path)
    return _raise_if_failed(result)


def acquire_token_device_code() -> dict[str, Any]:
    """Fallback when localhost redirect is blocked.

    Microsoft shows a code; the browser opens the device-login page. You sign
    in through the same JumpCloud/SSO flow. MSAL then receives the token.
    """
    path = cache_path()
    cache = _load_cache(path)
    app = _app(cache)
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "user_code" not in flow:
        raise SharePointAuthError(f"Could not start device login: {flow}")
    print(flow["message"])
    print(
        "\nThis script does not scrape the page; Microsoft returns the token\n"
        "to the library after you finish sign-in.\n"
    )
    uri = flow.get("verification_uri") or flow.get("verification_uri_complete")
    if uri:
        webbrowser.open(uri)
    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache, path)
    return _raise_if_failed(result)


def get_token(*, force_login: bool = False, device_code: bool = False) -> dict[str, Any]:
    if client_id_optional():
        if not force_login:
            cached = acquire_token_silent()
            if cached:
                print("Reusing cached Microsoft token (no browser).")
                return cached
        if device_code:
            result = acquire_token_device_code()
        else:
            result = acquire_token_interactive()
        print("Authenticated successfully.")
        return result

    if not force_login:
        cached = acquire_token_azure_cli(login_if_needed=False)
        if cached:
            print("Reusing Azure CLI Graph token (no Entra app registration).")
            return cached
    cli = acquire_token_azure_cli(login_if_needed=True)
    if cli:
        print("Authenticated successfully via Azure CLI.")
        return cli
    raise SharePointAuthError(_no_graph_auth_help())
