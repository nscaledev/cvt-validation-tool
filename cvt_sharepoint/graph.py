from __future__ import annotations

import os
from typing import Any
from urllib.parse import quote

import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_HOSTNAME = "dell.sharepoint.com"
DEFAULT_SITE_PATH = "/sites/NSCALE-WARDCTYTX-Phase1"
DEFAULT_FILE_ID = "4FB77D68-CC31-491D-B23E-6F7DFC720561"


class GraphError(RuntimeError):
    def __init__(self, message: str, status: int | None = None, body: str = "") -> None:
        super().__init__(message)
        self.status = status
        self.body = body


def _env(name: str, default: str) -> str:
    return (os.environ.get(name) or default).strip()


def site_hostname() -> str:
    return _env("SHAREPOINT_HOSTNAME", DEFAULT_HOSTNAME)


def site_path() -> str:
    path = _env("SHAREPOINT_SITE_PATH", DEFAULT_SITE_PATH)
    return path if path.startswith("/") else f"/{path}"


def file_id() -> str:
    return _env("SHAREPOINT_FILE_ID", DEFAULT_FILE_ID).strip("{}")


def _headers(access_token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def graph_request(
    access_token: str,
    method: str,
    path: str,
    json_body: dict[str, Any] | None = None,
) -> Any:
    url = f"{GRAPH_ROOT}{path}"
    response = requests.request(
        method,
        url,
        headers=_headers(access_token),
        json=json_body,
        timeout=60,
    )
    if not response.ok:
        raise GraphError(
            f"{method} {path} failed with HTTP {response.status_code}: {response.text[:800]}",
            response.status_code,
            response.text,
        )
    if not response.content:
        return None
    try:
        return response.json()
    except ValueError as exc:
        raise GraphError(f"{method} {path} returned invalid JSON: {exc}") from exc


def graph_get(access_token: str, path: str) -> Any:
    return graph_request(access_token, "GET", path)


def me(access_token: str) -> dict[str, Any]:
    payload = graph_get(access_token, "/me")
    if not isinstance(payload, dict):
        raise GraphError(f"unexpected /me payload: {type(payload)}")
    return payload


def site(access_token: str) -> dict[str, Any]:
    host = quote(site_hostname(), safe="")
    relative = quote(site_path(), safe="/")
    payload = graph_get(access_token, f"/sites/{host}:{relative}")
    if not isinstance(payload, dict):
        raise GraphError(f"unexpected site payload: {type(payload)}")
    return payload


def drive_item(access_token: str, site_id: str, item_id: str | None = None) -> dict[str, Any]:
    item = item_id or file_id()
    payload = graph_get(access_token, f"/sites/{site_id}/drive/items/{item}")
    if not isinstance(payload, dict):
        raise GraphError(f"unexpected drive item payload: {type(payload)}")
    return payload


def resolve_workbook(access_token: str) -> tuple[str, dict[str, Any], dict[str, Any]]:
    site_payload = site(access_token)
    site_id = site_payload.get("id")
    if not site_id:
        raise GraphError("site lookup returned no id")
    item = drive_item(access_token, site_id, file_id())
    return str(site_id), site_payload, item


def worksheets(access_token: str, site_id: str, item_id: str) -> list[dict[str, Any]]:
    payload = graph_get(
        access_token,
        f"/sites/{site_id}/drive/items/{item_id}/workbook/worksheets",
    )
    if not isinstance(payload, dict) or not isinstance(payload.get("value"), list):
        raise GraphError(f"unexpected worksheets payload: {type(payload)}")
    return payload["value"]


def _worksheet_path(name: str) -> str:
    escaped = name.replace("'", "''")
    return f"worksheets('{escaped}')"


def range_path(site_id: str, item_id: str, worksheet: str, address: str) -> str:
    return (
        f"/sites/{site_id}/drive/items/{item_id}/workbook/"
        f"{_worksheet_path(worksheet)}/range(address='{address}')"
    )


def update_range(
    access_token: str,
    site_id: str,
    item_id: str,
    worksheet: str,
    address: str,
    values: list[list[Any]],
) -> dict[str, Any]:
    payload = graph_request(
        access_token,
        "PATCH",
        range_path(site_id, item_id, worksheet, address),
        {"values": values},
    )
    if not isinstance(payload, dict):
        raise GraphError(f"unexpected range update payload: {type(payload)}")
    return payload
