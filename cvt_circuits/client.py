from __future__ import annotations

import json
import ssl
import urllib.error
import urllib.parse
import urllib.request
from http.cookiejar import CookieJar
from typing import Any, Iterable


class CvtApiError(RuntimeError):
    def __init__(self, message: str, status: int | None = None, body: str = "") -> None:
        super().__init__(message)
        self.status = status
        self.body = body


class CvtClient:
    def __init__(
        self,
        base_url: str,
        username: str,
        password: str,
        insecure: bool = True,
        timeout: int = 120,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.username = username
        self.password = password
        self.timeout = timeout
        self.cookie_jar = CookieJar()
        context = ssl._create_unverified_context() if insecure else ssl.create_default_context()
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.cookie_jar),
            urllib.request.HTTPSHandler(context=context),
        )

    def login(self) -> None:
        payload = urllib.parse.urlencode(
            {
                "httpd_username": self.username,
                "httpd_password": self.password,
            }
        ).encode()
        request = urllib.request.Request(
            f"{self.base_url}/cablevalidation/login",
            data=payload,
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        with self.opener.open(request, timeout=self.timeout) as response:
            # Apache returns 302 to /cables_validation after a good login.
            if response.status not in (200, 302) and not (300 <= response.status < 400):
                raise CvtApiError(f"login failed with HTTP {response.status}", response.status)

    def get_json(self, path: str, params: dict[str, Any] | None = None) -> Any:
        query = urllib.parse.urlencode({k: v for k, v in (params or {}).items() if v is not None})
        url = f"{self.base_url}{path}"
        if query:
            url = f"{url}?{query}"
        request = urllib.request.Request(url, method="GET")
        try:
            with self.opener.open(request, timeout=self.timeout) as response:
                raw = response.read().decode("utf-8")
                if not raw:
                    return None
                return json.loads(raw)
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise CvtApiError(f"GET {path} failed with HTTP {exc.code}: {body[:500]}", exc.code, body) from exc
        except json.JSONDecodeError as exc:
            raise CvtApiError(f"GET {path} returned invalid JSON: {exc}") from exc

    def ready(self) -> Any:
        return self.get_json("/cablevalidation/ready")

    def circuits_stats(self, context: str = "dc", items: str | None = None) -> dict[str, Any]:
        return self.get_json(
            "/cablevalidation/report/circuits/stats",
            {"context": context, "items": items},
        )

    def data_halls(self) -> list[str]:
        halls = self.get_json("/cablevalidation/resources/data_halls")
        if not isinstance(halls, list):
            raise CvtApiError(f"unexpected data_halls payload: {type(halls)}")
        return [str(item) for item in halls]

    def scalable_units(self, *, data_hall: str | None = None, context: str | None = None) -> list[str]:
        params: dict[str, Any] = {}
        if data_hall:
            params["data_hall"] = data_hall
        if context:
            params["context"] = context
        units = self.get_json("/cablevalidation/resources/scalable_units", params or None)
        if not isinstance(units, list):
            raise CvtApiError(f"unexpected scalable_units payload: {type(units)}")
        return [str(item) for item in units]

    def su_scopes(self) -> list[str]:
        """Return unique hall/SU scopes, e.g. ``EH1A/SU1``."""
        scopes: list[str] = []
        seen: set[str] = set()
        for hall in self.data_halls():
            for su in self.scalable_units(data_hall=hall):
                scope = su if "/" in su else f"{hall}/{su}"
                if scope in seen:
                    continue
                seen.add(scope)
                scopes.append(scope)
        return scopes

    def racks(self) -> list[str]:
        racks = self.get_json("/cablevalidation/resources/racks")
        if not isinstance(racks, list):
            raise CvtApiError(f"unexpected racks payload: {type(racks)}")
        return [str(item) for item in racks]

    def circuits(
        self,
        *,
        context: str,
        items: str | None = None,
        page: str = "circuit",
        healthy: bool | None = None,
        report: str | None = None,
        node: str | None = None,
    ) -> list[dict[str, Any]]:
        params: dict[str, Any] = {
            "context": context,
            "items": items,
            "page": page,
            "report": report,
            "node": node,
        }
        if healthy is not None:
            params["healthy"] = "true" if healthy else "false"
        payload = self.get_json("/cablevalidation/report/circuits", params)
        if payload is None:
            return []
        if not isinstance(payload, list):
            raise CvtApiError(f"unexpected circuits payload: {type(payload)}")
        return payload

    def iter_su_circuits(
        self,
        *,
        healthy: bool | None = False,
        report: str | None = None,
        page: str = "circuit",
    ) -> Iterable[tuple[str, list[dict[str, Any]]]]:
        """Yield circuits for each hall/SU scope.

        A single ``context=dc`` dump times out on 16K (~350k circuits). SU list
        requires ``data_hall``, so halls are used only to discover SU numbers.
        Each circuits request is ``context=su&items=<hall>/<su>``.
        """
        for scope in self.su_scopes():
            yield scope, self.circuits(
                context="su",
                items=scope,
                page=page,
                healthy=healthy,
                report=report,
            )
