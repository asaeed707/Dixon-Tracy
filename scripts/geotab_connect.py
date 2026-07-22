#!/usr/bin/env python3
"""Minimal MyGeotab JSON-RPC connector.

Required environment variables:
  GEOTAB_USERNAME
  GEOTAB_PASSWORD
  GEOTAB_DATABASE

Optional:
  GEOTAB_SERVER=my.geotab.com
"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any


def require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


class GeotabClient:
    def __init__(self, server: str = "my.geotab.com", timeout: int = 60) -> None:
        self.scheme, self.server = self._normalize_server(server)
        self.timeout = timeout
        self.credentials: dict[str, Any] | None = None

    @staticmethod
    def _normalize_server(server: str) -> tuple[str, str]:
        value = server.strip()
        parsed = urllib.parse.urlparse(value if "://" in value else f"https://{value}")
        scheme = parsed.scheme or "https"
        host = (parsed.netloc or parsed.path).removesuffix("/")
        return scheme, host

    @property
    def endpoint(self) -> str:
        return f"{self.scheme}://{self.server}/apiv1"

    def call(self, method: str, params: dict[str, Any]) -> Any:
        payload = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": method,
            "params": params,
        }
        request = urllib.request.Request(
            self.endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        for attempt in range(3):
            try:
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    body = json.loads(response.read().decode("utf-8"))
                break
            except urllib.error.HTTPError as exc:
                detail = exc.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"HTTP {exc.code} from {self.endpoint}: {detail}") from exc
            except TimeoutError:
                if attempt == 2:
                    raise
                time.sleep(2 * (attempt + 1))
            except urllib.error.URLError as exc:
                if attempt == 2 or not isinstance(exc.reason, TimeoutError):
                    raise
                time.sleep(2 * (attempt + 1))

        if "error" in body:
            raise RuntimeError(json.dumps(body["error"], indent=2))
        return body["result"]

    def authenticate(self, username: str, password: str, database: str) -> dict[str, Any]:
        result = self.call(
            "Authenticate",
            {
                "userName": username,
                "password": password,
                "database": database,
            },
        )

        path = result.get("path")
        if path and path != "ThisServer":
            self.scheme, self.server = self._normalize_server(path)

        self.credentials = result["credentials"]
        return self.credentials

    def get_devices(self, limit: int = 1) -> list[dict[str, Any]]:
        if not self.credentials:
            raise RuntimeError("Call authenticate() before get_devices().")

        return self.call(
            "Get",
            {
                "typeName": "Device",
                "credentials": self.credentials,
                "resultsLimit": limit,
                "propertySelector": {
                    "fields": ["id", "name"],
                    "isIncluded": True,
                },
            },
        )

    def get_driver_logs(
        self,
        limit: int = 10,
        from_date: str | None = None,
        to_date: str | None = None,
    ) -> list[dict[str, Any]]:
        if not self.credentials:
            raise RuntimeError("Call authenticate() before get_driver_logs().")

        search: dict[str, Any] = {}
        if from_date:
            search["fromDate"] = from_date
        if to_date:
            search["toDate"] = to_date

        params: dict[str, Any] = {
            "typeName": "DutyStatusLog",
            "credentials": self.credentials,
            "resultsLimit": limit,
        }
        if search:
            params["search"] = search

        return self.call("Get", params)


def main() -> int:
    client = GeotabClient(
        os.environ.get("GEOTAB_SERVER", "my.geotab.com"),
        timeout=int(os.environ.get("GEOTAB_TIMEOUT", "60")),
    )
    credentials = client.authenticate(
        require_env("GEOTAB_USERNAME"),
        require_env("GEOTAB_PASSWORD"),
        require_env("GEOTAB_DATABASE"),
    )
    devices = client.get_devices(limit=1)

    print("Connected to MyGeotab.")
    print(f"Server: {client.scheme}://{client.server}")
    print(f"Database: {credentials['database']}")
    print(f"Sample devices returned: {len(devices)}")
    if devices:
        print(json.dumps(devices[0], indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
