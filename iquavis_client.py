import os
import json
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
import urllib3


# Disable SSL warnings to match ref samples (internal network usage)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


DEFAULT_BASE_URL = "http://rdgpm0701/iquavis-api"


class IQuavisClient:
    """
    Lightweight API client for iQUAVIS, based on ref samples.
    - Auth via password grant to /token
    - Basic GET/POST helpers with Bearer token
    - Project and Task retrieval helpers
    """

    def __init__(self, base_url: Optional[str] = None, timeout: int = 30) -> None:
        self.base_url = base_url or os.getenv("IQUAVIS_BASE_URL", DEFAULT_BASE_URL)
        self.timeout = timeout
        self.session = requests.Session()
        # In trusted internal environments, certificate validation is disabled per samples
        self.session.verify = False
        self.access_token: Optional[str] = None

    # -------------------- Core HTTP --------------------
    def _auth_header(self) -> Dict[str, str]:
        if not self.access_token:
            return {}
        return {"Authorization": f"Bearer {self.access_token}"}

    def _get(self, path: str, params: Optional[Dict[str, Any]] = None) -> Any:
        url = f"{self.base_url}{path}"
        headers = {"Content-Type": "application/json", **self._auth_header()}
        eff_params = dict(params or {})
        # Increase default page size as done in refs
        eff_params.setdefault("count", 10000)
        r = self.session.get(url, headers=headers, params=eff_params, timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    def _post(self, path: str, json_body: Any, params: Optional[Dict[str, Any]] = None) -> Any:
        url = f"{self.base_url}{path}"
        headers = {"Content-Type": "application/json", **self._auth_header()}
        r = self.session.post(url, headers=headers, json=json_body, params=params or {}, timeout=self.timeout)
        if r.status_code in (200, 201):
            try:
                return r.json()
            except ValueError:
                return {"status": r.status_code}
        r.raise_for_status()
        return r.json()

    # -------------------- Auth --------------------
    def login(self, user_id: str, password: str) -> str:
        """
        Authenticate via password grant and store access token.
        Returns the access token on success; raises on failure.
        """
        url = f"{self.base_url}/token"
        payload = {"grant_type": "password", "username": user_id, "password": password}
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        r = self.session.post(url, headers=headers, data=payload, timeout=self.timeout)
        r.raise_for_status()
        token = r.json().get("access_token")
        if not token:
            raise RuntimeError("No access_token in response")
        self.access_token = token
        return token

    # -------------------- Projects --------------------
    def list_projects(self, name: Optional[str] = None) -> List[Dict[str, Any]]:
        params: Dict[str, Any] = {}
        if name:
            params["name"] = name
        items = self._get("/v1/projects", params=params)
        if not isinstance(items, list):
            # Some endpoints may return non-list; normalize
            return []
        return items

    # -------------------- Tasks --------------------
    def list_tasks(
        self,
        project_id: str,
        name: Optional[str] = None,
        include: Optional[Iterable[str]] = None,
        count: Optional[int] = 10000,
    ) -> List[Dict[str, Any]]:
        params: Dict[str, Any] = {}
        if name:
            params["name"] = name
        if include:
            # Many endpoints expect comma-separated includes rather than repeated keys
            # Convert iterable -> comma-separated string to maximize compatibility.
            try:
                include_list = list(include)
                params["include"] = ",".join(include_list)
            except TypeError:
                # If a plain string is passed, keep as-is
                params["include"] = include  # type: ignore[assignment]
        if count is not None:
            params["count"] = count
        items = self._get(f"/v1/projects/{project_id}/tasks", params=params)
        if not isinstance(items, list):
            return []
        return items

    # Utility to safely get canonical Id/Name from project objects
    @staticmethod
    def project_identity(p: Dict[str, Any]) -> Tuple[str, str]:
        pid = str(p.get("Id") or p.get("id") or p.get("ID") or "")
        name = str(p.get("Name") or p.get("name") or "")
        return pid, name

    # Utility to unwrap a task possibly wrapped in {"Task": {...}}
    @staticmethod
    def unwrap_task(item: Dict[str, Any]) -> Dict[str, Any]:
        if isinstance(item, dict) and "Task" in item and isinstance(item["Task"], dict):
            return item["Task"]
        return item
