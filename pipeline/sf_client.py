"""Minimal Salesforce REST client.

Requests go through `sf api request rest`, so the Salesforce CLI session
(`sf org login web`) signs them and the access token never touches this code.
If SF_INSTANCE_URL / SF_ACCESS_TOKEN are set, calls go direct over HTTPS instead.
No passwords live in this repo.
"""
import json
import os
import subprocess
import urllib.parse
import urllib.request

API_VERSION = "v62.0"


class SalesforceClient:
    def __init__(self, target_org=None):
        self.target_org = target_org
        self.token = os.environ.get("SF_ACCESS_TOKEN")
        self.instance_url = os.environ.get("SF_INSTANCE_URL") if self.token else self._instance_url()

    def _instance_url(self):
        cmd = ["sf", "org", "display", "--json"] + (["--target-org", self.target_org] if self.target_org else [])
        out = subprocess.run(cmd, capture_output=True, text=True, check=True).stdout
        return json.loads(out)["result"]["instanceUrl"]

    def _request(self, method, path, body=None):
        if path.startswith("http"):
            path = urllib.parse.urlsplit(path)._replace(scheme="", netloc="").geturl()
        if self.token:
            return self._direct(method, path, body)

        cmd = ["sf", "api", "request", "rest", path, "--method", method]
        if self.target_org:
            cmd += ["--target-org", self.target_org]
        if body is not None:
            cmd += ["--body", "-"]  # read the JSON body from stdin
        if body is not None:
            proc = subprocess.run(cmd, capture_output=True, text=True, input=json.dumps(body))
        else:
            proc = subprocess.run(cmd, capture_output=True, text=True, stdin=subprocess.DEVNULL)
        out = proc.stdout.strip()
        if proc.returncode != 0:
            raise RuntimeError(f"{method} {path} failed: {(out or proc.stderr).strip()[:500]}")
        parsed = json.loads(out) if out else None
        # The REST API reports errors as a list of {errorCode, message}.
        if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict) and "errorCode" in parsed[0]:
            raise RuntimeError(f"{method} {path} failed: {parsed[0]['errorCode']}: {parsed[0]['message']}")
        return parsed

    def _direct(self, method, path, body):
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(f"{self.instance_url}{path}", data=data, method=method, headers={
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        })
        with urllib.request.urlopen(req) as resp:
            raw = resp.read()
            return json.loads(raw) if raw else None

    def delete_ids(self, sobject, ids):
        """Delete via the Composite API (POST), 25 sub-requests per call."""
        for i in range(0, len(ids), 25):
            body = {"allOrNone": False, "compositeRequest": [
                {"method": "DELETE", "url": f"/services/data/{API_VERSION}/sobjects/{sobject}/{rid}",
                 "referenceId": f"d{n}"} for n, rid in enumerate(ids[i:i + 25])]}
            result = self._request("POST", f"/services/data/{API_VERSION}/composite", body)
            failed = [r for r in result["compositeResponse"] if r["httpStatusCode"] >= 300]
            if failed:
                raise RuntimeError(f"{len(failed)} deletes failed, first: {failed[0]['body']}")

    def query_all(self, soql):
        """Run SOQL and follow nextRecordsUrl pagination."""
        path = f"/services/data/{API_VERSION}/query?q={urllib.parse.quote(soql)}"
        records = []
        while path:
            page = self._request("GET", path)
            records.extend(page["records"])
            path = page.get("nextRecordsUrl")
        return records
