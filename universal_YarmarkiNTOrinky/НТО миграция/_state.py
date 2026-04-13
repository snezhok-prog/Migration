import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


def _utc_now_iso():
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def _load_json(path):
    if not path.exists():
        return {}
    try:
        raw = path.read_text(encoding="utf-8")
        return json.loads(raw) if raw.strip() else {}
    except Exception:
        return {}


def _save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass
class ResumeState:
    path: Path
    namespace: str
    enabled: bool = True

    def __post_init__(self):
        self._data = _load_json(self.path)
        version_raw = self._data.get("version")
        try:
            version_num = int(version_raw)
        except Exception:
            version_num = 0
        if version_num < 2:
            self._data["version"] = 2
        if "namespaces" not in self._data or not isinstance(self._data.get("namespaces"), dict):
            self._data["namespaces"] = {}
        if self.namespace not in self._data["namespaces"]:
            self._data["namespaces"][self.namespace] = {"rows": {}, "run": {}}

    @staticmethod
    def make_key(workbook_path, job_name, row_idx):
        return "%s::%s::%s" % (workbook_path, job_name, row_idx)

    def _namespace_data(self):
        ns = self._data["namespaces"].setdefault(self.namespace, {"rows": {}, "run": {}})
        if not isinstance(ns, dict):
            ns = {"rows": {}, "run": {}}
            self._data["namespaces"][self.namespace] = ns
        rows = ns.get("rows")
        if not isinstance(rows, dict):
            ns["rows"] = {}
        run = ns.get("run")
        if not isinstance(run, dict):
            ns["run"] = {}
        return ns

    def _rows(self):
        return self._namespace_data()["rows"]

    def _run(self):
        return self._namespace_data()["run"]

    def reset_namespace(self):
        self._data["namespaces"][self.namespace] = {"rows": {}, "run": {}}
        self.flush()

    def clear_rows(self):
        self._namespace_data()["rows"] = {}
        self.flush()

    def rows_count(self):
        return len(self._rows())

    def get_run_info(self):
        run = self._run()
        return dict(run) if isinstance(run, dict) else {}

    def begin_run(self, **meta):
        if not self.enabled:
            return
        run = {
            "status": "running",
            "startedAt": _utc_now_iso(),
            "finishedAt": None,
        }
        for key, value in (meta or {}).items():
            if value is not None:
                run[key] = value
        self._namespace_data()["run"] = run
        self.flush()

    def update_run(self, **fields):
        if not self.enabled:
            return
        run = self._run()
        for key, value in (fields or {}).items():
            if value is not None:
                run[key] = value
        run["updatedAt"] = _utc_now_iso()
        self.flush()

    def finish_run(self, *, status, summary=None, clear_rows=False):
        if not self.enabled:
            return
        ns = self._namespace_data()
        run = ns.setdefault("run", {})
        run["status"] = status
        run["finishedAt"] = _utc_now_iso()
        if isinstance(summary, dict):
            run["summary"] = summary
        run["rowsCountAtFinish"] = len(ns.get("rows") or {})
        if clear_rows:
            ns["rows"] = {}
            run["rowsClearedAt"] = _utc_now_iso()
        self.flush()

    def get(self, workbook_path, job_name, row_idx):
        if not self.enabled:
            return None
        key = self.make_key(workbook_path, job_name, row_idx)
        value = self._rows().get(key)
        return value if isinstance(value, dict) else None

    def mark_success(
        self,
        *,
        workbook_path,
        job_name,
        row_idx,
        collection,
        main_id,
        guid,
        had_errors=False,
        error_count=0,
    ):
        if not self.enabled:
            return
        key = self.make_key(workbook_path, job_name, row_idx)
        updated_at = _utc_now_iso()
        row_payload = {
            "workbook": workbook_path,
            "job": job_name,
            "row": row_idx,
            "collection": collection,
            "_id": main_id,
            "guid": guid,
            "hadErrors": bool(had_errors),
            "errorCount": int(error_count or 0),
            "updatedAt": updated_at,
        }
        rows = self._rows()
        rows[key] = row_payload

        run = self._run()
        run["lastCheckpointAt"] = updated_at
        run["lastCheckpoint"] = {
            "workbook": workbook_path,
            "job": job_name,
            "row": row_idx,
            "collection": collection,
            "_id": main_id,
            "guid": guid,
            "hadErrors": bool(had_errors),
            "errorCount": int(error_count or 0),
        }
        run["rowsCount"] = len(rows)
        self.flush()

    def clear_row(self, workbook_path, job_name, row_idx):
        if not self.enabled:
            return
        key = self.make_key(workbook_path, job_name, row_idx)
        rows = self._rows()
        if key in rows:
            del rows[key]
            self.flush()

    def flush(self):
        _save_json(self.path, self._data)
