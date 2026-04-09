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
        if "version" not in self._data:
            self._data["version"] = 1
        if "namespaces" not in self._data or not isinstance(self._data.get("namespaces"), dict):
            self._data["namespaces"] = {}
        if self.namespace not in self._data["namespaces"]:
            self._data["namespaces"][self.namespace] = {"rows": {}}

    @staticmethod
    def make_key(workbook_path, job_name, row_idx):
        return "%s::%s::%s" % (workbook_path, job_name, row_idx)

    def _rows(self):
        ns = self._data["namespaces"].setdefault(self.namespace, {"rows": {}})
        rows = ns.setdefault("rows", {})
        if not isinstance(rows, dict):
            ns["rows"] = {}
            rows = ns["rows"]
        return rows

    def reset_namespace(self):
        self._data["namespaces"][self.namespace] = {"rows": {}}
        self.flush()

    def get(self, workbook_path, job_name, row_idx):
        if not self.enabled:
            return None
        key = self.make_key(workbook_path, job_name, row_idx)
        value = self._rows().get(key)
        return value if isinstance(value, dict) else None

    def mark_success(self, *, workbook_path, job_name, row_idx, collection, main_id, guid):
        if not self.enabled:
            return
        key = self.make_key(workbook_path, job_name, row_idx)
        self._rows()[key] = {
            "workbook": workbook_path,
            "job": job_name,
            "row": row_idx,
            "collection": collection,
            "_id": main_id,
            "guid": guid,
            "updatedAt": _utc_now_iso(),
        }
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
