"""Saved Views — persist named filter/preset snapshots across sessions.

Views are stored in saved_views.json next to this file.  Each view is a dict
of session-state key bases (version tag stripped) → serialised values.
Stripping the version tag means saved views survive app version bumps.
"""

import json
import datetime
from pathlib import Path

VIEWS_FILE = Path(__file__).parent / "saved_views.json"


# ── Serialisation helpers ─────────────────────────────────────────────────────

def _serialize(v):
    if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
        return {"__type": "date", "value": v.isoformat()}
    if isinstance(v, datetime.datetime):
        return {"__type": "datetime", "value": v.isoformat()}
    return v


def _deserialize(v):
    if isinstance(v, dict):
        t = v.get("__type")
        if t == "date":
            return datetime.date.fromisoformat(v["value"])
        if t == "datetime":
            return datetime.datetime.fromisoformat(v["value"])
    return v


# ── Storage ───────────────────────────────────────────────────────────────────

def load_all() -> dict:
    """Return {name: {key_base: value}} for all saved views."""
    if not VIEWS_FILE.exists():
        return {}
    try:
        raw = json.loads(VIEWS_FILE.read_text(encoding="utf-8"))
        return {
            name: {k: _deserialize(v) for k, v in view.items()}
            for name, view in raw.items()
        }
    except Exception:
        return {}


def _persist(views: dict):
    def _enc(v):
        s = _serialize(v)
        if isinstance(s, dict):
            return s
        # fallback for any remaining non-serialisable types
        try:
            json.dumps(s)
            return s
        except TypeError:
            return str(s)

    serialised = {
        name: {k: _enc(val) for k, val in view.items()}
        for name, view in views.items()
    }
    VIEWS_FILE.write_text(
        json.dumps(serialised, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def save_view(name: str, state: dict):
    views = load_all()
    views[name] = state
    _persist(views)


def delete_view(name: str):
    views = load_all()
    views.pop(name, None)
    _persist(views)


def list_names() -> list[str]:
    return list(load_all().keys())
