import json
import os
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


def _default_log_path() -> str:
    env_path = os.environ.get("EXCEL_BOT_LOG_PATH")
    if env_path:
        return env_path
    return os.path.join(os.getcwd(), "logs", "events.jsonl")


DEFAULT_LOG_PATH = _default_log_path()


EVENTS: List[Dict[str, Any]] = []


def _append_jsonl(path: str, record: Dict[str, Any]) -> None:
    directory = os.path.dirname(path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record) + "\n")


def emit_event(
    event_type: str,
    user_id: str,
    payload: Optional[Dict[str, Any]] = None,
    level: str = "INFO",
    log_path: Optional[str] = None,
) -> Dict[str, Any]:
    event = {
        "id": str(uuid.uuid4()),
        "type": event_type,
        "user_id": user_id,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "level": level,
        "payload": payload or {},
    }
    EVENTS.append(event)
    if log_path is None:
        log_path = _default_log_path()
    if log_path:
        _append_jsonl(log_path, event)
    return event


def load_events(path: Optional[str] = None) -> List[Dict[str, Any]]:
    if path is None:
        path = _default_log_path()
    events: List[Dict[str, Any]] = []
    if not os.path.exists(path):
        return events
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            events.append(json.loads(line))
    return events
