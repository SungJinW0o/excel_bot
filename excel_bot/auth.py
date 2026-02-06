import json
import os
from datetime import datetime
from importlib import resources
from typing import Dict, Optional


ROLE_PERMISSIONS = {
    "admin": {"run_pipeline", "view_reports", "export_reports", "modify_config", "view_logs"},
    "analyst": {"run_pipeline", "view_reports", "export_reports"},
    "viewer": {"view_reports"},
}


class User:
    def __init__(self, id, email=None, role="viewer", status="active", created_at=None):
        self.id = id
        self.email = email
        self.role = role
        self.status = status
        self.created_at = created_at or datetime.utcnow().isoformat()

    def can(self, action: str) -> bool:
        if self.status != "active":
            return False
        allowed = ROLE_PERMISSIONS.get(self.role, set())
        return action in allowed

    def __repr__(self):
        return f"<User {self.email} ({self.role}, {self.status})>"


def _load_users_from_path(users_file: str) -> list:
    if not os.path.exists(users_file):
        raise FileNotFoundError(f"users.json not found at {users_file}")
    with open(users_file, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_users_from_package() -> list:
    with resources.open_text(__package__, "users.json", encoding="utf-8") as f:
        return json.load(f)


def load_users(users_file: Optional[str] = None) -> Dict[str, User]:
    resolved_path = users_file or os.environ.get("EXCEL_BOT_USERS")
    if resolved_path:
        data = _load_users_from_path(resolved_path)
    else:
        cwd_path = os.path.join(os.getcwd(), "users.json")
        if os.path.exists(cwd_path):
            data = _load_users_from_path(cwd_path)
        else:
            data = _load_users_from_package()

    users: Dict[str, User] = {}
    for u in data:
        user = User(
            id=u["id"],
            email=u.get("email"),
            role=u.get("role", "viewer"),
            status=u.get("status", "active"),
            created_at=u.get("created_at"),
        )
        users[user.email] = user
    return users


def authorize(user: User, action: str) -> bool:
    if not user.can(action):
        raise PermissionError(
            f"User {user.email} with role {user.role} cannot perform '{action}'"
        )
    return True


def get_user(email: str, users_file: Optional[str] = None) -> User:
    users = load_users(users_file)
    if email not in users:
        raise KeyError(f"User {email} not found")
    return users[email]


def load_user(email: str, users_file: Optional[str] = None) -> User:
    return get_user(email, users_file)


if __name__ == "__main__":
    user = get_user("analyst1@example.com")
    authorize(user, "run_pipeline")
    print(f"{user} authorized to run_pipeline")
