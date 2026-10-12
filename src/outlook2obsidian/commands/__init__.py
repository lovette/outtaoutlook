"""outlook2obsidian commands."""

from .auth import app as auth_app
from .export import app as export_app
from .ls import app as ls_app

__all__ = [
    "auth_app",
    "export_app",
    "ls_app",
]
