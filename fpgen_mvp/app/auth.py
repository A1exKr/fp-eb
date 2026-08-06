"""Authentication/authorization via exaBase oauth2-proxy + Keycloak.

The exaBase nginx layer authenticates every request and forwards identity as
headers (``X-Forwarded-User`` / ``X-Forwarded-Email`` / ``X-Forwarded-groups``).
This module surfaces that identity and provides role-gating dependencies. When
``FPGEN_AUTH_ENABLED`` is false (local dev) a synthetic admin user is returned.
"""
from dataclasses import dataclass, field

from fastapi import Depends, HTTPException, Request, status

from app.config import settings


def _split(value: str) -> list[str]:
    """Split a comma/semicolon/space-separated header into normalized tokens.

    Keycloak groups may arrive as ``/Finance`` etc.; the leading slash is stripped.
    """
    if not value:
        return []
    normalized = value.replace(";", ",").replace("\n", ",")
    tokens: list[str] = []
    for chunk in normalized.split(","):
        for part in chunk.split():
            cleaned = part.strip().lstrip("/").strip()
            if cleaned:
                tokens.append(cleaned)
    return tokens


@dataclass
class CurrentUser:
    user: str = ""
    email: str = ""
    groups: list[str] = field(default_factory=list)

    def has_any(self, allowed: set[str]) -> bool:
        return bool(set(self.groups) & allowed)


def get_current_user(request: Request) -> CurrentUser:
    if not settings.auth_enabled:
        return CurrentUser(
            user=settings.auth_dev_user,
            email=settings.auth_dev_user,
            groups=_split(settings.auth_dev_groups),
        )

    headers = request.headers
    user = headers.get("x-forwarded-user") or headers.get("x-forwarded-email") or ""
    email = headers.get("x-forwarded-email") or ""
    groups = _split(
        headers.get("x-forwarded-groups")
        or headers.get("x-forwarded-roles")
        or ""
    )
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED, detail="Not authenticated"
        )
    return CurrentUser(user=user, email=email, groups=groups)


def require_role(*allowed_csv: str):
    """Build a dependency requiring the user to belong to any of the given roles/groups."""
    allowed: set[str] = set()
    for item in allowed_csv:
        allowed |= set(_split(item))

    def _dep(user: CurrentUser = Depends(get_current_user)) -> CurrentUser:
        if allowed and not user.has_any(allowed):
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Insufficient permissions",
            )
        return user

    return _dep


# Pre-built dependencies for common gates.
require_admin = require_role(settings.admin_roles)
require_finance = require_role(settings.finance_roles)
# Gate for the setup/admin console shell: anyone who may edit at least one
# section (platform admins, plus the Finance group for unit rates).
require_setup = require_role(settings.admin_roles, settings.finance_roles)


def user_can_setup(user: CurrentUser) -> bool:
    """True when the user may reach the setup/admin console (admin or Finance).

    Mirrors ``require_setup`` but returns a boolean so callers (e.g. ``/v1/me``)
    can advertise the capability to the UI without raising.
    """
    allowed = set(_split(settings.admin_roles)) | set(_split(settings.finance_roles))
    return user.has_any(allowed)


def user_can_regenerate(user: CurrentUser) -> bool:
    """True when the user may re-run pipeline stages on a saved proposal."""
    return user.has_any(set(_split(settings.admin_roles)))


def user_can_recalculate_fee(user: CurrentUser) -> bool:
    """True when the user may re-run the fee engine on a saved proposal."""
    return user.has_any(set(_split(settings.finance_roles)))

