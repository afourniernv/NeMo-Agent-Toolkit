# SPDX-FileCopyrightText: Copyright (c) 2025-2026, NVIDIA CORPORATION & AFFILIATES.
# All rights reserved.
# SPDX-License-Identifier: Apache-2.0

"""Per-turn A365 identity propagated from front-end handlers to telemetry exporter.

The Microsoft Agents SDK exposes the active agent's app id and tenant on the
inbound `Activity` (`get_agentic_instance_id()` / `get_agentic_tenant_id()`).
The A365 backend validates that the agent_id in the export URL matches the
`appid` claim in the bearer token, so telemetry must be partitioned by the
same identity that authenticated the turn — not by static config.

This module owns the contextvar that ties the two together. The front-end
worker writes it for the lifetime of a turn; the telemetry exporter reads it
when stamping span attributes and when looking up the cached token.
"""

from __future__ import annotations

import contextlib
import logging
from collections.abc import Iterator
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Any

logger = logging.getLogger(__name__)

_AGENTIC_ROLES = frozenset({"agenticIdentity", "agenticUser"})


@dataclass(frozen=True, slots=True)
class A365TurnIdentity:
    """Identity of the agent serving the current turn."""

    agent_app_id: str
    tenant_id: str | None
    on_behalf_user_id: str | None


_TURN_IDENTITY: ContextVar[A365TurnIdentity | None] = ContextVar(
    "a365_turn_identity", default=None
)


def get_turn_identity() -> A365TurnIdentity | None:
    """Return the identity of the current turn, or None if not in a turn."""
    return _TURN_IDENTITY.get()


@contextlib.contextmanager
def set_turn_identity(identity: A365TurnIdentity | None) -> Iterator[None]:
    """Bind ``identity`` for the duration of the with-block, restoring on exit."""
    token = _TURN_IDENTITY.set(identity)
    try:
        yield
    finally:
        _TURN_IDENTITY.reset(token)


def _call_if_callable(obj: Any, name: str) -> Any:
    fn = getattr(obj, name, None)
    if callable(fn):
        try:
            return fn()
        except Exception:
            # Method exists but threw — surface it so silent fallback to static config is observable.
            logger.warning(
                "A365 turn-identity extraction: %s.%s() raised; falling back to None",
                type(obj).__name__,
                name,
                exc_info=True,
            )
            return None
    return None


def _is_agentic_via_role(activity: Any) -> bool:
    recipient = getattr(activity, "recipient", None)
    role = getattr(recipient, "role", None)
    return role in _AGENTIC_ROLES


def extract_identity_from_activity(activity: Any) -> A365TurnIdentity | None:
    """Pull A365 identity from a Microsoft Agents Activity (or duck-typed shape).

    Returns None when the activity is not an agentic request, or when no agent
    app id can be determined. Agent app id is required because the A365 export
    URL depends on it.
    """
    if activity is None:
        return None

    is_agentic_method = getattr(activity, "is_agentic_request", None)
    if callable(is_agentic_method):
        try:
            agentic = bool(is_agentic_method())
        except Exception:
            logger.warning(
                "A365 turn-identity extraction: %s.is_agentic_request() raised; treating as non-agentic",
                type(activity).__name__,
                exc_info=True,
            )
            agentic = False
    else:
        agentic = _is_agentic_via_role(activity)

    if not agentic:
        return None

    app_id = _call_if_callable(activity, "get_agentic_instance_id")
    if app_id is None:
        recipient = getattr(activity, "recipient", None)
        app_id = getattr(recipient, "agentic_app_id", None)

    if not app_id:
        return None

    tenant_id = _call_if_callable(activity, "get_agentic_tenant_id")
    if tenant_id is None:
        recipient = getattr(activity, "recipient", None)
        tenant_id = getattr(recipient, "tenant_id", None)
        if tenant_id is None:
            conversation = getattr(activity, "conversation", None)
            tenant_id = getattr(conversation, "tenant_id", None)

    user_id = _call_if_callable(activity, "get_agentic_user")
    if user_id is None:
        recipient = getattr(activity, "recipient", None)
        user_id = getattr(recipient, "agentic_user_id", None)

    return A365TurnIdentity(
        agent_app_id=str(app_id),
        tenant_id=str(tenant_id) if tenant_id is not None else None,
        on_behalf_user_id=str(user_id) if user_id is not None else None,
    )
