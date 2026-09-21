"""Behavior tests for the Microsoft Graph connector plugin."""

from __future__ import annotations

import json

import pytest

from plugins.msgraph import register
from plugins.msgraph import tools


class _ToolContext:
    def __init__(self):
        self.registered = []

    def register_tool(self, **kwargs):
        self.registered.append(kwargs)


class _FakeGraphClient:
    def __init__(self, payload=None):
        self.payload = payload or {"value": []}
        self.calls = []

    async def get_json(self, path, *, params=None, headers=None):
        self.calls.append((path, params, headers))
        return self.payload


@pytest.fixture
def graph_credentials(monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": "tenant-from-bitwarden",
        "MSGRAPH_CLIENT_ID": "client-from-bitwarden",
        "MSGRAPH_CLIENT_SECRET": "secret-from-bitwarden",
    }
    monkeypatch.setattr(tools, "get_scoped_secret", lambda name, default=None: values.get(name, default))
    tools._CLIENTS.clear()
    return values


def test_registers_async_tools_with_credential_gate():
    context = _ToolContext()

    register(context)

    assert {entry["name"] for entry in context.registered} == {
        "msgraph_calendar_view",
        "msgraph_get_event",
        "msgraph_list_messages",
        "msgraph_get_message",
        "msgraph_list_contacts",
        "msgraph_get_contact",
    }
    assert all(entry["toolset"] == "msgraph" for entry in context.registered)
    assert all(entry["is_async"] for entry in context.registered)
    assert all(entry["check_fn"] is tools.check_graph_credentials for entry in context.registered)


def test_credentials_are_read_from_active_scoped_secret_mapping(graph_credentials):
    assert tools.check_graph_credentials()
    credentials = tools._credentials()
    assert credentials is not None
    assert credentials.tenant_id == "tenant-from-bitwarden"
    assert credentials.client_id == "client-from-bitwarden"
    assert credentials.client_secret == "secret-from-bitwarden"


@pytest.mark.anyio
async def test_calendar_view_targets_explicit_mailbox_and_preserves_timezone(graph_credentials, monkeypatch):
    client = _FakeGraphClient({"value": [{"subject": "Planning"}]})
    monkeypatch.setattr(tools, "_client", lambda: client)

    result = await tools.calendar_view({
        "user_id": "person@example.com",
        "start_date_time": "2026-09-21T09:00:00-04:00",
        "end_date_time": "2026-09-21T17:00:00-04:00",
        "timezone": "Eastern Standard Time",
    })

    assert json.loads(result) == {"value": [{"subject": "Planning"}]}
    assert client.calls == [(
        "/users/person@example.com/calendar/calendarView",
        {"startDateTime": "2026-09-21T09:00:00-04:00", "endDateTime": "2026-09-21T17:00:00-04:00"},
        {"Prefer": 'outlook.timezone="Eastern Standard Time"'},
    )]


@pytest.mark.anyio
async def test_message_pagination_uses_opaque_next_link(graph_credentials, monkeypatch):
    client = _FakeGraphClient({"value": [{"subject": "Next page"}]})
    monkeypatch.setattr(tools, "_client", lambda: client)

    result = await tools.list_messages({
        "user_id": "person@example.com",
        "next_link": "https://graph.microsoft.com/v1.0/users/person@example.com/messages?$skip=10",
    })

    assert json.loads(result)["value"][0]["subject"] == "Next page"
    assert client.calls == [(
        "https://graph.microsoft.com/v1.0/users/person@example.com/messages?$skip=10",
        None,
        {"Prefer": 'outlook.body-content-type="text"'},
    )]


@pytest.mark.anyio
async def test_contacts_escape_email_filter_and_reject_path_injection(graph_credentials, monkeypatch):
    client = _FakeGraphClient()
    monkeypatch.setattr(tools, "_client", lambda: client)

    await tools.list_contacts({"user_id": "person@example.com", "email": "o'neal@example.com"})
    rejected = await tools.get_event({"user_id": "person@example.com", "event_id": "bad/id"})

    assert client.calls[0][0] == "/users/person@example.com/contacts"
    assert client.calls[0][1] == {"$filter": "emailAddresses/any(a:a/address eq 'o''neal@example.com')"}
    assert "opaque Graph resource ID" in rejected
