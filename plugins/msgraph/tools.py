"""Tool schemas and handlers for read-only Microsoft Graph access."""

from __future__ import annotations

import json
from typing import Any

from gateway.platforms._shared import get_scoped_secret
from tools.microsoft_graph_auth import GraphCredentials, MicrosoftGraphTokenProvider
from tools.microsoft_graph_client import MicrosoftGraphClient


_CLIENTS: dict[tuple[str, ...], MicrosoftGraphClient] = {}


def _credentials() -> GraphCredentials | None:
    tenant_id = str(get_scoped_secret("MSGRAPH_TENANT_ID", "") or "").strip()
    client_id = str(get_scoped_secret("MSGRAPH_CLIENT_ID", "") or "").strip()
    client_secret = str(get_scoped_secret("MSGRAPH_CLIENT_SECRET", "") or "").strip()
    if not all((tenant_id, client_id, client_secret)):
        return None
    return GraphCredentials(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        scope=str(get_scoped_secret("MSGRAPH_SCOPE", "") or "").strip()
        or "https://graph.microsoft.com/.default",
        authority_url=str(get_scoped_secret("MSGRAPH_AUTHORITY_URL", "") or "").strip()
        or "https://login.microsoftonline.com",
    )


def _client() -> MicrosoftGraphClient:
    credentials = _credentials()
    if credentials is None:
        raise RuntimeError(
            "MSGRAPH_CLIENT_ID, MSGRAPH_CLIENT_SECRET, and MSGRAPH_TENANT_ID are required"
        )
    key = (credentials.tenant_id, credentials.client_id, credentials.client_secret,
           credentials.scope, credentials.authority_url)
    client = _CLIENTS.get(key)
    if client is None:
        client = MicrosoftGraphClient(MicrosoftGraphTokenProvider(credentials))
        _CLIENTS[key] = client
    return client


def check_graph_credentials() -> bool:
    return _credentials() is not None


def _default_user_id() -> str:
    return str(get_scoped_secret("MSGRAPH_DEFAULT_USER_ID", "") or "").strip()


def _user_path(args: dict[str, Any], suffix: str) -> str:
    user_id = str(args.get("user_id") or _default_user_id()).strip()
    if not user_id or any(char in user_id for char in "/?#"):
        raise ValueError("user_id is required, or MSGRAPH_DEFAULT_USER_ID must be configured, and must be a user ID or UPN")
    return f"/users/{user_id}/{suffix.lstrip('/')}"


def _resource_id(args: dict[str, Any], key: str) -> str:
    value = str(args.get(key) or "").strip()
    if not value or any(char in value for char in "/?#"):
        raise ValueError(f"{key} is required and must be an opaque Graph resource ID")
    return value


async def _json_result(operation) -> str:
    try:
        return json.dumps(await operation(), default=str)
    except Exception as exc:
        return json.dumps({"error": str(exc)})


CALENDAR_VIEW_SCHEMA = {
    "name": "msgraph_calendar_view",
    "description": "List Microsoft Graph calendar occurrences, exceptions, and single events in a time range.",
    "parameters": {"type": "object", "properties": {
        "user_id": {"type": "string", "description": "Mailbox user ID or UPN authorized for app-only access."},
        "start_date_time": {"type": "string", "description": "ISO 8601 start time, including timezone offset."},
        "end_date_time": {"type": "string", "description": "ISO 8601 end time, including timezone offset."},
        "calendar_id": {"type": "string", "description": "Optional calendar ID; defaults to the default calendar."},
        "timezone": {"type": "string", "description": "Optional Outlook timezone for returned event times."},
        "top": {"type": "integer", "description": "Optional page size from 1 to 1000."},
        "next_link": {"type": "string", "description": "Opaque @odata.nextLink from an earlier call."},
    }, "required": ["start_date_time", "end_date_time"]},
}


async def calendar_view(args: dict[str, Any], **kwargs) -> str:
    calendar_id = str(args.get("calendar_id") or "").strip()
    path = _user_path(args, f"calendars/{calendar_id}/calendarView" if calendar_id else "calendar/calendarView")
    params = {"startDateTime": args["start_date_time"], "endDateTime": args["end_date_time"]}
    if args.get("top") is not None:
        params["$top"] = args["top"]
    if args.get("next_link"):
        path, params = args["next_link"], None
    headers = {"Prefer": f"outlook.timezone=\"{args['timezone']}\""} if args.get("timezone") else None
    return await _json_result(lambda: _client().get_json(path, params=params, headers=headers))


GET_EVENT_SCHEMA = {"name": "msgraph_get_event", "description": "Get one Microsoft Graph calendar event by ID.",
                    "parameters": {"type": "object", "properties": {"user_id": {"type": "string", "description": "Optional mailbox ID or UPN; defaults to MSGRAPH_DEFAULT_USER_ID."}, "event_id": {"type": "string"}}, "required": ["event_id"]}}


async def get_event(args: dict[str, Any], **kwargs) -> str:
    return await _json_result(lambda: _client().get_json(_user_path(args, f"events/{_resource_id(args, 'event_id')}")))


LIST_MESSAGES_SCHEMA = {"name": "msgraph_list_messages", "description": "List Microsoft Graph mail messages with optional OData filters and an opaque next link.",
                        "parameters": {"type": "object", "properties": {
                            "user_id": {"type": "string"}, "folder_id": {"type": "string"}, "filter": {"type": "string"}, "search": {"type": "string"},
                            "select": {"type": "string", "description": "Comma-separated properties; defaults to sender,subject,receivedDateTime,isRead,webLink."},
                            "orderby": {"type": "string"}, "top": {"type": "integer"}, "next_link": {"type": "string"},
                        }, "required": []}}


async def list_messages(args: dict[str, Any], **kwargs) -> str:
    folder_id = str(args.get("folder_id") or "").strip()
    path = _user_path(args, f"mailFolders/{folder_id}/messages" if folder_id else "messages")
    params: dict[str, Any] = {"$select": args.get("select") or "sender,subject,receivedDateTime,isRead,webLink"}
    for key, odata_key in (("filter", "$filter"), ("search", "$search"), ("orderby", "$orderby"), ("top", "$top")):
        if args.get(key) not in (None, ""):
            params[odata_key] = args[key]
    if args.get("next_link"):
        path, params = args["next_link"], None
    return await _json_result(lambda: _client().get_json(path, params=params,
                                                           headers={"Prefer": "outlook.body-content-type=\"text\""}))


GET_MESSAGE_SCHEMA = {"name": "msgraph_get_message", "description": "Get one Microsoft Graph email message by ID, optionally including its body.",
                     "parameters": {"type": "object", "properties": {"user_id": {"type": "string", "description": "Optional mailbox ID or UPN; defaults to MSGRAPH_DEFAULT_USER_ID."}, "message_id": {"type": "string"}, "select": {"type": "string"}}, "required": ["message_id"]}}


async def get_message(args: dict[str, Any], **kwargs) -> str:
    params = {"$select": args["select"]} if args.get("select") else None
    return await _json_result(lambda: _client().get_json(_user_path(args, f"messages/{_resource_id(args, 'message_id')}"), params=params,
                                                           headers={"Prefer": "outlook.body-content-type=\"text\""}))


LIST_CONTACTS_SCHEMA = {"name": "msgraph_list_contacts", "description": "List Microsoft Graph contacts, optionally filtering by email address.",
                       "parameters": {"type": "object", "properties": {"user_id": {"type": "string", "description": "Optional mailbox ID or UPN; defaults to MSGRAPH_DEFAULT_USER_ID."}, "folder_id": {"type": "string"}, "email": {"type": "string"}, "select": {"type": "string"}, "top": {"type": "integer"}, "next_link": {"type": "string"}}, "required": []}}


async def list_contacts(args: dict[str, Any], **kwargs) -> str:
    folder_id = str(args.get("folder_id") or "").strip()
    path = _user_path(args, f"contactFolders/{folder_id}/contacts" if folder_id else "contacts")
    params: dict[str, Any] = {}
    if args.get("email"):
        escaped = str(args["email"]).replace("'", "''")
        params["$filter"] = f"emailAddresses/any(a:a/address eq '{escaped}')"
    if args.get("select"):
        params["$select"] = args["select"]
    if args.get("top") is not None:
        params["$top"] = args["top"]
    if args.get("next_link"):
        path, params = args["next_link"], None
    return await _json_result(lambda: _client().get_json(path, params=params))


GET_CONTACT_SCHEMA = {"name": "msgraph_get_contact", "description": "Get one Microsoft Graph contact by ID.",
                     "parameters": {"type": "object", "properties": {"user_id": {"type": "string", "description": "Optional mailbox ID or UPN; defaults to MSGRAPH_DEFAULT_USER_ID."}, "contact_id": {"type": "string"}, "folder_id": {"type": "string"}}, "required": ["contact_id"]}}


async def get_contact(args: dict[str, Any], **kwargs) -> str:
    folder_id = str(args.get("folder_id") or "").strip()
    path = _user_path(args, f"contactFolders/{folder_id}/contacts/{_resource_id(args, 'contact_id')}" if folder_id else f"contacts/{_resource_id(args, 'contact_id')}")
    return await _json_result(lambda: _client().get_json(path))
