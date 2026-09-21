"""Read-only Microsoft Graph tools for calendar, mail, and contacts."""

from __future__ import annotations

from plugins.msgraph import tools


_TOOLS = (
    ("msgraph_calendar_view", tools.CALENDAR_VIEW_SCHEMA, tools.calendar_view, "calendar"),
    ("msgraph_get_event", tools.GET_EVENT_SCHEMA, tools.get_event, "calendar"),
    ("msgraph_list_messages", tools.LIST_MESSAGES_SCHEMA, tools.list_messages, "mail"),
    ("msgraph_get_message", tools.GET_MESSAGE_SCHEMA, tools.get_message, "mail"),
    ("msgraph_list_contacts", tools.LIST_CONTACTS_SCHEMA, tools.list_contacts, "contacts"),
    ("msgraph_get_contact", tools.GET_CONTACT_SCHEMA, tools.get_contact, "contacts"),
)


def register(ctx) -> None:
    """Register the opt-in Microsoft Graph toolset."""
    for name, schema, handler, emoji in _TOOLS:
        ctx.register_tool(
            name=name,
            toolset="msgraph",
            schema=schema,
            handler=handler,
            check_fn=tools.check_graph_credentials,
            is_async=True,
            emoji=emoji,
        )