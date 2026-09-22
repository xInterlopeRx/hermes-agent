# Microsoft Graph Connector Progress

Last updated: 2026-09-22

## Current status

The connector is implemented as a native Hermes plugin and is also installable
as a standalone user plugin under `~/.hermes/plugins/msgraph/`.

The local plugin is enabled and has been validated through Hermes plugin
discovery and Plugin Doctor.

## Implemented

- Microsoft Entra app-only client-credentials authentication.
- Profile-scoped secret lookup for Graph credentials.
- Access-token caching and refresh.
- HTTP timeout and retry handling for transport errors, `401`, `429`, and `5xx`.
- `Retry-After` handling.
- OData pagination through opaque `@odata.nextLink` values.
- Resource-ID path validation.
- Six read-only tools:
  - `msgraph_calendar_view`
  - `msgraph_get_event`
  - `msgraph_list_messages`
  - `msgraph_get_message`
  - `msgraph_list_contacts`
  - `msgraph_get_contact`
- Optional `MSGRAPH_DEFAULT_USER_ID` profile setting. When configured, tools
  use that mailbox when `user_id` is omitted; an explicit `user_id` overrides
  it.

## Local installation

The standalone plugin is installed at:

```text
~/.hermes/plugins/msgraph/
```

The source copy is:

```text
plugins/msgraph/
```

Required profile secrets:

```text
MSGRAPH_TENANT_ID
MSGRAPH_CLIENT_ID
MSGRAPH_CLIENT_SECRET
```

Optional profile settings:

```text
MSGRAPH_DEFAULT_USER_ID
MSGRAPH_SCOPE
MSGRAPH_AUTHORITY_URL
```

Secrets remain in the local Hermes profile and are not part of the plugin
repository.

## Validation completed

The following local tests pass:

```bash
scripts/run_tests.sh \
  tests/plugins/test_msgraph_plugin.py \
  tests/tools/test_microsoft_graph_auth.py \
  tests/tools/test_microsoft_graph_client.py
```

Current result: 15 tests passed. The Plugin Doctor also confirms six tools are
registered through the standalone user-plugin path.

Live validation against the configured mailbox confirmed:

- Calendar view request succeeds; the tested seven-day window returned zero events.
- Mail listing succeeds and returned five message metadata records with an
  available next page.
- Default-mailbox operation works without passing `user_id` in the tool call.
- Token acquisition and caching succeed.

The live contacts request returned `403 ErrorAccessDenied`, indicating that the
Entra app still needs the appropriate Microsoft Graph application Contacts
permission and admin consent.

## Teams status

The shared Graph authentication/client foundation can support Teams Graph
endpoints, but the `msgraph` connector does not currently expose Teams tools.

Existing Hermes Teams-related functionality is separate:

- `teams` platform: Bot Framework messaging.
- `msgraph_webhook`: Graph change notifications.
- Teams meeting pipeline: transcript/recording processing.
- Graph-based Teams summary delivery: existing pipeline output path.

Potential future read-only connector tools include chat, team, channel, and
channel-message listing. Sending Teams messages should be a separate explicit
write toolset with confirmation and narrowly scoped application permissions.

## Next steps

1. Grant and validate the required Contacts application permission if contact
   access is needed.
2. Decide whether Teams access should be read-only, write-capable, or both.
3. Add Teams read tools as a separate opt-in toolset.
4. Add confirmation-gated Teams write tools only after the read surface is
   validated.
5. Publish the standalone plugin repository and optionally add it to the
   Hermes plugin catalog with a pinned commit.
