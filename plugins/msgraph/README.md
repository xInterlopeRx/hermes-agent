# Hermes Microsoft Graph plugin

Read-only Microsoft Graph access for Hermes. The plugin exposes calendar,
mail, and contact tools using Microsoft Entra app-only client-credentials
authentication.

## Install as a standalone plugin

Copy this directory into the active Hermes profile:

```bash
mkdir -p ~/.hermes/plugins
cp -a plugins/msgraph ~/.hermes/plugins/msgraph
hermes plugins enable msgraph
```

The directory must contain `plugin.yaml` and `__init__.py` directly under the
plugin directory. Restart Hermes after installing or enabling it.

The plugin uses the Graph authentication/client helpers shipped by the Hermes
version it runs with. Use a Hermes build that includes:

```text
tools.microsoft_graph_auth
tools.microsoft_graph_client
```

## Required credentials

Configure these in the active Hermes profile's secret scope:

```text
MSGRAPH_TENANT_ID
MSGRAPH_CLIENT_ID
MSGRAPH_CLIENT_SECRET
MSGRAPH_DEFAULT_USER_ID=<mailbox-upn-or-user-id>
```

Optional overrides are supported for sovereign clouds or custom token scopes:

```text
MSGRAPH_SCOPE
MSGRAPH_AUTHORITY_URL
```

`MSGRAPH_DEFAULT_USER_ID` is optional. When configured, tools use that mailbox
by default, so normal calls do not need to include `user_id`. An explicit
`user_id` still overrides it for a one-off request.

Do not put secrets in `plugin.yaml` or commit them to the plugin repository.

## Tools

- `msgraph_calendar_view`
- `msgraph_get_event`
- `msgraph_list_messages`
- `msgraph_get_message`
- `msgraph_list_contacts`
- `msgraph_get_contact`

Tools use `MSGRAPH_DEFAULT_USER_ID` when configured, or accept an explicit
`user_id` mailbox ID or UPN for one-off requests. They do not use `/me`, because
the plugin uses app-only authentication. Apply an Exchange or Microsoft Graph
application access policy to restrict which mailboxes the app can read.

See the repository guide at
`website/docs/guides/microsoft-graph-app-registration.md` for Entra setup,
permissions, admin consent, Bitwarden Secrets Manager, and token validation.

See [PROGRESS.md](./PROGRESS.md) for the current implementation status, live
validation results, known permission limitations, and planned Teams support.

## Local validation

From the Hermes checkout:

```bash
scripts/run_tests.sh \
  tests/plugins/test_msgraph_plugin.py \
  tests/tools/test_microsoft_graph_auth.py \
  tests/tools/test_microsoft_graph_client.py
```
