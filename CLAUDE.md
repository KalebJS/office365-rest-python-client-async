# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

An **async-only** fork of `office365-rest-python-client`. The entire public API is `async`/`await`-based, using `httpx.AsyncClient` as the HTTP transport instead of `requests`. Python 3.9+ required.

## Commands

```bash
# Install dependencies
uv sync

# Lint and format (run before every commit)
prek run --all-files

# Run a single integration test (requires credentials — see below)
. .env
pytest tests/directory/test_user.py::TestGraphUser::test2_list_users -v

# Run a full test module
pytest tests/sharepoint/test_list.py -v
```

### Test credentials

Integration tests hit a real Microsoft 365 tenant. Create a `.env` file (gitignored):

```bash
export office365_python_sdk_securevars='{username};{password};{client_id};{client_secret}'
```

Indices in the semicolon-separated string: `[1]` = password, `[3]` = client secret.
Tenant/URLs are configured in `tests/settings.cfg`.

Tests use `pytest-asyncio` with `asyncio_mode = "auto"` (set in `pyproject.toml`) — no `@pytest.mark.asyncio` decorator needed.

## Architecture

### Two top-level clients

| Client | File | API surface |
|--------|------|-------------|
| `GraphClient` | `office365/graph_client.py` | Microsoft Graph (users, teams, OneDrive, Outlook, etc.) |
| `ClientContext` | `office365/sharepoint/client_context.py` | SharePoint REST + OneDrive for Business |

Both are used as `async with` context managers (they own an `httpx.AsyncClient` instance).

### Execution pipeline

Every API call follows the same path regardless of which client or resource is used:

1. **Domain object** (e.g. `User`, `Drive`) calls a method that creates a **Query** (e.g. `ReadEntityQuery`, `CreateEntityQuery`, `ServiceOperationQuery`) and registers it on the context via `context.add_query(qry)`.
2. Calling `await obj.execute_query()` (on `ClientObject`, `ClientResult`, or `ClientQuery`) delegates to `await context.execute_query()`.
3. `ClientRuntimeContext.execute_query()` (`office365/runtime/client_runtime_context.py`) fires the pending query: calls `before_execute` event handlers, then `await pending_request().execute_query(query)`.
4. `ClientRequest.execute_query()` (`office365/runtime/client_request.py`) fires `beforeExecute` (auth injects the token here), dispatches via `httpx`, fires `afterExecute`.
5. `ODataRequest.process_response()` / batch equivalents deserialize the response back into the domain object.

### Event handler pattern

`EventHandler` (`office365/runtime/types/event_handler.py`) is a simple pub/sub used throughout. `notify()` is `async` and supports both sync and async listeners (detected via `asyncio.iscoroutinefunction`). Auth providers, header injection, and response callbacks are all wired as event listeners.

### Authentication

`AuthenticationProvider` (`office365/runtime/auth/authentication_provider.py`) is the abstract base — `async def authenticate_request(request)` is the single method all providers must implement.

- **Graph**: `GraphRequest` holds an `AuthenticationContext` (Entra/MSAL); token acquisition runs via `asyncio.to_thread` since MSAL is synchronous.
- **SharePoint**: `AuthenticationContext` (`office365/runtime/auth/authentication_context.py`) supports multiple providers (ACS, SAML, cookie, certificate). Each provider is an `httpx`-based async implementation.
- **NTLM**: raises `NotImplementedError` — incompatible with async transport.

### OData layer

`ODataRequest` (`office365/runtime/odata/request.py`) subclasses `ClientRequest` and adds OData-specific request building and response parsing. There are two OData dialects:

- **v3** (`office365/runtime/odata/v3/`) — SharePoint REST (JSON Light format)
- **v4** (`office365/runtime/odata/v4/`) — Microsoft Graph (JSON format)

Batch requests (`ODataBatchV3Request`, `ODataV4BatchRequest`) construct a multipart body, fire a single HTTP call, then fan out sub-responses by constructing `httpx.Response(status_code=..., headers=..., content=...)` objects for each part.

### Domain objects

All domain objects (under `office365/directory/`, `office365/sharepoint/`, `office365/onedrive/`, etc.) inherit from `ClientObject` or `ClientValue`. They do no I/O themselves — they build queries and register them. The runtime layer executes everything.

`ClientObjectCollection` supports `async for` iteration (`__aiter__`/`__anext__`) with automatic paging.

### Test structure

- `tests/graph_case.py` — `GraphTestCase(IsolatedAsyncioTestCase)` base for all Graph tests
- `tests/sharepoint/sharepoint_case.py` — `SPTestCase(IsolatedAsyncioTestCase)` base for SharePoint tests
- `setUpClass`/`tearDownClass` that need `await` wrap their body in `async def _async_setup(): ... ; asyncio.run(_async_setup())`
- All test methods are `async def`; all `execute_query()` calls are `await`ed
