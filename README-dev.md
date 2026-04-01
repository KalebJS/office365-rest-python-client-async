# Installing to virtualenv

```bash
uv sync
```

## Running tests

Most of the tests are end-to-end — operations are invoked against an actual tenant (not mocked).
Configure credentials before running them.

### 1. Create a `.env` file (never committed — it's in `.gitignore`)

```bash
export office365_python_sdk_securevars='{username};{password};{client_id};{client_secret}'
```

Indices in the semicolon-separated string:
- `[1]` — password (user credentials)
- `[3]` — client secret (app credentials)

### 2. Source and run

```bash
. .env
pytest tests/directory/test_user.py -v        # single file
pytest tests/ -v                               # full suite
```

Tests use `pytest-asyncio` with `asyncio_mode = "auto"` (configured in `pyproject.toml`),
so no per-test `@pytest.mark.asyncio` decorators are needed. All test methods are
`async def` and all `execute_query()` / `execute_query_retry()` calls are `await`ed.

## Configure Tenant

Required roles:

- Global reader
- Groups admin
- Search admin
- SharePoint admin
- Teams service admin
- Users admin
