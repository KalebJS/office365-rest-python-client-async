
# About
Microsoft 365 & Microsoft Graph library for Python — **async edition**

This is an async-only fork of [office365-rest-python-client](https://github.com/vgrem/office365-rest-python-client) that replaces the `requests`-based transport with [`httpx`](https://www.python-httpx.org/) and exposes a fully `async`/`await` API. Python 3.9+ is required.

# Usage

1. [Installation](#Installation)
2. [Working with SharePoint API](#Working-with-SharePoint-API)
3. [Working with Outlook API](#Working-with-Outlook-API)
4. [Working with OneDrive and SharePoint API v2 APIs](#working-with-onedrive-and-sharepoint-v2-apis)
5. [Working with Teams API](#Working-with-Microsoft-Teams-API)
6. [Working with OneNote API](#Working-with-Microsoft-OneNote-API)
7. [Working with Planner API](#Working-with-Microsoft-Planner-API)

## Status
[![PyPI](https://img.shields.io/pypi/v/office365-rest-python-client.svg)](https://pypi.python.org/pypi/office365-rest-python-client)
[![PyPI pyversions](https://img.shields.io/pypi/pyversions/office365-rest-python-client.svg)](https://pypi.python.org/pypi/office365-rest-python-client/)

# Installation

Use pip:

```
pip install office365-rest-python-client
```

### Note
>
>Alternatively the _latest_ version could be directly installed via GitHub:
>```
>pip install git+https://github.com/vgrem/office365-rest-python-client.git
>```

# Authentication
For the following examples, relevant credentials can be found in the Azure Portal.

Steps to access:
1. Login to the home page of the Azure Portal
2. Navigate to "Azure Active Directory" using the three bars in the top right corner of the portal
3. Select "App registrations" in the navigation panel on the left
4. Search for and select your relevant application
5. In the application's "Overview" page, the client id can be found under "Application (client) id"
6. In the application's "Certificates & Secrets" page, the client secret can be found under the "Value" of the "Client Secrets." If there is no client secret yet, create one here.


# Working with SharePoint API

   The `ClientContext` client provides support for the legacy SharePoint REST and OneDrive for Business REST APIs,
   including:
   -   [SharePoint 2013 REST API](https://msdn.microsoft.com/en-us/library/office/jj860569.aspx) and above
   -   [SharePoint Online REST API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
   -   OneDrive for Business REST API

### Authentication

   The following auth flows are supported:

#### 1. Using a SharePoint App-Only principal (client credentials flow)

   - `ClientContext.with_credentials(client_credentials)`
   - `ClientContext.with_client_credentials(client_id, client_secret)`

   Usage:
   ```python
   import asyncio
   from office365.sharepoint.client_context import ClientContext
   from office365.runtime.auth.client_credential import ClientCredential

   async def main():
       client_credentials = ClientCredential('{client_id}', '{client_secret}')
       async with ClientContext('{url}').with_credentials(client_credentials) as ctx:
           web = await ctx.web.get().execute_query()
           print(web.title)

   asyncio.run(main())
   ```

   Documentation:
   - [Granting access using SharePoint App-Only](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs)

   Example: [connect_with_app_principal.py](examples/sharepoint/auth/with_app_only.py)

#### 2. Using username and password

   Usage:
   ```python
   import asyncio
   from office365.sharepoint.client_context import ClientContext
   from office365.runtime.auth.user_credential import UserCredential

   async def main():
       user_credentials = UserCredential('{username}', '{password}')
       async with ClientContext('{url}').with_credentials(user_credentials) as ctx:
           web = await ctx.web.get().execute_query()
           print(web.title)

   asyncio.run(main())
   ```

   Example: [connect_with_user_credential.py](examples/sharepoint/auth/with_user_credential.py)

#### 3. Using an Azure AD application (certificate credentials flow)

  Documentation:
   - [Granting access via Azure AD App-Only](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread)

  Example: [with_certificate.py](examples/sharepoint/auth/with_certificate.py)

#### 4. Interactive

   Login interactively via a local browser.

   Prerequisite:
   > In Azure Portal, configure the Redirect URI of your
   "Mobile and Desktop application" as `http://localhost`.

  Example: [connect_interactive.py](examples/sharepoint/auth/with_interactive.py)

  Usage:
```python
import asyncio
from office365.sharepoint.client_context import ClientContext

async def main():
    async with ClientContext("{site-url}").with_interactive("{tenant-name-or-id}", "{client-id}") as ctx:
        me = await ctx.web.current_user.get().execute_query()
        print(me.login_name)

asyncio.run(main())
```

#### 5. Browser session cookies (SharePoint Online)

Authenticate using cookies from a real browser session (e.g., Playwright). No Azure AD app registration required.

Usage:
```python
import asyncio
from office365.sharepoint.client_context import ClientContext

def cookie_source():
    return {"FedAuth": "...", "rtFa": "...", "SPOIDCRL": "..."}

async def main():
    async with ClientContext("https://contoso.sharepoint.com/sites/demo").with_cookies(cookie_source) as ctx:
        web = await ctx.web.get().execute_query()
        print(web.title)

asyncio.run(main())
```

Example: [auth_cookies.py](examples/sharepoint/auth_cookies.py)

### Examples

```python
import asyncio
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

async def main():
    site_url = "https://{your-tenant-prefix}.sharepoint.com"
    async with ClientContext(site_url).with_credentials(UserCredential("{username}", "{password}")) as ctx:
        web = await ctx.web.get().execute_query()
        print("Web title: {0}".format(web.title))

asyncio.run(main())
```

### Support for Azure environments

  To enable authentication to specific Azure environment endpoints, pass the `environment` parameter:

   ```python
   import asyncio
   from office365.azure_env import AzureEnvironment
   from office365.sharepoint.client_context import ClientContext
   from office365.runtime.auth.client_credential import ClientCredential

   async def main():
       client_credentials = ClientCredential('{client_id}', '{client_secret}')
       async with ClientContext(
           '{site-url}',
           environment=AzureEnvironment.USGovernmentHigh
       ).with_credentials(client_credentials) as ctx:
           web = await ctx.web.get().execute_query()

   asyncio.run(main())
   ```

# Working with Outlook API

The list of supported APIs:
-   [Outlook Contacts REST API](https://msdn.microsoft.com/en-us/office/office365/api/contacts-rest-operations)
-   [Outlook Calendar REST API](https://msdn.microsoft.com/en-us/office/office365/api/calendar-rest-operations)
-   [Outlook Mail REST API](https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations)

Since Outlook REST APIs are available in both Microsoft Graph and the Outlook API endpoint,
use `GraphClient` which targets the Microsoft Graph `v1.0` endpoint.

### Authentication

[The Microsoft Authentication Library (MSAL) for Python](https://pypi.org/project/msal/) is used as the default library to obtain tokens. MSAL calls are automatically run in a thread pool so they don't block the event loop.

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(tenant='{tenant_name_or_id}').with_client_secret(
        client_id='{client_id}',
        client_secret='{client_secret}'
    ) as client:
        me = await client.me.get().execute_query()
        print(me.display_name)

asyncio.run(main())
```

Example: [with_client_secret](examples/auth/with_client_secret.py)

Custom token acquisition using any OAuth-compliant library is also supported:

```python
import asyncio
import adal
from office365.graph_client import GraphClient

def acquire_token_func():
    authority_url = 'https://login.microsoftonline.com/{tenant_id_or_name}'
    auth_ctx = adal.AuthenticationContext(authority_url)
    token = auth_ctx.acquire_token_with_client_credentials(
        "https://graph.microsoft.com",
        "{client_id}",
        "{client_secret}")
    return token

async def main():
    async with GraphClient(acquire_token_func) as client:
        me = await client.me.get().execute_query()
        print(me.display_name)

asyncio.run(main())
```

#### Example: send an email

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(tenant='{tenant_name_or_id}').with_username_and_password(
        '{client_id}', '{username}', '{password}'
    ) as client:
        await client.me.send_mail(
            subject="Meet for lunch?",
            body="The new cafeteria is open.",
            to_recipients=["fannyd@contoso.onmicrosoft.com"]
        ).execute_query()

asyncio.run(main())
```

Additional examples & scenarios:

-  [download a message](examples/outlook/messages/download.py)
-  [list messages](examples/outlook/messages/list_all.py)
-  [move messages to a different folder](examples/outlook/messages/move.py)
-  [search messages](examples/outlook/messages/search.py)
-  [send messages](examples/outlook/messages/send.py)
-  [send messages with attachments](examples/outlook/messages/send_with_attachment.py)

Refer to [examples section](examples/outlook) for other scenarios


# Working with OneDrive and SharePoint v2 APIs

#### Documentation

[OneDrive Graph API reference](https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)

#### Authentication

[The Microsoft Authentication Library (MSAL) for Python](https://pypi.org/project/msal/) is used to obtain tokens.

```python
import asyncio
import msal
from office365.graph_client import GraphClient

def acquire_token_func():
    authority_url = 'https://login.microsoftonline.com/{tenant_id_or_name}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id='{client_id}',
        client_credential='{client_secret}'
    )
    return app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
```

#### Examples

##### Example: list available drives

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(acquire_token_func) as client:
        drives = await client.drives.get().execute_query()
        async for drive in drives:
            print("Drive url: {0}".format(drive.web_url))

asyncio.run(main())
```

##### Example: download the contents of a DriveItem (folder)

```python
import asyncio
import os
import tempfile
from office365.graph_client import GraphClient

async def download_files(remote_folder, local_path):
    drive_items = await remote_folder.children.get().execute_query()
    async for drive_item in drive_items:
        if drive_item.file is not None:
            with open(os.path.join(local_path, drive_item.name), 'wb') as local_file:
                await drive_item.download(local_file).execute_query()

async def main():
    async with GraphClient(acquire_token_func) as client:
        drive = await client.users["{user_id_or_principal_name}"].drive.get().execute_query()
        with tempfile.TemporaryDirectory() as path:
            await download_files(drive.root, path)

asyncio.run(main())
```

Additional examples:

-  [create list column](examples/onedrive/columns/create_text.py)
-  [download file](examples/onedrive/files/download.py)
-  [export files](examples/onedrive/files/export.py)
-  [upload folder](examples/onedrive/folders/upload.py)
-  [list drives](examples/onedrive/drives/list.py)
-  [list files](examples/onedrive/folders/list_files.py)

Refer to [OneDrive examples section](examples/onedrive) for more examples.


# Working with Microsoft Teams API

#### Examples

##### Example: create a new team under a group

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(acquire_token_func) as client:
        new_team = await client.groups["{group-id}"].add_team().execute_query_retry()

asyncio.run(main())
```

Additional examples:

-  [create a team](examples/teams/create_team.py)
-  [create team from group](examples/teams/create_from_group.py)
-  [list all teams](examples/teams/list_all.py)
-  [list my teams](examples/teams/list_my_teams.py)
-  [send messages](examples/teams/send_message.py)

Refer to [examples section](examples/teams) for other scenarios

# Working with Microsoft Onenote API

The library supports OneNote API calls to a user's notebooks, sections, and pages.

Example: Create a new page

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(tenant='{tenant_name_or_id}').with_username_and_password(
        '{client_id}', '{username}', '{password}'
    ) as client:
        files = {}
        with open("./MyPage.html", 'rb') as f, \
            open("./MyImage.png", 'rb') as img_f, \
            open("./MyDoc.pdf", 'rb') as pdf_f:
            files["imageBlock1"] = img_f
            files["fileBlock1"] = pdf_f
            page = await client.me.onenote.pages.add(
                presentation_file=f, attachment_files=files
            ).execute_query()

asyncio.run(main())
```

# Working with Microsoft Planner API

Example: create a new planner task

```python
import asyncio
from office365.graph_client import GraphClient

async def main():
    async with GraphClient(acquire_token_func) as client:
        task = await client.planner.tasks.add(
            title="New task", planId="--plan id goes here--"
        ).execute_query()

asyncio.run(main())
```

# Third Party Libraries and Dependencies
The following libraries will be installed when you install the client library:
* [httpx](https://www.python-httpx.org/) — async HTTP transport
* [Microsoft Authentication Library (MSAL) for Python](https://pypi.org/project/msal/)


# ThanksTo

Powerful Python IDE [`Pycharm`](https://www.jetbrains.com/pycharm) from [`Jetbrains`](https://jb.gg/OpenSourceSupport).

[<img src="https://resources.jetbrains.com/storage/products/company/brand/logos/jb_beam.svg">](https://jb.gg/OpenSourceSupport)
