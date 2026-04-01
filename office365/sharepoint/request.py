from typing import Callable, List, Optional

import httpx
from typing_extensions import Self

from office365.azure_env import AzureEnvironment
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.token_response import TokenResponse
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.http.request_options import RequestOptions
from office365.runtime.odata.request import ODataRequest
from office365.runtime.odata.v3.json_light_format import JsonLightFormat


class SharePointRequest(ODataRequest):
    def __init__(
        self,
        base_url,
        http_client,
        environment=AzureEnvironment.Global,
        allow_ntlm=False,
        browser_mode=False,
    ):
        """
        :param str base_url: Absolute Web or Site Url
        :param httpx.AsyncClient http_client: Shared async HTTP client
        :param str environment: The Office 365 Cloud Environment endpoint used for authentication
        :param bool allow_ntlm: Flag indicates whether NTLM scheme is enabled. Disabled by default
        :param bool browser_mode: Allow browser authentication
        """
        super().__init__(JsonLightFormat(), http_client)
        self._auth_context = AuthenticationContext(
            url=base_url,
            environment=environment,
            allow_ntlm=allow_ntlm,
            browser_mode=browser_mode,
        )
        self.beforeExecute += self._authenticate_request

    async def execute_request(self, path):
        # type: (str) -> httpx.Response
        request_url = "{0}/{1}".format(self.service_root_url, path)
        return await self.execute_request_direct(RequestOptions(request_url))

    def with_credentials(self, credentials):
        # type: (UserCredential|ClientCredential) -> Self
        """
        Initializes a client to acquire a token via user or client credentials
        """
        self._auth_context.with_credentials(credentials)
        return self

    def with_cookies(self, cookie_source, ttl_seconds=None):
        # type: (object, object) -> Self
        """
        Initializes authentication using browser-session cookies.

        :param object cookie_source: Callable returning Dict[str, str] or an AuthCookies instance.
        :param object ttl_seconds: Optional max age for cached cookies before reloading from source.
        """
        self._auth_context.with_cookies(cookie_source, ttl_seconds)
        return self

    def with_client_certificate(
        self,
        tenant,
        client_id,
        thumbprint,
        cert_path=None,
        private_key=None,
        scopes=None,
        passphrase=None,
    ):
        # type: (str, str, str, Optional[str], Optional[str], Optional[List[str]], Optional[str]) -> Self
        """
        Creates authenticated SharePoint context via certificate credentials
        """
        self._auth_context.with_client_certificate(
            tenant, client_id, thumbprint, cert_path, private_key, scopes, passphrase
        )
        return self

    def with_device_flow(self, tenant, client_id, scopes=None):
        # type: (str, str, Optional[List[str]]) -> Self
        """
        Initializes a client to acquire a token via device flow auth.
        """
        self._auth_context.with_device_flow(tenant, client_id, scopes)
        return self

    def with_interactive(self, tenant, client_id, scopes=None):
        # type: (str, str, Optional[List[str]]) -> Self
        """
        Initializes a client to acquire a token interactively i.e. via a local browser.
        """
        self._auth_context.with_interactive(tenant, client_id, scopes)
        return self

    def with_access_token(self, token_func):
        # type: (Callable[[], TokenResponse]) -> Self
        """
        Initializes a client to acquire a token from a callback
        """
        self._auth_context.with_access_token(token_func)
        return self

    async def _authenticate_request(self, request):
        # type: (RequestOptions) -> None
        """Authenticate request"""
        await self._auth_context.authenticate_request(request)

    @property
    def authentication_context(self):
        return self._auth_context

    @property
    def base_url(self):
        """Represents Base Url"""
        return self._auth_context.url

    @property
    def service_root_url(self):
        return "{0}/_api".format(self._auth_context.url)
