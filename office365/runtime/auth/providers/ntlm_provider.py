from office365.runtime.auth.authentication_provider import AuthenticationProvider


class NtlmProvider(AuthenticationProvider):
    def __init__(self, username, password):
        """
        NTLM authentication is not supported in the async client.
        """
        super(NtlmProvider, self).__init__()

    async def authenticate_request(self, request):
        raise NotImplementedError(
            "NTLM authentication is not supported in the async client. "
            "NTLM is only available for SharePoint On-Premises which requires the "
            "synchronous requests-ntlm library."
        )
