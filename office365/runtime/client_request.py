from abc import abstractmethod

import httpx

from office365.runtime.client_request_exception import ClientRequestException
from office365.runtime.http.http_method import HttpMethod
from office365.runtime.http.request_options import RequestOptions
from office365.runtime.queries.client_query import ClientQuery
from office365.runtime.types.event_handler import EventHandler


class ClientRequest(object):
    def __init__(self, http_client):
        # type: (httpx.AsyncClient) -> None
        """
        Abstract request client
        """
        self._http_client = http_client
        self.beforeExecute = EventHandler()
        self.afterExecute = EventHandler()

    @abstractmethod
    def build_request(self, query):
        # type: (ClientQuery) -> RequestOptions
        """Builds a request"""
        pass

    @abstractmethod
    async def process_response(self, response, query):
        # type: (httpx.Response, ClientQuery) -> None
        pass

    async def execute_query(self, query):
        # type: (ClientQuery) -> None
        """Submits a pending request to the server"""
        try:
            request = self.build_request(query)
            response = await self.execute_request_direct(request)
            self.process_response(response, query)
            await self.afterExecute.notify(response)
        except httpx.HTTPStatusError as e:
            raise ClientRequestException(*e.args, response=e.response)

    async def execute_request_direct(self, request):
        # type: (RequestOptions) -> httpx.Response
        """Execute the client request"""
        await self.beforeExecute.notify(request)

        method = request.method
        url = request.url
        headers = request.headers
        auth = request.auth

        if request.stream:
            # For streaming responses, send without reading body eagerly
            req = self._http_client.build_request(method, url, headers=headers)
            response = await self._http_client.send(req, stream=True)
            response.raise_for_status()
            return response

        if method == HttpMethod.Post:
            if request.is_bytes or request.is_file:
                response = await self._http_client.post(
                    url=url,
                    headers=headers,
                    content=request.data,
                    auth=auth,
                )
            else:
                response = await self._http_client.post(
                    url=url,
                    headers=headers,
                    json=request.data,
                    auth=auth,
                )
        elif method == HttpMethod.Patch:
            response = await self._http_client.patch(
                url=url,
                headers=headers,
                json=request.data,
                auth=auth,
            )
        elif method == HttpMethod.Delete:
            response = await self._http_client.delete(
                url=url,
                headers=headers,
                auth=auth,
            )
        elif method == HttpMethod.Put:
            response = await self._http_client.put(
                url=url,
                content=request.data,
                headers=headers,
                auth=auth,
            )
        else:
            response = await self._http_client.get(
                url=url,
                headers=headers,
                auth=auth,
            )
        response.raise_for_status()
        return response
