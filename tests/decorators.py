import asyncio
from functools import lru_cache, wraps
from typing import Any, Callable, TypeVar
from unittest import IsolatedAsyncioTestCase

from office365.directory.applications.roles.collection import AppRoleCollection
from office365.graph_client import GraphClient
from office365.runtime.types.collections import StringCollection
from tests import test_client_id

T = TypeVar("T", bound=Callable[..., Any])


async def _fetch_permissions(client, client_id):
    # type: (GraphClient, str) -> AppRoleCollection
    resource = client.service_principals.get_by_name("Microsoft Graph")
    result = await resource.get_application_permissions(client_id).execute_query()
    return result.value


async def _fetch_delegated_permissions(client, client_id):
    # type: (GraphClient, str) -> StringCollection
    resource = client.service_principals.get_by_name("Microsoft Graph")
    result = await resource.get_delegated_permissions(client_id).execute_query()
    return result.value


def requires_app_permission(*app_roles):
    # type: (*str) -> Callable[[T], T]
    def decorator(test_method):
        # type: (T) -> T
        @wraps(test_method)
        async def wrapper(self, *args, **kwargs):
            # type: (IsolatedAsyncioTestCase, *Any, **Any) -> Any
            client = getattr(self, "client", None)
            if not client:
                self.skipTest("No client available for permission check")

            permissions = await _fetch_permissions(client, test_client_id)

            if not any(role.value in app_roles for role in permissions):
                required_roles = ", ".join(f"'{role}'" for role in app_roles)
                self.skipTest(f"Required app permission '{required_roles}' not granted")

            return await test_method(self, *args, **kwargs)

        return wrapper

    return decorator


def requires_delegated_permission(*scopes):
    # type: (*str) -> Callable[[T], T]
    """Decorator to verify delegated permissions before test execution"""

    def decorator(test_method):
        # type: (T) -> T
        @wraps(test_method)
        async def wrapper(self, *args, **kwargs):
            # type: (IsolatedAsyncioTestCase, *Any, **Any) -> Any
            client = getattr(self, "client", None)
            if not client:
                self.skipTest("No client available for permission check")

            granted_scopes = await _fetch_delegated_permissions(client, test_client_id)

            if not any(scope in granted_scopes for scope in scopes):
                self.skipTest(f"Required delegated permission '{', '.join(scopes)}' not granted")

            return await test_method(self, *args, **kwargs)

        return wrapper  # type: ignore[return-value]

    return decorator
