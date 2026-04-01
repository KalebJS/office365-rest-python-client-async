import asyncio

from office365.directory.applications.application import Application
from office365.directory.password_credential import PasswordCredential
from office365.directory.serviceprincipals.service_principal import ServicePrincipal
from tests import create_unique_name
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestServicePrincipal(GraphTestCase):
    target_object = None  # type: ServicePrincipal
    target_app = None  # type: Application
    password_creds = None  # type: PasswordCredential

    @classmethod
    def setUpClass(cls):
        super(TestServicePrincipal, cls).setUpClass()
        app_name = create_unique_name("App")

        async def _async_setup():
            cls.target_app = await cls.client.applications.add(app_name).execute_query()

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_app.delete_object(True).execute_query()

        asyncio.run(_async_teardown())

    async def test1_create_service_principal(self):
        service_principal = await self.client.service_principals.add(
            self.target_app.app_id
        ).execute_query()
        self.assertIsNotNone(service_principal.resource_path)
        self.__class__.target_object = service_principal

    @requires_delegated_permission(
        "Application.Read.All",
        "Application.ReadWrite.All",
        "Directory.Read.All",
        "Directory.ReadWrite.All",
    )
    async def test2_list_service_principals(self):
        result = await self.client.service_principals.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test3_get_service_principals_count(self):
        result = await self.client.service_principals.count().execute_query()
        self.assertIsNotNone(result.value)

    async def test4_get_by_app_id(self):
        principal = await (
            self.client.service_principals.get_by_app_id(self.target_app.app_id)
            .get()
            .execute_query()
        )
        self.assertIsNotNone(principal.resource_path)

    async def test5_add_password(self):
        result = await self.__class__.target_object.add_password(
            "Password friendly name"
        ).execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.password_creds = result.value

    async def test6_remove_password(self):
        key_id = self.__class__.password_creds.keyId
        await self.__class__.target_object.remove_password(key_id).execute_query()

    async def test7_delete_service_principal(self):
        await self.__class__.target_object.delete_object().execute_query()

    async def test8_list_deleted(self):
        result = (
            await self.__class__.client.directory.deleted_service_principals.get().execute_query()
        )
        self.assertIsNotNone(result.resource_path)
        self.assertGreater(len(result), 0)
