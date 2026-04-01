import asyncio

from office365.directory.applications.application import Application
from office365.directory.extensions.extension_property import ExtensionProperty
from tests import create_unique_name
from tests.graph_case import GraphTestCase


class TestExtensions(GraphTestCase):
    target_app = None  # type: Application
    target_extension = None  # type: ExtensionProperty

    @classmethod
    def setUpClass(cls):
        super(TestExtensions, cls).setUpClass()
        app_name = create_unique_name("App")

        async def _async_setup():
            cls.target_app = await cls.client.applications.add(app_name).execute_query()

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_app.delete_object(True).execute_query()

        asyncio.run(_async_teardown())

    async def test1_create_extension(self):
        new_extension = await self.__class__.target_app.extension_properties.add(
            name="extensionName"
        ).execute_query()
        self.assertIsNotNone(new_extension.resource_path)
        self.__class__.target_extension = new_extension

    async def test2_list_extensions(self):
        extensions = (
            await self.client.directory_objects.get_available_extension_properties(
                False
            ).execute_query()
        )
        self.assertIsNotNone(extensions.resource_path)

    async def test3_delete_extension(self):
        await self.__class__.target_extension.delete_object().execute_query()
