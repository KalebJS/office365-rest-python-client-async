import asyncio
from random import randint

from office365.sharepoint.files.file import File
from office365.sharepoint.recyclebin.item_collection import RecycleBinItemCollection
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSharePointRecycleBin(SPTestCase):
    target_file = None  # type: File

    @classmethod
    def setUpClass(cls):
        super(TestSharePointRecycleBin, cls).setUpClass()
        file_name = "Sample{0}.txt".format(str(randint(0, 10000)))

        async def _async_setup():
            target_file = await (
                cls.client.web.default_document_library()
                .root_folder.upload_file(file_name, "--some content goes here--")
                .execute_query()
            )
            cls.target_file = target_file

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_recycle_file(self):
        result = await self.__class__.target_file.recycle().execute_query()
        self.assertIsNotNone(result.value)

    async def test2_find_removed_file(self):
        file_name = self.__class__.target_file.name
        items = await self.client.site.recycle_bin.filter("LeafName eq '{0}'".format(file_name)).get().execute_query()
        self.assertGreater(len(items), 0)

    async def test3_restore_file(self):
        items = await self.client.web.recycle_bin.get().execute_query()
        self.assertGreater(len(items), 0)
        await items[0].restore().execute_query()
        items_after = await self.client.web.recycle_bin.get().execute_query()
        self.assertEqual(len(items_after), len(items) - 1)

    async def test4_get_site_recycle_bin_items(self):
        items = await self.client.site.get_recycle_bin_items().execute_query()
        self.assertIsInstance(items, RecycleBinItemCollection)

    async def test5_get_web_recycle_bin_items(self):
        items = await self.client.web.get_recycle_bin_items().execute_query()
        self.assertIsInstance(items, RecycleBinItemCollection)

    async def test6_clear_recycle_bin(self):
        await self.client.site.recycle_bin.delete_all().execute_query()
        items_after = await self.client.site.recycle_bin.get().execute_query()
        self.assertEqual(len(items_after), 0)
