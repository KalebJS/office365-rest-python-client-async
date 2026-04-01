import asyncio
import os
from io import BytesIO

from office365.sharepoint.changes.query import ChangeQuery
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from tests import create_unique_name, test_client_credentials
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSharePointFile(SPTestCase):
    folder_from = None  # type: Folder
    folder_to = None  # type: Folder
    file = None  # type: File
    deleted_file_guid = None
    text_content = b"updated content goes here..."

    @classmethod
    def setUpClass(cls):
        super(TestSharePointFile, cls).setUpClass()
        cls.folder_from = cls.client.web.default_document_library().root_folder.add(
            create_unique_name("from")
        )
        cls.folder_to = cls.client.web.default_document_library().root_folder.add(
            create_unique_name("to")
        )

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.folder_from.delete_object().execute_query()
            await cls.folder_to.delete_object().execute_query()

        asyncio.run(_async_teardown())

    async def test1_upload_file_as_content(self):
        path = "{0}/../data/Sample.txt".format(os.path.dirname(__file__))
        uploaded_file = await self.folder_from.files.upload(path).execute_query()
        self.assertEqual(uploaded_file.name, os.path.basename(path))
        self.assertIsNotNone(uploaded_file.resource_path)
        self.__class__.file = uploaded_file

    async def test3_get_first_file(self):
        files = await self.folder_from.files.top(1).get().execute_query()
        self.assertEqual(len(files), 1)

    async def test4_get_file_from_absolute_url(self):
        result = await self.__class__.file.get_absolute_url().execute_query()
        file = await (
            File.from_url(result.value)
            .with_credentials(test_client_credentials)
            .get()
            .execute_query()
        )
        self.assertIsNotNone(file.serverRelativeUrl)

    async def test5_create_file_anon_link(self):
        result = await self.__class__.file.create_anonymous_link(False).execute_query()
        self.assertIsNotNone(result.value)

    async def test6_load_file_metadata(self):
        list_item = (
            await self.__class__.file.listItemAllFields.expand(["File"])
            .get()
            .execute_query()
        )
        self.assertIsInstance(list_item.file, File)

    async def test7_load_file_metadata_alt(self):
        list_item = self.__class__.file.listItemAllFields
        self.client.load(list_item, ["File"])
        await self.client.execute_query()
        self.assertIsInstance(list_item.file, File)

    async def test8_update_file_content(self):
        file = await self.__class__.file.save_binary_stream(
            self.text_content
        ).execute_query()
        self.assertTrue(file.resource_path)

    async def test9_update_file_metadata(self):
        list_item = self.__class__.file.listItemAllFields  # get metadata
        list_item.set_property("Title", "Updated")
        await list_item.update().execute_query()

    async def test_10_list_file_versions(self):
        file = await self.__class__.file.expand(["Versions"]).get().execute_query()
        self.assertGreater(len(file.versions), 0)

    async def test_11_delete_file_version(self):
        versions = await self.__class__.file.versions.top(1).get().execute_query()
        self.assertEqual(len(versions), 1)
        first_version = versions[0]
        self.assertIsNotNone(first_version.resource_path)
        await first_version.delete_object().execute_query()

    async def test_13_download_file_content(self):
        result = await self.__class__.file.get_content().execute_query()
        self.assertEqual(result.value, self.text_content)

    async def test_14_download_file_content_alt(self):
        with BytesIO() as f:
            await self.__class__.file.download(f).execute_query()
            content = f.getvalue()
        self.assertEqual(content, self.text_content)

    async def test_15_copy_file(self):
        copied_file = await self.__class__.file.copyto(
            self.folder_to, True
        ).execute_query()
        self.assertIsNotNone(copied_file.serverRelativeUrl)

    async def test_16_move_file(self):
        file = self.__class__.file
        moved_file = await file.moveto(self.folder_to, 1).get().execute_query()
        self.assertIsNotNone(moved_file.serverRelativeUrl)

    async def test_17_recycle_file(self):
        files_before = await self.folder_to.files.get().execute_query()
        result = await self.__class__.file.recycle().execute_query()
        self.assertIsNotNone(result.value)
        files_after = await self.folder_to.files.get().execute_query()
        self.assertEqual(len(files_before) - 1, len(files_after))
        self.__class__.deleted_file_guid = result.value

    async def test_18_restore_file(self):
        recycle_item = self.client.web.recycle_bin.get_by_id(
            self.__class__.deleted_file_guid
        )
        await recycle_item.restore().execute_query()
        self.assertIsNotNone(recycle_item.resource_path)

    # def test_18_create_template_file(self):
    #    file_url = "WikiPage.aspx"
    #    file = self.parent_folder.files.add_template_file(file_url, TemplateFileType.WikiPage).execute_query()
    #    self.assertEqual(file.name, file_url)

    async def test_19_get_files_changes(self):
        changes = await self.__class__.file.listItemAllFields.get_changes(
            ChangeQuery(item=True)
        ).execute_query()
        self.assertGreater(len(changes), 0)

    async def test_20_delete_file(self):
        files_before = await self.folder_to.files.get().execute_query()
        self.assertGreater(len(files_before), 0)
        await self.__class__.file.delete_object().execute_query()
        files_after = await self.folder_to.files.get().execute_query()
        self.assertEqual(len(files_after), len(files_before) - 1)

    async def test_22_upload_large_file(self):
        path = "{0}/../data/big_buck_bunny.mp4".format(os.path.dirname(__file__))
        file_size = os.path.getsize(path)
        size_1mb = 1000000
        file = await self.folder_from.files.create_upload_session(
            path, size_1mb
        ).execute_query()
        self.assertEqual(file_size, file.length)
