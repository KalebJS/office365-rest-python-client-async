import asyncio
import os
import uuid
from datetime import datetime, timedelta

from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.drives.drive import Drive
from office365.onedrive.lists.template_type import ListTemplateType
from tests import create_unique_name
from tests.graph_case import GraphTestCase


class TestFile(GraphTestCase):
    """OneDrive specific test case base class"""

    target_drive = None  # type: Drive
    target_file = None  # type: DriveItem
    target_folder = None  # type: DriveItem

    @classmethod
    def setUpClass(cls):
        super(TestFile, cls).setUpClass()
        lib_name = create_unique_name("Lib")

        async def _async_setup():
            lib = await cls.client.sites.root.lists.add(lib_name, ListTemplateType.documentLibrary).execute_query()
            cls.target_drive = lib.drive

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_drive.list.delete_object().execute_query()

        asyncio.run(_async_teardown())

    async def test1_create_folder(self):
        target_folder_name = "New_" + uuid.uuid4().hex
        folder = await self.target_drive.root.create_folder(target_folder_name).execute_query()
        self.assertEqual(folder.name, target_folder_name)
        self.__class__.target_folder = folder

    async def test2_get_folder_permissions(self):
        folder_perms = await self.__class__.target_folder.permissions.get().execute_query()
        self.assertIsNotNone(folder_perms.resource_path)

    async def test3_upload_file(self):
        file_name = "SharePoint User Guide.docx"
        file_path = "{0}/../data/{1}".format(os.path.dirname(__file__), file_name)
        self.__class__.target_file = await self.target_drive.root.upload_file(file_path).execute_query()
        self.assertIsNotNone(self.target_file.web_url)

    async def test4_preview_file(self):
        result = await self.__class__.target_file.preview("1").execute_query()
        self.assertIsNotNone(result.value)

    # def test5_validate_permission(self):
    #    self.__class__.target_file.validate_permission().execute_query()

    async def test6_checkout(self):
        await self.__class__.target_file.checkout().execute_query()
        target_item = await self.__class__.target_file.get().select(["publication"]).execute_query()
        self.assertEqual(target_item.publication.level, "checkout")

    async def test7_checkin(self):
        await self.__class__.target_file.checkin("").execute_query()
        target_item = await self.__class__.target_file.get().select(["publication"]).execute_query()
        self.assertEqual(target_item.publication.level, "published")

    async def test8_list_versions(self):
        versions = await self.__class__.target_file.versions.get().execute_query()
        self.assertGreater(len(versions), 1)

    # def test9_follow(self):
    #    target_item = self.__class__.target_file.follow().execute_query()
    #    self.assertIsNotNone(target_item.resource_path)

    # def test_10_unfollow(self):
    #    target_item = self.__class__.target_file.unfollow().execute_query()
    #    self.assertIsNotNone(target_item.resource_path)

    async def test_11_upload_file_session(self):
        file_name = "big_buck_bunny.mp4"
        local_path = "{0}/../data/{1}".format(os.path.dirname(__file__), file_name)
        target_file = await self.target_drive.root.resumable_upload(local_path).get().execute_query()
        self.assertIsNotNone(target_file.web_url)

    async def test_12_download_file(self):
        result = await self.__class__.target_file.get_content().execute_query()
        self.assertIsNotNone(result.value)

    async def test_13_convert_file(self):
        result = await self.__class__.target_file.convert("pdf").execute_query()
        self.assertIsNotNone(result.value)

    async def test_14_copy_file(self):
        file_name = "Copied_{0}_SharePoint User Guide.docx".format(uuid.uuid4().hex)
        result = await self.__class__.target_file.copy(file_name).execute_query()
        self.assertIsNotNone(result.value)

    # def test_14_move_file(self):
    #    target_folder = self.__class__.target_folder.parentReference

    #    file_name = "Moved_{0}_SharePoint User Guide.docx".format(uuid.uuid4().hex)
    #    result = self.__class__.target_file.move(file_name, target_folder)
    #    self.client.execute_query()
    #    self.assertIsNotNone(result.value)

    async def test_15_get_activities_by_interval(self):
        end_time = datetime.now()
        start_time = end_time - timedelta(days=14)
        result = await self.__class__.target_file.get_activities_by_interval(start_time, end_time, "day").execute_query()
        self.assertIsNotNone(result)

    async def test_16_get_item_analytics(self):
        result = await self.__class__.target_file.analytics.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test_17_extract_sensitivity_labels(self):
        result = await self.__class__.target_file.extract_sensitivity_labels().execute_query()
        self.assertIsNotNone(result.value)

    async def test_18_delete_file(self):
        items = await self.target_drive.root.children.top(2).get().execute_query()
        for item in items:
            await item.delete_object().execute_query()
