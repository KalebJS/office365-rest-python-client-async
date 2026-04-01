import asyncio
import uuid

from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.drives.drive import Drive
from office365.onedrive.lists.template_type import ListTemplateType
from office365.runtime.paths.v4.entity import EntityPath
from tests import create_unique_name
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestFolder(GraphTestCase):
    """OneDrive test case for a Folder"""

    target_drive = None  # type: Drive
    target_folder = None  # type: DriveItem
    target_folder_name = "Archive_" + uuid.uuid4().hex

    @classmethod
    def setUpClass(cls):
        super(TestFolder, cls).setUpClass()
        lib_name = create_unique_name("Lib")

        async def _async_setup():
            lib = await cls.client.sites.root.lists.add(
                lib_name, ListTemplateType.documentLibrary
            ).execute_query()
            cls.target_drive = lib.drive

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_drive.list.delete_object().execute_query()

        asyncio.run(_async_teardown())

    @requires_delegated_permission(
        "Files.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All"
    )
    async def test1_create_root_folder(self):
        folder = await self.target_drive.root.create_folder(
            self.target_folder_name
        ).execute_query()
        self.assertEqual(folder.name, self.target_folder_name)
        self.__class__.target_folder = folder

    @requires_delegated_permission(
        "Files.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All"
    )
    async def test2_create_child_folder(self):
        target_folder_name = "2018"
        folder = await self.__class__.target_folder.create_folder(
            target_folder_name
        ).execute_query()
        self.assertEqual(folder.name, target_folder_name)

    @requires_delegated_permission(
        "Files.Read",
        "Files.ReadWrite",
        "Files.Read.All",
        "Files.ReadWrite.All",
        "Group.Read.All",
        "Group.ReadWrite.All",
        "Sites.Read.All",
        "Sites.ReadWrite.All",
    )
    async def test3_get_folder_by_path(self):
        root_folder = await (
            self.target_drive.root.get_by_path(self.target_folder_name)
            .get()
            .execute_query()
        )
        folder = await root_folder.get_by_path("2018").get().execute_query()
        self.assertEqual(
            folder.resource_path,
            EntityPath(folder.id, self.target_drive.items.resource_path),
        )

    async def test4_get_folder_permissions(self):
        result = await self.__class__.target_folder.permissions.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    @requires_delegated_permission(
        "Files.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All"
    )
    async def test5_update_folder(self):
        folder = self.__class__.target_folder
        await folder.update().execute_query()

    async def test6_get_analytics(self):
        result = await self.__class__.target_folder.analytics.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test7_delete_folder(self):
        await self.__class__.target_folder.delete_object().execute_query()
