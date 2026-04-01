import asyncio
import uuid
from unittest import TestCase

from office365.graph_client import GraphClient
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.permissions.permission import Permission
from tests import (
    test_client_credentials,
    test_client_id,
    test_client_secret,
    test_team_site_url,
    test_tenant,
    test_user_principal_name_alt,
)
from tests.decorators import requires_app_permission


class TestPermissions(TestCase):
    target_drive_item = None  # type: DriveItem
    target_permission = None  # type: Permission

    @classmethod
    def setUpClass(cls):
        super(TestPermissions, cls).setUpClass()
        client = GraphClient(tenant=test_tenant).with_client_secret(
            test_client_id, test_client_secret
        )
        folder_name = "New_" + uuid.uuid4().hex

        async def _async_setup():
            cls.target_drive_item = await client.sites.root.drive.root.create_folder(
                folder_name
            ).execute_query()
            cls.client = client

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            item_to_delete = await cls.target_drive_item.get().execute_query()
            await item_to_delete.delete_object().execute_query()

        asyncio.run(_async_teardown())

    @requires_app_permission("Files.ReadWrite.All", "Sites.ReadWrite.All")
    async def test1_create_anonymous_link(self):
        permission = await self.__class__.target_drive_item.create_link(
            "view", "anonymous"
        ).execute_query()
        self.assertIsNotNone(permission.id)
        self.assertIsNotNone(permission.roles[0], "read")

    @requires_app_permission("Files.ReadWrite.All", "Sites.ReadWrite.All")
    async def test2_create_company_link(self):
        permission = await self.__class__.target_drive_item.create_link(
            "edit", "organization"
        ).execute_query()
        self.assertIsNotNone(permission.id)
        self.assertIsNotNone(permission.roles[0], "write")

    @requires_app_permission(
        "Files.Read.All", "Files.ReadWrite.All", "Sites.Read.All", "Sites.ReadWrite.All"
    )
    async def test4_driveitem_list_permissions(self):
        permissions = (
            await self.__class__.target_drive_item.permissions.get().execute_query()
        )
        self.assertIsNotNone(permissions.resource_path)
        self.assertGreater(len(permissions), 0)

    @requires_app_permission(
        "Files.Read.All", "Files.ReadWrite.All", "Sites.Read.All", "Sites.ReadWrite.All"
    )
    async def test5_driveitem_get_permission(self):
        result = (
            await self.__class__.target_drive_item.permissions.get()
            .top(1)
            .execute_query()
        )
        self.assertEqual(len(result), 1)
        perm_id = result[0].id
        perm = (
            await self.__class__.target_drive_item.permissions[perm_id]
            .get()
            .execute_query()
        )
        self.assertIsNotNone(perm.resource_path)
        self.__class__.target_permission = result[0]

    async def test6_driveitem_update_permission(self):
        # perm_to_update = self.__class__.target_permission
        # perm_to_update.roles = ["read"]
        # perm_to_update.update().execute_query()
        pass

    @requires_app_permission("Files.ReadWrite.All", "Sites.ReadWrite.All")
    async def test7_driveitem_delete_permission(self):
        perm_to_delete = self.__class__.target_permission
        await perm_to_delete.delete_object().execute_query()

    async def test8_driveitem_grant_access(self):
        file_abs_url = "{0}/Shared Documents/Financial Sample.xlsx".format(
            test_team_site_url
        )
        permissions = await (
            self.client.shares.by_url(file_abs_url)
            .permission.grant(recipients=[test_user_principal_name_alt], roles=["read"])
            .execute_query()
        )
        self.assertIsNotNone(permissions.resource_path)

    async def test9_create_site_permission(self):
        app = self.client.applications.get_by_app_id(test_client_credentials.clientId)
        new_site_permission = await self.client.sites.root.permissions.add(
            ["write"], app
        ).execute_query()
        self.assertIsNotNone(new_site_permission.resource_path)
        self.target_permission = new_site_permission

    async def test_10_list_site_permissions(self):
        site_permissions = (
            await self.client.sites.root.permissions.get().execute_query()
        )
        self.assertIsNotNone(site_permissions.resource_path)

    async def test_11_delete_site_permission(self):
        await self.target_permission.delete_object().execute_query()
