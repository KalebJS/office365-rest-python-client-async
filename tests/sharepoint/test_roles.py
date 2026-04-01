from office365.sharepoint.permissions.base_permissions import BasePermissions
from office365.sharepoint.permissions.kind import PermissionKind
from office365.sharepoint.permissions.roles.definitions.definition import RoleDefinition
from tests.sharepoint.sharepoint_case import SPTestCase


class TestRoles(SPTestCase):
    target_object = None  # type: RoleDefinition
    role_name = "Create and Manage Alerts 123"

    async def test1_create_role(self):
        permissions = BasePermissions()
        permissions.set(PermissionKind.CreateAlerts)
        permissions.set(PermissionKind.ManageAlerts)
        result = await self.client.web.role_definitions.add(
            permissions, self.role_name
        ).execute_query()
        self.assertIsNotNone(result.resource_path)
        self.__class__.target_object = result

    # def test2_get_by_type(self):
    #    result = self.client.web.role_definitions.get_by_type(PermissionKind.CreateAlerts).get().execute_query()
    #    self.assertIsNotNone(result.resource_path)

    async def test3_get_by_name(self):
        result = await (
            self.client.web.role_definitions.get_by_name(self.role_name)
            .get()
            .execute_query()
        )
        self.assertIsNotNone(result.resource_path)

    async def test4_add_role_assignment(self):
        target_user = self.client.web.current_user
        result = await self.client.web.add_role_assignment(
            target_user, self.__class__.target_object
        ).execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test5_remove_role_assignment(self):
        target_user = self.client.web.current_user
        result = await self.client.web.remove_role_assignment(
            target_user, self.__class__.target_object
        ).execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test6_delete_role(self):
        await self.__class__.target_object.delete_object().execute_query()
