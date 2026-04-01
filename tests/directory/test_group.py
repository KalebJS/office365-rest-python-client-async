import unittest

from office365.directory.groups.group import Group
from office365.directory.users.user import User
from office365.runtime.client_request_exception import ClientRequestException
from tests import create_unique_name, test_user_principal_name
from tests.graph_case import GraphTestCase


class TestGraphGroup(GraphTestCase):
    """Tests for Azure Active Directory (Azure AD) groups"""

    target_group = None  # type: Group
    target_user = None  # type: User
    directory_quota_exceeded = False

    async def test1_create_group(self):
        try:
            name = create_unique_name("Group")
            new_group = await self.client.groups.create_m365(name).execute_query()
            self.assertIsNotNone(new_group.id)
            self.__class__.target_group = new_group
        except ClientRequestException as e:
            if e.code == "Directory_QuotaExceeded":
                self.directory_quota_exceeded = True
                result = await self.client.me.get_member_groups().execute_query()
                self.assertIsNotNone(result.value)
                filter_expr = "displayName eq '{0}'".format(result.value[0])
                result = (
                    await self.client.groups.filter(filter_expr).get().execute_query()
                )
                self.__class__.target_group = result[0]

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not be created")
    async def test2_list_groups(self):
        groups = await self.client.groups.top(1).get().execute_query()
        self.assertEqual(len(groups), 1)

    async def test3_get_groups_count(self):
        result = await self.client.groups.count().execute_query()
        self.assertIsNotNone(result.value)

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not be created")
    async def test4_get_group(self):
        existing_group = self.__class__.target_group
        target_group = await self.client.groups[existing_group.id].get().execute_query()
        self.assertIsInstance(target_group, Group)

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not be created")
    async def test5_add_group_owner(self):
        users = await (
            self.client.users.filter(
                "mail eq '{mail}'".format(mail=test_user_principal_name)
            )
            .get()
            .execute_query()
        )
        self.assertEqual(len(users), 1)

        owner = users[0]
        grp = self.__class__.target_group
        await grp.owners.add(owner).execute_query()
        self.__class__.target_user = users[0]

    async def test6_list_group_owners(self):
        owners = await self.__class__.target_group.owners.get().execute_query()
        self.assertGreater(len(owners), 0)

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not created")
    async def test7_remove_group_owner(self):
        owner_id = self.__class__.target_user.id
        grp = self.__class__.target_group
        await grp.owners.remove(owner_id).execute_query()

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not created")
    async def test8_add_group_member(self):
        member = self.__class__.target_user
        grp = self.__class__.target_group
        await grp.members.add(member).execute_query()

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not created")
    async def test9_remove_group_member(self):
        member_id = self.__class__.target_user.id
        grp = self.__class__.target_group
        await grp.members.remove(member_id).execute_query()

    @unittest.skipIf(directory_quota_exceeded, "Skipping, group was not created")
    async def test_10_delete_group(self):
        grp_to_delete = self.__class__.target_group
        await grp_to_delete.delete_object(True).execute_query()

    async def test_11_get_changes(self):
        changed_groups = (
            await self.client.groups.delta.select(["displayName"]).get().execute_query()
        )
        self.assertGreater(len(changed_groups), 0)
