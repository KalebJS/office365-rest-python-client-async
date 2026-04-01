import asyncio
import uuid

from office365.sharepoint.principal.groups.group import Group
from tests import test_user_principal_name
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSharePointGroup(SPTestCase):
    target_group = None  # type: Group

    @classmethod
    def setUpClass(cls):
        super(TestSharePointGroup, cls).setUpClass()

        async def _async_setup():
            cls.target_user = await cls.client.web.ensure_user(
                test_user_principal_name
            ).execute_query()

        asyncio.run(_async_setup())

    async def test1_create_group(self):
        grp_title = "Custom Group" + uuid.uuid4().hex
        result = await self.client.web.site_groups.add(grp_title).execute_query()
        self.assertIsNotNone(result.resource_path)
        self.__class__.target_group = result

    async def test2_add_user_to_group(self):
        target_user = await self.__class__.target_group.users.add_user(
            self.target_user.login_name
        ).execute_query()
        self.assertIsNotNone(target_user.id)

    async def test3_get_group_users(self):
        result = await self.__class__.target_group.users.get().execute_query()
        self.assertGreaterEqual(len(result), 1)

    async def test4_expand_to_principals(self):
        result = (
            await self.__class__.target_group.expand_to_principals().execute_query()
        )
        self.assertIsNotNone(result.value)

    async def test5_remove_user_from_group(self):
        result = await self.__class__.target_group.users.remove_by_id(
            self.target_user.id
        ).execute_query()
        self.assertEqual(len(result), 0)

    async def test6_delete_group(self):
        grp_id = self.__class__.target_group.id
        await self.client.web.site_groups.remove_by_id(grp_id).execute_query()
