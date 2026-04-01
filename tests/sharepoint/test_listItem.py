import asyncio
from random import randint
from time import sleep

from office365.sharepoint.listitems.caml.query import CamlQuery
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.lists.creation_information import ListCreationInformation
from office365.sharepoint.lists.list import List
from office365.sharepoint.lists.template_type import ListTemplateType
from office365.sharepoint.types.wopi_action import SPWOPIAction
from tests import create_unique_name
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSharePointListItem(SPTestCase):
    target_list = None  # type: List
    target_item = None  # type: ListItem
    deleted_item_guid = None  # type: str

    @classmethod
    def setUpClass(cls):
        super(TestSharePointListItem, cls).setUpClass()
        target_list_title = "Tasks" + str(randint(0, 10000))
        cls.default_title = create_unique_name("Task")
        cls.batch_items_count = 3

        async def _async_setup():
            cls.target_list = await cls.ensure_list(
                cls.client.web,
                ListCreationInformation(target_list_title, None, ListTemplateType.Tasks),
            )

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_list.delete_object().execute_query()

        asyncio.run(_async_teardown())

    async def test1_create_list_item(self):
        item_properties = {"Title": self.default_title}
        new_item = await self.target_list.add_item(item_properties).execute_query()
        self.assertIsNotNone(new_item.properties["Title"])
        self.__class__.target_item = new_item

    async def test2_enable_folders_in_list(self):
        async def _init_list():
            if not self.target_list.enable_folder_creation:
                self.target_list.enable_folder_creation = True
                await self.target_list.update().execute_query()
            self.assertTrue(self.target_list.enable_folder_creation, "Folder creation enabled")

        await self.target_list.ensure_property("EnableFolderCreation", _init_list).execute_query()

    async def test3_create_folder_in_list(self):
        new_folder = await self.target_list.root_folder.add("Archive").execute_query()
        self.assertIsNotNone(new_folder.serverRelativeUrl)

    async def test4_get_list_item(self):
        item_id = self.__class__.target_item.id
        item = await self.target_list.get_item_by_id(item_id).get().execute_query()
        self.assertIsNotNone(item.id)

    async def test5_get_list_item_via_caml(self):
        item_id = self.__class__.target_item.id
        caml_query = CamlQuery.parse(
            "<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>{0}</Value></Eq></Where>".format(item_id)
        )
        result = await self.target_list.get_items(caml_query).execute_query()
        self.assertEqual(len(result), 1)

    async def test6_get_wopi_frame_url(self):
        result = await self.__class__.target_item.get_wopi_frame_url(SPWOPIAction.default).execute_query()
        self.assertIsNotNone(result.value)

    async def test7_update_listItem(self):
        item_to_update = await self.__class__.target_item.get().execute_query()
        last_updated = item_to_update.properties["Modified"]

        sleep(1)
        new_title = create_unique_name("Task item")
        item_to_update.set_property("Title", new_title).update()
        self.client.load(item_to_update)  # retrieve updated
        await self.client.execute_query()
        self.assertNotEqual(item_to_update.properties["Modified"], last_updated)
        self.assertNotEqual(self.default_title, new_title)

    async def test8_systemUpdate_listItem(self):
        item_to_update = await self.__class__.target_item.get().execute_query()
        last_updated = item_to_update.properties["Modified"]

        new_title = create_unique_name("Task item %s")
        item_to_update.set_property("Title", new_title).system_update()
        self.client.load(item_to_update)  # retrieve updated
        await self.client.execute_query()
        self.assertEqual(item_to_update.properties["Modified"], last_updated)
        self.assertNotEqual(self.default_title, new_title)

    async def test9_update_overwrite_version(self):
        item_to_update = self.__class__.target_item
        await item_to_update.update_overwrite_version().execute_query()

    async def test_11_get_versions(self):
        versions = await self.__class__.target_item.versions.get().execute_query()
        self.assertIsNotNone(versions.resource_path)

    async def test_12_get_dlp_policy_tip(self):
        result = await self.__class__.target_item.get_dlp_policy_tip.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test_13_enable_comments(self):
        comments = await self.__class__.target_item.set_comments_disabled(False).execute_query()
        self.assertIsNotNone(comments.resource_path)

    async def test_10_get_comments(self):
        comments = await self.__class__.target_item.get_comments().execute_query()
        self.assertIsNotNone(comments.resource_path)

    async def test_14_recycle_item(self):
        item_to_recycle = self.__class__.target_item
        result = await item_to_recycle.recycle().execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.deleted_item_guid = result.value

    async def test_15_restore_item(self):
        recycle_item = self.client.web.recycle_bin.get_by_id(self.__class__.deleted_item_guid)
        await recycle_item.restore().execute_query()
        self.assertIsNotNone(recycle_item.resource_path)

    async def test_16_set_rating(self):
        result = await self.__class__.target_item.set_rating(1).execute_query()
        self.assertIsNotNone(result.value)

    async def test_17_delete_list_item(self):
        item_id = self.__class__.target_item.properties["Id"]
        item_to_delete = self.__class__.target_item
        await item_to_delete.delete_object().execute_query()

        result = await self.target_list.items.filter("Id eq {0}".format(item_id)).get().execute_query()
        self.assertEqual(0, len(result))

    async def test_18_create_multiple_items(self):
        for i in range(0, self.batch_items_count):
            item_properties = {"Title": "Task {0}".format(i)}
            self.target_list.add_item(item_properties)
        self.client.execute_batch()
        result = await self.target_list.items.get().execute_query()
        self.assertEqual(len(result), self.batch_items_count)

    async def test_19_get_multiple_items_with_params(self):
        # test case for when .load with set properties_to_retrieve
        # would ignore all other previously set query params (like top(2))

        items = self.target_list.items.top(2)
        self.client.load(items, ["Id", "AttachmentFiles"])
        await self.client.execute_query()
        self.assertLessEqual(len(items), 2)

    async def test_20_delete_multiple_items(self):
        items = await self.target_list.items.get().execute_query()  # get existing items
        self.assertGreater(len(items), 0)
        [item.delete_object() for item in items]
        self.client.execute_batch()
        items_after = await self.target_list.items.get().execute_query()
        self.assertEqual(len(items_after), 0)

    async def test_21_get_all_items(self):
        users_list = self.client.web.lists.get_by_title("User Information List")
        items = await users_list.items.get_all(page_size=1000).execute_query()
        items_count = len(items)
        items = await users_list.items.get_all(page_size=1000).execute_query()
        self.assertEqual(len(items), items_count)
