from office365.onedrive.columns.definition import ColumnDefinition
from office365.onedrive.lists.list import List
from tests import create_unique_name
from tests.graph_case import GraphTestCase


class TestList(GraphTestCase):
    """OneDrive specific test case base class"""

    target_list = None  # type: List
    target_column = None  # type: ColumnDefinition
    list_name = create_unique_name("Documents")

    @classmethod
    def setUpClass(cls):
        super(TestList, cls).setUpClass()

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_create_list(self):
        result = await self.client.sites.root.lists.add(self.list_name, "documentLibrary").execute_query()
        self.__class__.target_list = result

    async def test2_get_list(self):
        target_list = await self.client.sites.root.lists[self.list_name].get().execute_query()
        self.assertIsNotNone(target_list.resource_path)

    async def test3_get_list_items(self):
        items = await self.target_list.items.get().execute_query()
        self.assertIsNotNone(items.resource_path)

    async def test4_get_list_columns(self):
        columns = await self.target_list.columns.get().execute_query()
        self.assertIsNotNone(columns.resource_path)

    async def test5_create_list_column(self):
        column_name = create_unique_name("Text")
        text_column = await self.target_list.columns.add_text(column_name).execute_query()
        self.assertIsNotNone(text_column.resource_path)
        self.__class__.target_column = text_column

    async def test6_delete_list_column(self):
        column_to_del = self.__class__.target_column
        await column_to_del.delete_object().execute_query()

    async def test7_delete_list(self):
        await self.__class__.target_list.delete_object().execute_query()

    async def test8_get_pages_list(self):
        result = await self.client.sites.root.lists.get_by_name("Site Pages").get().execute_query()
        self.assertIsNotNone(result.resource_path)
