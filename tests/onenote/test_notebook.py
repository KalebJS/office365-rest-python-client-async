from office365.onenote.notebooks.notebook import Notebook
from tests import create_unique_name
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestNotebook(GraphTestCase):
    target_notebook = None  # type: Notebook

    @requires_delegated_permission("Notes.Create", "Notes.ReadWrite", "Notes.ReadWrite.All")
    async def test1_create_notebook(self):
        notebook_name = create_unique_name("My Private notebook")
        new_notebook = await self.client.me.onenote.notebooks.add(notebook_name).execute_query()
        self.assertIsNotNone(new_notebook.resource_path)
        self.__class__.target_notebook = new_notebook

    @requires_delegated_permission("Notes.Create")
    async def test2_list_notebooks(self):
        my_notebooks = await self.client.me.onenote.notebooks.get().execute_query()
        self.assertIsNotNone(my_notebooks.resource_path)

    @requires_delegated_permission(
        "Notes.Create",
        "Notes.Read",
        "Notes.Read.All",
        "Notes.ReadWrite",
        "Notes.ReadWrite.All",
    )
    async def test3_get_recent_notebooks(self):
        result = await self.client.me.onenote.notebooks.get_recent_notebooks().execute_query()
        self.assertIsNotNone(result.value)

    @requires_delegated_permission(
        "Notes.Create",
        "Notes.ReadWrite",
        "Notes.ReadWrite.All",
    )
    async def test4_create_section(self):
        name = create_unique_name("Section name")
        new_section = await self.__class__.target_notebook.sections.add(displayName=name).execute_query()
        self.assertIsNotNone(new_section.resource_path)
        self.__class__.target_section = new_section
