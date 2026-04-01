from random import randint

from office365.sharepoint.changes.query import ChangeQuery
from office365.sharepoint.contenttypes.collection import ContentTypeCollection
from office365.sharepoint.contenttypes.content_type import ContentType
from office365.sharepoint.contenttypes.creation_information import (
    ContentTypeCreationInformation,
)
from tests.sharepoint.sharepoint_case import SPTestCase


class TestContentType(SPTestCase):
    target_ct = None  # type: ContentType
    localized_title = "Contoso Dokumentti"

    async def test1_list_site_content_types(self):
        web_cts = await self.client.site.root_web.content_types.get().execute_query()
        self.assertIsInstance(web_cts, ContentTypeCollection)

    async def test2_get_content_type_by_id(self):
        ct = await (
            self.client.site.root_web.content_types.get_by_id("0x0101")
            .get()
            .execute_query()
        )
        self.assertIsNotNone(ct.name)

    async def test3_create_content_type(self):
        cti = ContentTypeCreationInformation("Contoso Document" + str(randint(0, 1000)))
        ct = await self.client.site.root_web.content_types.add(cti).execute_query()
        self.assertIsNotNone(ct.name)
        self.__class__.target_ct = ct

    async def test4_update_content_type(self):
        ct_to_update = self.__class__.target_ct
        ct_to_update.description = "New desc"
        await ct_to_update.update(True).execute_query()
        self.assertIsNotNone(ct_to_update.description)

    async def test5_set_value_for_ui_culture(self):
        ct = self.__class__.target_ct
        result = await ct.name_resource.set_value_for_ui_culture(
            "fi-FI", self.localized_title
        ).execute_query()
        self.assertIsNotNone(result.value)

    async def test6_get_value_for_ui_culture(self):
        ct = self.__class__.target_ct
        result = await ct.name_resource.get_value_for_ui_culture(
            "fi-FI"
        ).execute_query()
        self.assertIsNotNone(result.value)
        # self.assertEqual(result.value, self.localized_title)

    async def test8_delete_content_type(self):
        web_cts = await self.client.site.root_web.content_types.get().execute_query()
        before_count = len(web_cts)
        await self.__class__.target_ct.delete_object().execute_query()
        web_cts = await self.client.site.root_web.content_types.get().execute_query()
        self.assertTrue(before_count, len(web_cts) + 1)

    async def test9_get_content_types_changes(self):
        changes = await self.client.web.get_changes(
            ChangeQuery(content_type=True)
        ).execute_query()
        self.assertGreater(len(changes), 0)
