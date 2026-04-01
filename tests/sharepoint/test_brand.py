from tests.sharepoint.sharepoint_case import SPTestCase


class TestBrand(SPTestCase):
    async def test1_get_site_themes(self):
        result = await self.client.brand_center.get_site_themes().execute_query()
        self.assertIsNotNone(result.value)

    async def test2_get_configuration(self):
        result = await self.client.brand_center.configuration().execute_query()
        self.assertIsNotNone(result.value)
