from tests.sharepoint.sharepoint_case import SPTestCase


class TestOrgNews(SPTestCase):
    async def test_1_get_org_news(self):
        result = await self.client.org_news.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test2_sites_reference(self):
        result = await self.client.org_news.sites_reference().execute_query()
        self.assertIsNotNone(result.value)
