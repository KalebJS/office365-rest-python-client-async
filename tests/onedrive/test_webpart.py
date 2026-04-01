import asyncio

from office365.onedrive.sitepages.site_page import SitePage
from tests import create_unique_name, test_team_site_url
from tests.graph_case import GraphTestCase


class TestWebPart(GraphTestCase):
    """OneDrive specific test case for web parts"""

    target_page = None  # type: SitePage

    @classmethod
    def setUpClass(cls):
        super(TestWebPart, cls).setUpClass()
        test_site = cls.client.sites.get_by_url(test_team_site_url)
        page_name = create_unique_name("Test Page")

        async def _async_setup():
            cls.target_page = await test_site.pages.add(page_name).checkin("").execute_query()

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_page.delete_object().execute_query()

        asyncio.run(_async_teardown())

    async def test1_list_web_parts(self):
        result = await self.target_page.web_parts.get().execute_query()
        self.assertIsNotNone(result.resource_path)
