from office365.sharepoint.publishing.pages.collection import SitePageCollection
from office365.sharepoint.publishing.pages.page import SitePage
from tests import create_unique_name
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSitePages(SPTestCase):
    target_page = None  # type: SitePage

    @classmethod
    def setUpClass(cls):
        super(TestSitePages, cls).setUpClass()

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_create_draft_page(self):
        page_title = create_unique_name("Site Page ")
        page = await self.client.site_pages.create_page(
            page_title, "Site Page.aspx"
        ).execute_query()
        self.assertIsNotNone(page.resource_path)
        self.__class__.target_page = page

    async def test2_list_site_pages(self):
        result = await self.client.site_pages.pages.get().execute_query()
        self.assertIsInstance(result, SitePageCollection)
        self.assertIsNotNone(result.resource_path)

    async def test3_publish_site_page(self):
        page = self.__class__.target_page
        await page.publish().execute_query()
        self.assertIsNotNone(page.first_published)

    async def test5_delete_site_page(self):
        page = self.__class__.target_page
        await page.delete_object().execute_query()
