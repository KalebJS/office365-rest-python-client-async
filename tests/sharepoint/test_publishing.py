from office365.sharepoint.publishing.pages.collection import SitePageCollection
from office365.sharepoint.publishing.pages.service import SitePageService
from office365.sharepoint.publishing.video.service_discoverer import (
    VideoServiceDiscoverer,
)
from tests.sharepoint.sharepoint_case import SPTestCase


class TestPublishing(SPTestCase):
    @classmethod
    def setUpClass(cls):
        super(TestPublishing, cls).setUpClass()

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_init_site_page_service(self):
        svc = await self.client.site_pages.get().execute_query()
        self.assertIsNotNone(svc.resource_path)

    # def test3_get_time_zone(self):
    #    time_zone = SitePageService.get_time_zone(self.client, "Moscow").execute_query()
    #    self.assertIsInstance(time_zone, PrimaryCityTime)

    async def test4_compute_file_name(self):
        result = await SitePageService.compute_file_name(self.client, "Test page").execute_query()
        self.assertIsNotNone(result.value)

    async def test5_file_picker_tab_options(self):
        result = await SitePageService.file_picker_tab_options(self.client).execute_query()
        self.assertIsNotNone(result.value)

    async def test6_org_assets(self):
        result = await SitePageService.org_assets(self.client).execute_query()
        self.assertIsNotNone(result.value)

    async def test7_get_video_service_manager(self):
        discoverer = await VideoServiceDiscoverer(self.client).get().execute_query()
        self.assertIsNotNone(discoverer.video_portal_url)

    async def test8_get_page_by_name(self):
        page = await self.client.site_pages.pages.get_by_name("Home.aspx").get().execute_query()
        self.assertIsNotNone(page.resource_path)

    async def test9_can_create_page(self):
        result = await self.client.site_pages.can_create_page().execute_query()
        self.assertIsNotNone(result.value)

    async def test_10_get_current_user_memberships(self):
        result = await SitePageService.get_current_user_memberships(self.client).execute_query()
        self.assertIsNotNone(result.value)

    async def test_11_get_page_diagnostics(self):
        result = await self.client.page_diagnostics.by_page("/sites/team/SitePages/Home.aspx").execute_query()
        self.assertIsNotNone(result.value)

    async def test_12_checkout_page(self):
        page = self.client.site_pages.pages.get_by_name("Home.aspx")
        await page.checkout_page().execute_query()
        self.assertIsNotNone(page.resource_path)
        self.assertTrue(page.is_page_checked_out_to_current_user)

    async def test_13_discard_page(self):
        page = self.client.site_pages.pages.get_by_name("Home.aspx")
        await page.discard_page().execute_query()
        self.assertFalse(
            page.is_page_checked_out_to_current_user,
            "Page is expected to be checked in",
        )

    # def test_12_share_page_preview_by_email(self):
    #    page = self.client.site_pages.pages.get_by_url("/sites/team/SitePages/Home.aspx")
    #    page.share_page_preview_by_email("This page has been shared with you",
    #                                     [test_user_principal_name_alt]).execute_query()
