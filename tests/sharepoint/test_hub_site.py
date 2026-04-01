import asyncio

from office365.sharepoint.portal.hub_sites_utility import SPHubSitesUtility
from office365.sharepoint.sites.site import Site
from tests.sharepoint.sharepoint_case import SPTestCase


class TestHubSite(SPTestCase):
    @classmethod
    def setUpClass(cls):
        super(TestHubSite, cls).setUpClass()

        async def _async_setup():
            cls.target_site = await cls.client.site.get().execute_query()  # type: Site

        asyncio.run(_async_setup())

    async def test1_register_hub_site(self):
        if not self.target_site.is_hub_site and not self.target_site.hub_site_id:
            site = await self.target_site.register_hub_site().get().execute_query()
            self.assertTrue(site.is_hub_site)

    async def test2_get_hub_sites(self):
        hub_sites = await SPHubSitesUtility(self.client).get_hub_sites().execute_query()
        self.assertGreater(len(hub_sites), 0)

    async def test3_unregister_hub_site(self):
        if self.target_site.is_hub_site:
            site = await self.target_site.unregister_hub_site().get().execute_query()
            self.assertFalse(site.is_hub_site)
