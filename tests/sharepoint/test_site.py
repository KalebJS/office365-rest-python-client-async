import uuid

from office365.sharepoint.lists.template_type import ListTemplateType
from office365.sharepoint.portal.sites.creation_response import SPSiteCreationResponse
from office365.sharepoint.portal.sites.status import SiteStatus
from office365.sharepoint.sites.site import Site
from tests import test_admin_credentials, test_site_url, test_user_principal_name_alt
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSite(SPTestCase):
    site_response = None  # type: SPSiteCreationResponse

    async def test1_if_site_loaded(self):
        site = await self.client.site.get().execute_query()
        self.assertIs(site.is_property_available("Url"), True, "Site resource was not requested")
        self.assertIs(site.is_property_available("RootWeb"), False)

    async def test2_if_site_exists(self):
        site_url = self.client.site.url
        result = await Site.exists(self.client, site_url).execute_query()
        self.assertIsNotNone(result.value)

    async def test3_get_site_by_id(self):
        site_id = self.client.site.id
        result = await Site.get_url_by_id(self.client, site_id).execute_query()
        self.assertIsNotNone(result.value)

    async def test4_check_is_deletable(self):
        result = await self.client.site.check_is_deletable().execute_query()
        self.assertIsNotNone(result.value)

    async def test5_get_site_catalog(self):
        catalog = await self.client.site.get_catalog(ListTemplateType.AppDataCatalog).get().execute_query()
        self.assertIsNotNone(catalog.title)

    async def test6_get_web_templates(self):
        web_templates = await self.client.site.get_web_templates().execute_query()
        self.assertIsNotNone(web_templates)

    async def test7_get_web_template_by_name(self):
        template_name = "GLOBAL#0"
        web_template = await self.client.site.get_web_templates().get_by_name(template_name).get().execute_query()
        self.assertIsNotNone(web_template)

    async def test8_get_site_logo(self):
        result = await self.client.site.get_site_logo().execute_query()
        self.assertIsNotNone(result.value)

    async def test_10_open_web_by_id(self):
        web = await self.client.web.get().execute_query()
        sub_site = await self.client.site.open_web_by_id(web.id).execute_query()
        self.assertIsNotNone(sub_site.id)

    # def test_10_get_site_links(self):
    #    result = self.client.site_linking_manager.get_site_links().execute_query()
    #    self.assertIsNotNone(result.value)

    async def test_11_create_site(self):
        site_url = "{0}/sites/{1}".format(test_site_url, uuid.uuid4().hex)
        result = await self.client.site_manager.create(
            "Comm Site", site_url, test_user_principal_name_alt
        ).execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.site_response = result.value

    async def test_12_get_site_status(self):
        site_url = self.__class__.site_response.SiteUrl
        result = await self.client.site_manager.get_status(site_url).execute_query()
        self.assertIsNotNone(result.value.SiteStatus)
        self.assertTrue(result.value.SiteStatus != SiteStatus.Error)

    async def test_13_get_site_url(self):
        site_id = self.__class__.site_response.SiteId
        result = await self.client.site_manager.get_site_url(site_id).execute_query()
        self.assertIsNotNone(result.value)
        self.assertTrue(self.__class__.site_response.SiteUrl == result.value)

    async def test_14_is_deletable(self):
        result = await self.client.site.is_deletable().execute_query()
        self.assertIsNotNone(result.value)

    async def test_15_delete_site(self):
        from office365.sharepoint.client_context import ClientContext

        admin_ctx = ClientContext(self.client.base_url).with_credentials(test_admin_credentials)
        site_id = self.__class__.site_response.SiteId
        await admin_ctx.site_manager.delete(site_id).execute_query()

    # def test_16_get_block_download_policy_for_files_data(self):
    #    result = self.client.site.get_block_download_policy_for_files_data().execute_query()
    #    self.assertIsNotNone(result.value)

    async def test_17_site_font_packages(self):
        result = await self.client.site_font_packages.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    # def test_18_get_block_download_policy_for_files_data(self):
    #    result = self.client.site.get_block_download_policy_for_files_data().execute_query()
    #    self.assertIsNotNone(result.value)

    # def test_19_get_top_files(self):
    #    result = self.client.site_manager_svc.top_files().execute_query()
    #    self.assertIsNotNone(result.value)
