from uuid import UUID

from office365.sharepoint.sitedesigns.creation_info import SiteDesignCreationInfo
from office365.sharepoint.sitedesigns.metadata import SiteDesignMetadata
from office365.sharepoint.sitescripts.utility import SiteScriptUtility
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSiteDesign(SPTestCase):
    site_design_metadata = None  # type: SiteDesignMetadata
    site_design_count = None

    async def test_1_create(self):
        info = SiteDesignCreationInfo(
            title="Contoso customer tracking",
            description="Creates customer list and applies standard theme",
            site_script_ids=[UUID("07702c07-0485-426f-b710-4704241caad9")],
            web_template="64",
        )
        result = await SiteScriptUtility.create_site_design(
            self.client, info
        ).execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.site_design_metadata = result.value

    async def test_2_list(self):
        result = await SiteScriptUtility.get_site_designs(self.client).execute_query()
        self.assertIsNotNone(result.value)
        self.assertGreater(len(result.value), 0)
        self.__class__.site_design_count = len(result.value)

    async def test_3_delete(self):
        await SiteScriptUtility.delete_site_design(
            self.client, self.site_design_metadata.Id
        ).execute_query()
        result = await SiteScriptUtility.get_site_designs(self.client).execute_query()
        self.assertEqual(self.site_design_count - 1, len(result.value))
