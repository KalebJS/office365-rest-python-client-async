from office365.sharepoint.sitescripts.metadata import SiteScriptMetadata
from office365.sharepoint.sitescripts.utility import SiteScriptUtility
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSiteScript(SPTestCase):
    site_script_meta = None  # type: SiteScriptMetadata
    site_script_count = None

    async def test_1_create(self):
        script = {
            "$schema": "schema.json",
            "actions": [{"verb": "applyTheme", "themeName": "Contoso Theme"}],
            "bindata": {},
            "version": 1,
        }

        result = await SiteScriptUtility.create_site_script(
            self.client, "Contoso theme script", "", script
        ).execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.site_script_meta = result.value

    async def test_2_list(self):
        result = await SiteScriptUtility.get_site_scripts(self.client).execute_query()
        self.assertIsNotNone(result.value)
        self.__class__.site_script_count = len(result.value)

    async def test_3_delete(self):
        await SiteScriptUtility.delete_site_script(
            self.client, self.site_script_meta.Id
        ).execute_query()
        result_after = await SiteScriptUtility.get_site_scripts(
            self.client
        ).execute_query()
        self.assertEqual(self.site_script_count - 1, len(result_after.value))
