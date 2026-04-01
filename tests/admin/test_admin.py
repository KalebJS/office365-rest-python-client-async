from tests.graph_case import GraphTestCase


class TestAdmin(GraphTestCase):
    """SharePoint specific test case base class"""

    async def test1_get_sharepoint_settings(self):
        result = await self.client.admin.sharepoint.settings.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test2_update_sharepoint_settings(self):
        settings = self.client.admin.sharepoint.settings
        settings.sharing_blocked_domain_list = ["contoso.com", "fabrikam.com"]
        await settings.update().execute_query()

    async def test3_list_issues(self):
        result = await self.client.admin.service_announcement.issues.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test4_list_microsoft365_apps(self):
        result = await self.client.admin.microsoft365_apps.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    # def test5_get_admin_people(self):
    #    result = self.client.admin.people.get().execute_query()
    #    self.assertIsNotNone(result.resource_path)
