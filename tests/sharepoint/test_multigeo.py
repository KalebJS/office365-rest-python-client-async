from office365.sharepoint.multigeo.api_versions import MultiGeoApiVersions
from tests.sharepoint.sharepoint_case import SPTestCase


class TestMultiGeo(SPTestCase):
    async def test1_get_api_versions(self):
        result = await MultiGeoApiVersions(self.client).get().execute_query()
        self.assertTrue(result.resource_path)
