import asyncio
import os

from tests import test_team_site_url
from tests.graph_case import GraphTestCase


class TestShares(GraphTestCase):
    """Shares API specific test case"""

    @classmethod
    def setUpClass(cls):
        super(TestShares, cls).setUpClass()
        path = "{0}/../data/Financial Sample.xlsx".format(os.path.dirname(__file__))

        async def _async_setup():
            cls.file_item = await (
                cls.client.sites.get_by_url(test_team_site_url)
                .drive.root.upload_file(path)
                .execute_query()
            )
            assert cls.file_item.resource_path is not None

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.file_item.delete_object().execute_query_retry()

        asyncio.run(_async_teardown())

    async def test1_get_file_by_abs_url(self):
        file_abs_url = "{0}/Shared Documents/Financial Sample.xlsx".format(
            test_team_site_url
        )
        result = (
            await self.client.shares.by_url(file_abs_url)
            .drive_item.get()
            .execute_query()
        )
        self.assertIsNotNone(result.resource_path)
        self.assertTrue(result.is_file)
