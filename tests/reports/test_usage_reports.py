from tests.graph_case import GraphTestCase


class TestUsageReports(GraphTestCase):
    @classmethod
    def setUpClass(cls):
        super(TestUsageReports, cls).setUpClass()

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_billed_usage_export(self):
        result = await self.client.reports.partners.billing.usage.billed.export(
            "G016907411"
        ).execute_query()
        self.assertIsNotNone(result.value)
