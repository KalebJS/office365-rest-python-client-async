from unittest import TestCase

from office365.graph_client import GraphClient
from tests import test_client_id, test_client_secret, test_tenant


class TestTenant(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = GraphClient(tenant=test_tenant).with_client_secret(test_client_id, test_client_secret)

    async def test1_find_tenant_information(self):
        result = await self.client.tenant_relationships.find_tenant_information_by_domain_name(
            test_tenant
        ).execute_query()
        self.assertIsNotNone(result.value)
