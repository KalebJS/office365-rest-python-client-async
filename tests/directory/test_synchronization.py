import asyncio

from office365.directory.serviceprincipals.service_principal import ServicePrincipal
from tests import test_client_id
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestSynchronization(GraphTestCase):
    target_sp = None  # type: ServicePrincipal

    # "salesforce"

    @classmethod
    def setUpClass(cls):
        super(TestSynchronization, cls).setUpClass()

        async def _async_setup():
            cls.target_sp = await (
                cls.client.service_principals.get_by_app_id(test_client_id)
                .get()
                .execute_query()
            )

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        pass

    @requires_delegated_permission(
        "Synchronization.Read.All", "Synchronization.ReadWrite.All"
    )
    async def test1_list_synchronization_jobs(self):
        result = await self.target_sp.synchronization.jobs.get().execute_query()
        self.assertIsNotNone(result.resource_path)
