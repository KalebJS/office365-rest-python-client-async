from office365.intune.devices.device import Device
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestDevices(GraphTestCase):
    device = None  # type: Device

    @requires_delegated_permission("Device.Read.All", "Directory.Read.All", "Directory.ReadWrite.All")
    async def test3_list_devices(self):
        result = await self.client.devices.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    # @requires_delegated_permission("Device.Read.All")
    async def test4_get_delta(self):
        result = await self.client.devices.delta.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    @requires_delegated_permission("Directory.AccessAsUser.All")
    async def test5_create_device(self):
        result = await self.client.devices.add("Test device", "linux", "1").execute_query()
        self.assertIsNotNone(result.resource_path)
        self.__class__.device = result

    @requires_delegated_permission("Directory.AccessAsUser.All")
    async def test6_add_registered_owner(self):
        result = await self.__class__.device.registered_owners.add(self.client.me).execute_query()
        self.assertIsNotNone(result.resource_path)

    @requires_delegated_permission("Device.Read.AllDirectory.Read.All", "Directory.ReadWrite.All")
    async def test7_list_registered_owners(self):
        result = await self.__class__.device.registered_owners.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    @requires_delegated_permission("Directory.AccessAsUser.All")
    async def test8_delete_device(self):
        await self.__class__.device.delete_object().execute_query()
