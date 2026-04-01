from datetime import datetime, timedelta

from office365.sharepoint.lists.list import List
from office365.sharepoint.webhooks.subscription import Subscription
from tests.sharepoint.sharepoint_case import SPTestCase


class TestSPWebHooks(SPTestCase):
    target_list = None  # type: List
    target_subscription = None  # type: Subscription
    push_service_url = "https://westeurope0.pushnp.svc.ms/notifications?token=526a9d28-d4ec-45b7-81b9-4e1599524784"

    @classmethod
    def setUpClass(cls):
        super(TestSPWebHooks, cls).setUpClass()
        cls.target_list = cls.client.web.lists.get_by_title("Documents")

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_create_subscription(self):
        subscription = await self.target_list.subscriptions.add(self.push_service_url).execute_query()
        self.assertIsNotNone(subscription.notification_url)
        self.__class__.target_subscription = subscription

    async def test2_list_webhooks(self):
        subscriptions = await self.target_list.subscriptions.get().execute_query()
        self.assertGreater(len(subscriptions), 0)

    async def test3_update_subscription(self):
        subscription = self.__class__.target_subscription
        subscription.expiration_datetime = datetime.utcnow() + timedelta(days=10)
        await subscription.update().execute_query()

    async def test4_delete_subscription(self):
        subscription = self.__class__.target_subscription
        await subscription.delete_object().execute_query()
