import asyncio

from office365.outlook.mail.attachments.attachment_item import AttachmentItem
from office365.outlook.mail.attachments.attachment_type import AttachmentType
from office365.outlook.mail.messages.message import Message
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestAttachments(GraphTestCase):
    target_message = None  # type: Message

    @classmethod
    def setUpClass(cls):
        super(TestAttachments, cls).setUpClass()

        async def _async_setup():
            cls.target_message = await cls.client.me.messages.add(
                subject="Meet for lunch?",
                body="The new cafeteria is open.",
                to_recipients=["fannyd@contoso.onmicrosoft.com"],
            ).execute_query()

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_message.delete_object().execute_query()

        asyncio.run(_async_teardown())

    @requires_delegated_permission("Files.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All")
    async def test1_create_upload_session(self):
        message = self.__class__.target_message
        attachment_item = AttachmentItem(attachment_type=AttachmentType.file, name="flower", size=3483322)
        result = await message.attachments.create_upload_session(attachment_item).execute_query()
        self.assertIsNotNone(result.value)
