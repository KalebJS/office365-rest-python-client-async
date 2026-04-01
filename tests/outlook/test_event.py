from datetime import datetime, timedelta

from office365.outlook.calendar.events.event import Event
from tests import test_user_principal_name
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestOutlookEvent(GraphTestCase):
    target_event = None  # type: Event

    @requires_delegated_permission("Calendars.ReadWrite", "Calendars.ReadWrite.Shared")
    async def test2_create_event(self):
        when = datetime.now() + timedelta(days=1)
        new_event = await self.client.me.calendar.events.add(
            subject="Let's go for lunch",
            body="Does mid month work for you?",
            start=when,
            end=when + timedelta(hours=1),
            attendees=[test_user_principal_name],
        ).execute_query()
        self.assertIsNotNone(new_event.id)
        self.__class__.target_event = new_event

    @requires_delegated_permission(
        "Calendars.ReadBasic",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
    )
    async def test3_list_my_events(self):
        events = await self.client.me.events.get().execute_query()
        self.assertGreaterEqual(len(events), 1)

    @requires_delegated_permission("Calendars.ReadWrite", "Calendars.ReadWrite.Shared")
    async def test4_update_event(self):
        event = self.__class__.target_event
        event.subject = "Let's go for lunch (updated)"
        await event.update().execute_query()

    # def test5_cancel_event(self):
    #    event = self.__class__.target_event
    #    event.cancel().execute_query()

    @requires_delegated_permission("Calendars.ReadWrite", "Calendars.ReadWrite.Shared")
    async def test6_delete_event(self):
        event_to_delete = self.__class__.target_event
        await event_to_delete.delete_object().execute_query()
        # verify
        events = await self.client.me.events.get().execute_query()
        results = [e for e in events if e.id == event_to_delete.id]
        self.assertEqual(len(results), 0)
