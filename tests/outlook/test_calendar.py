from datetime import datetime, timedelta
from typing import Optional

from office365.outlook.calendar.calendar import Calendar
from office365.outlook.calendar.email_address import EmailAddress
from tests import create_unique_name, test_user_principal_name
from tests.decorators import requires_delegated_permission
from tests.graph_case import GraphTestCase


class TestCalendar(GraphTestCase):
    """Tests for Calendar"""

    cal_name = create_unique_name("Volunteer")
    target_cal = None  # type: Optional[Calendar]

    @requires_delegated_permission(
        "Calendars.Read.Shared", "Calendars.ReadWrite.Shared"
    )
    async def test1_find_my_meeting_times(self):
        result = await self.client.me.find_meeting_times().execute_query()
        self.assertIsNotNone(result.value.meetingTimeSuggestions)

    @requires_delegated_permission(
        "Calendars.ReadBasic",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
    )
    async def test2_get_my_schedule(self):
        end_time = datetime.utcnow()
        start_time = end_time - timedelta(days=7)
        result = await self.client.me.calendar.get_schedule(
            schedules=[test_user_principal_name],
            start_time=start_time,
            end_time=end_time,
        ).execute_query()
        self.assertIsNotNone(result.value)

    @requires_delegated_permission(
        "Calendars.ReadBasic",
        "Calendars.ReadWrite",
        "Calendars.Read",
        "Calendars.ReadWrite.Shared",
    )
    async def test3_list_my_cal_groups(self):
        cal_groups = await self.client.me.calendar_groups.get().execute_query()
        self.assertIsNotNone(cal_groups.resource_path)

    @requires_delegated_permission(
        "Calendars.ReadBasic",
        "Calendars.ReadWrite",
        "Calendars.Read",
        "Calendars.ReadWrite.Shared",
    )
    async def test4_list_my_cal_permissions(self):
        result = (
            await self.client.me.calendar.calendar_permissions.get().execute_query()
        )
        self.assertIsNotNone(result.resource_path)

    @requires_delegated_permission(
        "Calendars.ReadBasic", "Calendars.Read", "Calendars.ReadWrite"
    )
    async def test5_list_my_cal_view(self):
        end_time = datetime.utcnow()
        start_time = end_time - timedelta(days=14)
        result = await self.client.me.get_calendar_view(
            start_dt=start_time, end_dt=end_time
        ).execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test6_get_my_reminder_view(self):
        end_time = datetime.utcnow()
        start_time = end_time - timedelta(days=14)
        result = await self.client.me.get_reminder_view(
            start_dt=start_time, end_dt=end_time
        ).execute_query()
        self.assertIsNotNone(result.value)

    async def test7_list_my_events(self):
        result = await self.client.me.calendar.events.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test6_get_my_calendars(self):
        result = await self.client.me.calendars.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test8_create_cal(self):
        result = await self.client.me.calendars.add(name=self.cal_name).execute_query()
        self.assertIsNotNone(result.resource_path)
        self.__class__.target_cal = result

    async def test9_update_cal(self):
        cal = self.__class__.target_cal
        self.__class__.cal_name = self.cal_name + "_Updated"
        await cal.set_property("name", self.cal_name).update().execute_query()

    async def test_10_get_cal(self):
        cal_id = self.__class__.target_cal.id
        result = await (
            self.client.me.calendars[cal_id]
            .select(["name", "owner"])
            .get()
            .execute_query()
        )
        self.assertEqual(result.name, self.cal_name)
        self.assertIsInstance(result.owner, EmailAddress)

    async def test_11_delete_cal(self):
        cal = self.__class__.target_cal
        await cal.delete_object().execute_query()

    async def test_12_allowed_calendar_sharing_roles(self):
        result = await self.client.me.calendar.allowed_calendar_sharing_roles(
            test_user_principal_name
        ).execute_query()
        self.assertIsNotNone(result.value)
