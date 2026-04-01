import asyncio
import uuid

from office365.teams.team import Team
from tests.graph_case import GraphTestCase


class TestTeamApps(GraphTestCase):
    """Tests for team Apps"""

    target_team = None  # type: Team

    @classmethod
    def setUpClass(cls):
        super(TestTeamApps, cls).setUpClass()
        team_name = "Team_" + uuid.uuid4().hex

        async def _async_setup():
            new_team = (
                await cls.client.teams.create(team_name).get().execute_query_retry()
            )
            cls.target_team = new_team

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.target_team.delete_object().execute_query_retry()

        asyncio.run(_async_teardown())

    async def test1_get_team_apps(self):
        apps = await self.__class__.target_team.installed_apps.get().execute_query()
        self.assertIsNotNone(apps.resource_path)
