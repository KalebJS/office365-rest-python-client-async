import asyncio

from office365.directory.groups.group import Group
from office365.planner.plans.plan import PlannerPlan
from tests import create_unique_name
from tests.graph_case import GraphTestCase


class TestPlanner(GraphTestCase):
    target_group = None  # type: Group
    target_plan = None  # type: PlannerPlan

    @classmethod
    def setUpClass(cls):
        super(TestPlanner, cls).setUpClass()
        # Ensure Group for a Planner
        filter_str = "groupTypes/any(a:a eq 'unified')"

        async def _async_setup():
            groups = await cls.client.groups.filter(filter_str).get().execute_query()
            cls.target_group = groups[0]
            cls.assertIsNotNone(cls.target_group.resource_path, "Group not found!")

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        pass

    async def test1_create_plan(self):
        plan_name = create_unique_name("My Plan")
        group = self.client.groups.get_by_name("My Sample Team")
        new_plan = await self.client.planner.plans.add(plan_name, group).execute_query()
        self.assertIsNotNone(new_plan.id)
        self.__class__.target_plan = new_plan

    async def test2_get_plan_details(self):
        result = await self.__class__.target_plan.details.get().execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test3_list_my_plans(self):
        my_plans = await self.client.me.planner.plans.get().execute_query()
        self.assertIsNotNone(my_plans.resource_path)
        self.assertGreaterEqual(len(my_plans), 0)

    async def test4_create_task(self):
        task = await self.client.planner.tasks.add("Update client list", self.__class__.target_plan).execute_query()
        self.assertIsNotNone(task.resource_path)

    async def test5_list_tasks(self):
        tasks = await self.__class__.target_plan.tasks.get().execute_query()
        self.assertGreaterEqual(len(tasks), 0)

    async def test6_delete_plan(self):
        plan_to_del = self.__class__.target_plan
        await plan_to_del.delete_object().execute_query()
