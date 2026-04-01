import asyncio
import os

from examples.sharepoint.lists.assessment.broken_tax_field_value import fields
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.workbooks.sort_field import WorkbookSortField
from office365.onedrive.workbooks.tables.table import WorkbookTable
from office365.onedrive.workbooks.worksheets.worksheet import WorkbookWorksheet
from tests.graph_case import GraphTestCase


class TestExcelTables(GraphTestCase):
    excel_file = None  # type: DriveItem
    worksheet = None  # type: WorkbookWorksheet
    table = None  # type: WorkbookTable

    @classmethod
    def setUpClass(cls):
        super(TestExcelTables, cls).setUpClass()
        path = "{0}/../../data/Financial Sample.xlsx".format(os.path.dirname(__file__))

        async def _async_setup():
            cls.excel_file = await cls.client.me.drive.root.upload_file(path).execute_query()
            assert cls.excel_file.resource_path is not None
            cls.worksheet = await cls.excel_file.workbook.worksheets["Sheet1"].get().execute_query()
            assert cls.worksheet.resource_path is not None
            cls.table = await cls.worksheet.tables["financials"].get().execute_query()
            assert cls.table.resource_path is not None

        asyncio.run(_async_setup())

    @classmethod
    def tearDownClass(cls):
        async def _async_teardown():
            await cls.excel_file.delete_object().execute_query_retry()

        asyncio.run(_async_teardown())

    async def test1_sort_apply(self):
        sort_fields = [WorkbookSortField()]
        result = await self.__class__.table.sort.apply(sort_fields).execute_query()
        self.assertIsNotNone(result.resource_path)
