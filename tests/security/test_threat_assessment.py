from tests.graph_case import GraphTestCase


class TestThreatAssessment(GraphTestCase):
    threat_assessment_request = None

    async def test1_create_url_assessment(self):
        result = await self.client.information_protection.create_url_assessment(
            "http://test.com", "block", "phishing"
        ).execute_query()
        self.assertIsNotNone(result.resource_path)
        self.__class__.threat_assessment_request = result

    async def test2_create_file_assessment(self):
        result = await self.client.information_protection.create_file_assessment(
            "test.txt", "VGhpcyBpcyBhIHRlc3QgZmlsZQ==", "block", "malware"
        ).execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test3_create_email_file_assessment(self):
        result = await self.client.information_protection.create_email_file_assessment(
            "tifc@contoso.com", "VGhpcyBpcyBhIHRlc3QgZmlsZQ==", "block", "malware"
        ).execute_query()
        self.assertIsNotNone(result.resource_path)

    async def test4_list_threat_assessment_requests(self):
        col = (
            await self.client.information_protection.threat_assessment_requests.get().execute_query()
        )
        self.assertIsNotNone(col.resource_path)
