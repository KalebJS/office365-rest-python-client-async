from office365.sharepoint.publishing.pages.page import SitePage
from tests.sharepoint.sharepoint_case import SPTestCase


class TestMultilingual(SPTestCase):
    """"""

    site_page = None  # type: SitePage

    async def test1_is_web_multilingual(self):
        web = await (
            self.client.web.select(["IsMultilingual", "SupportedUILanguageIds"])
            .expand(["MultilingualSettings"])
            .get()
            .execute_query()
        )
        self.assertIsNotNone(web.is_multilingual)
        self.assertIsNotNone(web.supported_ui_language_ids)
        self.assertIsNotNone(web.multilingual_settings)

    async def test2_create_page(self):
        page_title = "My Page"
        site_page = await self.client.site_pages.create_page(
            page_title, language="en-us"
        ).execute_query()
        self.assertIsNotNone(site_page.resource_path)
        self.__class__.site_page = site_page

    async def test3_get_page_language(self):
        site_page = (
            await self.__class__.site_page.get().select(["Language"]).execute_query()
        )
        self.assertIsNotNone(site_page.language)

    # The Machine Translations Service API will no longer be supported as of the end of July 2022
    # def test4_get_page_language(self):
    #    from office365.sharepoint.translation.job import TranslationJob
    #    job = TranslationJob.is_service_enabled(self.client, "en").execute_query()
    #    self.assertIsNotNone(job.value)

    # def test5_export_items_variations(self):
    #    from office365.sharepoint.translation.variations_timer_job import (
    #        VariationsTranslationTimerJob,
    #    )
    #
    #    result = VariationsTranslationTimerJob.export_items(
    #        self.client, "/sites/team/SitePages", [1, 2, 3]
    #    ).execute_query()
    #    self.assertIsNotNone(result.resource_path)
