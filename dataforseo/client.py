import os
import requests
from requests.auth import HTTPBasicAuth
from typing import Any


BASE_URL = "https://api.dataforseo.com/v3"


class DataForSEOClient:
    def __init__(self, login: str = None, password: str = None):
        self.login = login or os.environ["DATAFORSEO_LOGIN"]
        self.password = password or os.environ["DATAFORSEO_PASSWORD"]
        self._session = requests.Session()
        self._session.auth = HTTPBasicAuth(self.login, self.password)
        self._session.headers.update({"Content-Type": "application/json"})

    def _get(self, path: str) -> dict:
        response = self._session.get(f"{BASE_URL}/{path.lstrip('/')}")
        response.raise_for_status()
        return response.json()

    def _post(self, path: str, payload: list[dict]) -> dict:
        response = self._session.post(f"{BASE_URL}/{path.lstrip('/')}", json=payload)
        response.raise_for_status()
        return response.json()

    # ── SERP ──────────────────────────────────────────────────────────────────

    def serp_google_organic_live(
        self,
        keyword: str,
        location_code: int = 2840,  # United States
        language_code: str = "en",
        depth: int = 10,
        **kwargs: Any,
    ) -> dict:
        payload = [{"keyword": keyword, "location_code": location_code,
                    "language_code": language_code, "depth": depth, **kwargs}]
        return self._post("serp/google/organic/live/advanced", payload)

    def serp_task_post(self, keyword: str, location_code: int = 2840,
                       language_code: str = "en", **kwargs: Any) -> dict:
        payload = [{"keyword": keyword, "location_code": location_code,
                    "language_code": language_code, **kwargs}]
        return self._post("serp/google/organic/task_post", payload)

    def serp_task_get(self, task_id: str) -> dict:
        return self._get(f"serp/google/organic/task_get/advanced/{task_id}")

    # ── Keywords Data ──────────────────────────────────────────────────────────

    def keywords_for_site(self, target: str, location_code: int = 2840,
                          language_code: str = "en", **kwargs: Any) -> dict:
        payload = [{"target": target, "location_code": location_code,
                    "language_code": language_code, **kwargs}]
        return self._post("keywords_data/google_ads/keywords_for_site/live", payload)

    def search_volume(self, keywords: list[str], location_code: int = 2840,
                      language_code: str = "en", **kwargs: Any) -> dict:
        payload = [{"keywords": keywords, "location_code": location_code,
                    "language_code": language_code, **kwargs}]
        return self._post("keywords_data/google_ads/search_volume/live", payload)

    def keyword_suggestions(self, keyword: str, location_code: int = 2840,
                            language_code: str = "en", limit: int = 100,
                            **kwargs: Any) -> dict:
        payload = [{"keyword": keyword, "location_code": location_code,
                    "language_code": language_code, "limit": limit, **kwargs}]
        return self._post("dataforseo_labs/google/keyword_suggestions/live", payload)

    # ── Backlinks ──────────────────────────────────────────────────────────────

    def backlinks_summary(self, target: str, **kwargs: Any) -> dict:
        payload = [{"target": target, **kwargs}]
        return self._post("backlinks/summary/live", payload)

    def backlinks_list(self, target: str, limit: int = 100, **kwargs: Any) -> dict:
        payload = [{"target": target, "limit": limit, **kwargs}]
        return self._post("backlinks/backlinks/live", payload)

    # ── Domain Analytics ──────────────────────────────────────────────────────

    def domain_rank_overview(self, target: str, location_code: int = 2840,
                             language_code: str = "en", **kwargs: Any) -> dict:
        payload = [{"target": target, "location_code": location_code,
                    "language_code": language_code, **kwargs}]
        return self._post("dataforseo_labs/google/domain_rank_overview/live", payload)

    def domain_organic_keywords(self, target: str, location_code: int = 2840,
                                language_code: str = "en", limit: int = 100,
                                **kwargs: Any) -> dict:
        payload = [{"target": target, "location_code": location_code,
                    "language_code": language_code, "limit": limit, **kwargs}]
        return self._post("dataforseo_labs/google/ranked_keywords/live", payload)

    # ── On-Page ────────────────────────────────────────────────────────────────

    def onpage_task_post(self, target: str, max_crawl_pages: int = 10,
                         **kwargs: Any) -> dict:
        payload = [{"target": target, "max_crawl_pages": max_crawl_pages, **kwargs}]
        return self._post("on_page/task_post", payload)

    def onpage_summary(self, task_id: str) -> dict:
        return self._get(f"on_page/summary/{task_id}")

    # ── Business Data ──────────────────────────────────────────────────────────

    def google_my_business_info(self, keyword: str, location_code: int = 2840,
                                language_code: str = "en", **kwargs: Any) -> dict:
        payload = [{"keyword": keyword, "location_code": location_code,
                    "language_code": language_code, **kwargs}]
        return self._post("business_data/google/my_business_info/live", payload)
