from dotenv import load_dotenv
from dataforseo import DataForSEOClient

load_dotenv()

client = DataForSEOClient()

# SERP – live Google organic results
result = client.serp_google_organic_live("international marketing strategy", depth=10)
tasks = result.get("tasks", [])
if tasks and tasks[0].get("result"):
    for item in tasks[0]["result"][0].get("items", [])[:5]:
        print(f"[{item.get('rank_absolute')}] {item.get('title')} — {item.get('url')}")

# Search volume for a list of keywords
volumes = client.search_volume(["digital marketing", "SEO strategy", "startup growth"])
print(volumes)

# Backlinks summary for a domain
backlinks = client.backlinks_summary("ymove.app")
print(backlinks)

# Domain organic keyword overview
overview = client.domain_rank_overview("ymove.app")
print(overview)
