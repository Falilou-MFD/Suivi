import reflex as rx

config = rx.Config(
    app_name="suivi_app",
    api_url="http://localhost:8000",
    plugins=["reflex_base.plugins.sitemap.SitemapPlugin"],
)