import scrapy


class HoldingsSpider(scrapy.Spider):
    name = "holdings"
    allowed_domains = ["upstox.com"]
    start_urls = [
        "https://upstox.com/mutual-funds/quant-small-cap-fund-growth-option-direct-plan-105739?isin=INF966L01689"
    ]

    def parse(self, response):
        holdings = response.css(".holdingsListItem")

        for holding in holdings:
            name = holding.css("h3::text").get()
            percentage = holding.css(".listValue::text").get()
            print(f"Name: {name}, Percentage: {percentage}")

        self.logger.info("Data printed to terminal")
