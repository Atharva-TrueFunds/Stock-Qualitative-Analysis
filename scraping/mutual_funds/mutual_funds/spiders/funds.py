import scrapy
import pandas as pd


class FundsSpider(scrapy.Spider):
    name = "funds"
    allowed_domains = ["upstox.com"]
    start_urls = [
        "https://upstox.com/mutual-funds/quant-small-cap-fund-growth-option-direct-plan-105739?isin=INF966L01689"
    ]

    def parse(self, response):
        try:
            table = response.css(".holdingsList")

            if table:
                rows = table.css("tr")
                data = []

                for row in rows:
                    columns = row.css("th")
                    row_data = [
                        column.css("::text").get().strip() for column in columns
                    ]
                    data.append(row_data)

                df = pd.DataFrame(data[1:], columns=data[0])

                df.to_excel("funds_data.xlsx", index=False)
                self.logger.info("Data written to funds_data.xlsx")
        except Exception as e:
            self.logger.error("Error parsing table: %s", e)
