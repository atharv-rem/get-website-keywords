import scrapy
import openpyxl

class MetaSpider(scrapy.Spider):
    filename = "C://Users//athar//Downloads//color database.xlsx"
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    website_url = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        website_url.append(row_values[3])
    
    name = "meta_spider"

    def start_requests(self):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
        }
        for url in self.website_url:
            yield scrapy.Request(url=url, callback=self.parse,headers=headers)

    def parse(self, response):
        data = {
            'url': response.url,
            'title': response.xpath('//title/text()').get(),
            'meta_description': response.xpath('//meta[@name="description"]/@content').get(),
            'meta_keywords': response.xpath('//meta[@name="keywords"]/@content').get(),
            'og_title': response.xpath('//meta[@property="og:title"]/@content').get(),
            'og_description': response.xpath('//meta[@property="og:description"]/@content').get(),
        }
        yield data
