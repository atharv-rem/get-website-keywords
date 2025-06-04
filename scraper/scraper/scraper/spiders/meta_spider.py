import scrapy
import openpyxl
import ollama
import json
class MetaSpider(scrapy.Spider):
    name = "meta_spider"

    def extract_url_from_excel(self):
        filename = "C://Users//athar//Downloads//color database.xlsx"
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
        website_urls = []
        for row in sheet.iter_rows(min_row=2):
            url = row[3].value  
            if url:
                website_urls.append(url)

        return website_urls

    def start_requests(self):
        website_urls = self.extract_url_from_excel()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
        }
        for url in website_urls:
            yield scrapy.Request(url=url, callback=self.parse, headers=headers)

    def parse(self, response):
        data = {
            'title': response.xpath('//title/text()').get(),
            'url': response.url,
            'meta_description': response.xpath('//meta[@name="description"]/@content').get(),
            'meta_keywords': response.xpath('//meta[@name="keywords"]/@content').get(),
            'og_image': response.xpath('//meta[@property="og:image"]/@content').get(),
        }
        yield data
    
    

