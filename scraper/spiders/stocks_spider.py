from lxml import etree
import openpyxl
import scrapy

class StocksSpider(scrapy.Spider):
    '''
    This class is the initial spider which will start the scraping.
    Returns links to each specific stock
    '''
    name = "stocks"
    data = []

    def start_requests(self):
        urls = ['http://www.moneycontrol.com/india/stockpricequote/']
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):

        table = response.css('table.pcq_tbl')
        rows = table.xpath('.//td//a')

        #Extracting all the links for all the stocks here
        for row in rows:
            link = row.xpath('@href')
            name = row.xpath('text()')
            if row.xpath('@href').extract() and row.xpath('text()').extract():
                args = (link.extract(), name.extract())
                self.data.append(args)

        #Scraping each link here
        for datum in self.data:
            yield scrapy.Request(url=datum[0][0], callback=self.parse_result)

    def parse_result(self, response):
        '''
        Second callback triggered by each link

        Yields object containing name, price, price_change, price_change_perc,
        div_yield, p_e
        '''
        name = response.css('h1.b_42').xpath('.//text()').extract()[0]
        curr_price = response.css('span#Bse_Prc_tick').xpath('.//strong/text()').extract()[0]
        prc_chg = response.css('div#b_changetext').xpath('.//span//strong/text()').extract()[0]
        prc_chg_perc = response.css('div#b_changetext').xpath('./text()').extract()[1].strip()
        div_yield = ''
        p_e = ''

        #Extracting div yield here and P/E ratio here
        data = response.css('div#mktdet_2').css('div.PA7')
        for datum in data:
            if datum.xpath('.//div[contains(., "DIV YIELD")]'):
                div_yield = datum.css('div.FR::text').extract()[0]
            if datum.xpath('.//div/text()').re(r'^P/E'):
                p_e = datum.css('div.FR::text').extract()[0]


        yield {
            'name': name,
            'price': curr_price,
            'price_change': prc_chg,
            'price_change_perc': prc_chg_perc,
            'div_yield': div_yield,
            'p_e': p_e
        }
