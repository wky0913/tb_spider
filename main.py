import datetime
import time
import requests
import json
import re
import xlwt
import pandas as pd
import logging


logger = logging.getLogger(__name__)
logger.setLevel(level = logging.INFO)
handler = logging.FileHandler("log.txt")
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)



class WedDownload(object):
    def __init__(self):
        self.headers = {'user-agent': \
                        'Mozilla/5.0 (Windows NT 10.0;WOW64) \
                        AppleWebKit/537.36 (KHTML, like Gecko) \
                        Chrome/63.0.3239.108 Safari/537.36'}
        self.url = ''

    def download(self, url):
        resp = requests.get(url,headers=self.headers)
        if not resp.ok:
            logger.info("WedDownload %s" % resp.text)
        return resp

    def tb_download(self, se_key, pages, se_type, se_date):
        xlsdata = []
        xlsdata.append(['商品名', '商品图片', '商品链接', '价格',
                        '货源', '运费', '收货量', '商家ID', '店铺名'])
        selections = {'0': 'default',
                      '1': 'renqi-desc',
                      '2': 'sale-desc'}
        url = 'https://s.taobao.com/search?q={}&imgfile=&js=1&stats_click=search_radio_all%3A1&initiative_id={}&ie=utf8&sort={}'.format(
            se_key, se_date, selections[se_type])

        df = pd.DataFrame(columns=[])
        for i in range(pages):
            print("开始下载第"+str(i+1)+"页")
            time.sleep(5)
            resp = self.download(url + '&s={}'.format(str(i * 44)))
            data = re.search(r'g_page_config = (.+);', resp.text)
            data = json.loads(data.group(1), encoding='utf-8')
            for auction in data['mods']['itemlist']['data']['auctions']:
                xlsdata.append([
                    auction['raw_title'],    # 商品名
                    'https:' + auction['pic_url'],  # 商品图片
                    'https://item.taobao.com/item.htm?ft=t&id=' + \
                    auction['nid'],          # 商品链接
                    auction['view_price'],   # 价格
                    auction['item_loc'],     # 货源
                    auction['view_fee'],    # 运费
                    auction['view_sales'].replace('人收货', ''),   # 卖出数量
                    auction['user_id'],      # 商家id
                    auction['nick']          # 店铺名
                ])

            print(data['mods']['itemlist']['data']['auctions'][0])
            x=re.findall('"auctions":(.*?),"recommendAuctions"',resp.text)
            # print(data)
            # for xx in x:
            #     print('===========================')
            #     print(xx)

            # df=pd.read_json(data)
            # print(df)
        return xlsdata


def ModifyExcel(data):
    filename = se_key + '-' + datetime.date.today().strftime('%Y%m%d') + ".xls"
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(se_key, cell_overwrite_ok=True)
    for j in range(len(data)):
        values = data[j]
        for k in range(len(values)):
            sheet.write(j, k, values[k])
    wbk.save(filename)

if __name__ == '__main__':
    # se_key = input('输入商品名\n')
    # pages = int(input('爬多少页\n'))
    # se_type = str(input('输入0按默认，输入1按人气，输入2按销量\n'))
    se_key = '沙发'
    pages = 1
    se_type = '2'
    se_date = 'staobaoz_' + str(datetime.date.today()).replace('-', '')
    xdata = WedDownload().tb_download(se_key, pages, se_type, se_date)
    ModifyExcel(xdata)

    print("数据下载完成")
