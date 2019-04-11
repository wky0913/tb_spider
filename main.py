import datetime
import time
import requests
import json
import re





def taobao(keyword, pages, select_type, date_):
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36'}
    url = 'https://s.taobao.com/search?q={}&imgfile=&js=1&stats_click=search_radio_all%3A1&initiative_id={}&ie=utf8&sort={}'.format(
        keyword, date_, selections[select_type])
    xlsdata = []
    xlsdata.append(['商品名', '商品图片', '商品链接', '价格',
                    '货源', '运费', '收货量', '商家ID', '店铺名'])
    j = 1
    for i in range(pages):
        try:
            r = requests.get(
                url + '&s={}'.format(str(i * 44)), headers=headers,)
            print(r.text)

            data = re.search(r'g_page_config = (.+);', r.text)  # 捕捉json字符串
            data = json.loads(data.group(1), encoding='utf-8')  # json转dict
            for auction in data['mods']['itemlist']['data']['auctions']:
                print('正在下载第 %d 页,第 %d 条数据...' % (i + 1, j))
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
                j += 1
        # 如果预设的Pages数量超过淘宝中实际能查找到的页数,则跳出For循环
        except Exception as e:
            pass
            # print(e)
    return xlsdata


# 将数据写入Excel中
import xlwt  # 写Excel


def ModifyExcel(data):
    "建立Excel表格"
    Wb_Path = "淘宝数据-" + keyword + '-' + \
        datetime.date.today().strftime('%Y%m%d') + ".xls"
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(keyword, cell_overwrite_ok=True)  # 允许更改表单已有内容
    for j in range(len(data)):
        values = data[j]
        for k in range(len(values)):
            sheet.write(j, k, values[k])
    wbk.save(Wb_Path)


if __name__ == '__main__':
    start_time = time.clock()
    selections = {'0': 'default',
                  '1': 'renqi-desc',
                  '2': 'sale-desc'}
    keyword = input('输入商品名\n')
    pages = int(input('爬多少页\n'))
    date_ = 'staobaoz_' + str(datetime.date.today()).replace('-', '')
    select_type = str(input('输入0按默认，输入1按人气，输入2按销量\n'))
    xdata = taobao(keyword, pages, select_type, date_)
    ModifyExcel(xdata)
    end_time = time.clock()
    print('已完成 %s 项产品的下载,共计花费: %s 秒' %
          (len(xdata), int(end_time - start_time))
          )
    input('Tip: Press Enter to close window!')
