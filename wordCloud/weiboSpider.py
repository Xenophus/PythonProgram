# coding = utf-8
'''
  微博爬虫程序，获取微博下面的评论并保存到 luhan.txt 文件中
'''

import re
import time
import requests
import codecs

def spider(weiboID):
    commentList = []      #评论列表，暂时存放评论
  
   #传入的参数是微博id列表，循环爬取每一条微博的评论
    for id in weiboID:
        time.sleep(2)       #爬取每条微博时暂停2秒，减小被封ip的概率
    
     #pc端反爬机制太叼了，惹不起，多次尝试后选择了爬取移动端的数据，简单许多
     #通过chorm的开发者工具分析NetWork的xhr发现热门评论的变化规律为 'https://m.weibo.cn/api/comments/show?id=' + 微博id + '&page= 页码
      #那么接下来构建 链接
        url = 'https://m.weibo.cn/api/comments/show?id=' + id + '&page={}'
    
      #构建请求头，伪装成浏览器，Cookie使用的是我的微博账号登陆后的信息
        headers = {
              'User-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
              'Host': 'm.weibo.cn',
              'Accept': 'application/json, text/plain, */*',
              'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
              'Accept-Encoding': 'gzip, deflate, br',
              'Referer': 'https://m.weibo.cn/status/' + id,
              'Cookie': '_T_WM=52817851920; ALF=1559613826; " \
                   "SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WhUoA-n65_nV-zM7BunEn1Z5JpX5K-hUgL.Fo-Neo2NeK27eh-2dJLoI0YLxKqL1heLBoq" \
                   "LxKqL1heLBoqLxKBLBonL12BLxK-L12qLB-qLxKBLBonLBoqLxKMLB.zL1hnLxKqL1KMLBK-t; " \
                   "SUHB=0BL4EAfwmQ6I1B; SSOLoginState=1557023614; MLOGIN=1; XSRF-TOKEN=ea82b1',
              'DNT': '1',
              'Connection': 'keep-alive',
          }
        i = 1
        dataList = ['data', 'hot_data']
        proxies = {'http': 'http://112.91.218.21', 'https': 'http://112.91.218.21'} #代理IP，被封的话可以换，在西刺代理网找免费的IP（很多都不能用）
    
        while True:
            # r = requests.get(url=url.format(i), headers=headers, proxies=proxies) #加代理
            r = requests.get(url=url.format(i), headers=headers)  #不加代理
            if r.status_code == 200:
                try:
                    commentPage = r.json()['data']  #json格式的数据，包括评论者的id、评论时间等一系列信息，我只截取的评论内容（感兴趣的话可以获取其他信息做一些分析）
                    for data in dataList:
                        for i in range(0, len(commentPage[data])):
                            #正则，定位评论内容，并去除颜表情
                            text = re.sub('<.*?>|回复<.*?>:|[\U00010000-\U0010ffff]|[\uD800-\uDBFF][\uDC00-\uDFFF]', '',commentPage[data][i]['text'])
                            if text != '':
                                #将非空评论暂放到评论列表
                                commentList.append(text)
                except:
                    i += 1
                    pass
            else:break
            time.sleep(2)   #爬完后暂停2秒
    
    #将评论保存到文件        
    file = codecs.open('luhan.txt', 'w', 'utf-8')
    file.writelines(commentList)
    
if __name__ == '__main__':
    #通过selenium自动获得的25条微博id列表
    weiboID = ['4367559862762009', '4366340364008196', '4364369351183903', '4362948351242446', '4362766036448691',
               '4361281282187187', '4361010792271464', '4358894317286512', '4356691758414200', '4355872891067131',
               '4354041293106532', '4353110391968687', '4352793928710746', '4352368734889377', '4351133021310015',
               '4349894255389235', '4347659915439421', '4347091997696169', '4344448882293788', '4343530573277369',
               '4341487217635293', '4338521328096857', '4337500439050362', '4337153829599348', '4336126749360667']
    #爬取期间丢失了不少数据，没有找到原因，算了...
    spider(weiboID)
    
