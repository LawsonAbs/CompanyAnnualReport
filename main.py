"""
使用正则表达式： ^\s*\n 替换空行
"""
import logging
import requests
import random
import time
import sys
import os
download_path= 'http://www.cninfo.com.cn/new/index'
saving_path= '/home/lawson/program/CompanyAnnualReport/report/'
User_Agent= [
"Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
"Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
"Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
"Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
"Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
"Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
"Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0"
] #User_Agent的集合

headers= {'Accept': 'application/json, text/javascript, */*; q=0.01',
"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
"Accept-Encoding": "gzip, deflate",
"Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-HK;q=0.6,zh-TW;q=0.5",
'Host': 'www.cninfo.com.cn',
'Origin': 'http://www.cninfo.com.cn',
'Referer': 'http://www.cninfo.com.cn/new/commonUrl?url=disclosure/list/notice',
'X-Requested-With': 'XMLHttpRequest'
}

"""
功能：请求单页的内容
"""
def single_page(page_num,date):
    # 这个url是从 F12 获取到的
    query_path= 'http://www.cninfo.com.cn/new/hisAnnouncement/query'
    headers['User-Agent']= random.choice(User_Agent) #定义User_Agent
    
    # 弄清楚下面每项的参数含义
    query= {
    'pageNum': page_num,
    'pageSize': 30,
    'column': 'szse',   
    'tabName': 'fulltext',
    'plate': '',
    'stock':'' ,
    'searchkey':'',
    'secid':'' ,
    'category': 'category_ndbg_szsh',
    'trade': '',
    'seDate': date, # 设置一个更新
    'sortName':'' ,
    'sortType': '',
    'isHLtitle': 'true',
    }
    namelist= requests.post(query_path,headers = headers,data = query)
    # <class 'requests.models.Response'>    ['announcements']说明这个 namelist.json() 就是一个字典
    # if page_num == 101:
    #     print("2w")
    
    try:
        # if namelist.content != '':
        #     return namelist.json()['announcements']
        return namelist.json()['announcements'] #json中的年度报告信息
    except:
        return [] # 


"""
功能：下载单页中的年报
parameters:
1. single_page: 单个页面中的信息
"""
def download(single_page,year): #下载年报
    try:
        headers= {'Accept': 'application/json, text/javascript, */*; q=0.01',
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-HK;q=0.6,zh-TW;q=0.5",
        'Host': 'www.cninfo.com.cn',
        'Origin': 'http://www.cninfo.com.cn'
        }
        for i in single_page:
            #print(i)
            title = i['announcementTitle']
            aim_1 = year+"年年度报告(更新后)"
            aim_2 = year+"年年度报告"
            if title == aim_1 or title == aim_2:
                url='http://www.cninfo.com.cn/new/announcement/download?bulletinId=' + i['announcementId']+'&announceTime='+i['adjunctUrl'][10:20]
                name= i["secCode"]+ '_' + i['secName']+ '_' + i['announcementTitle']+ '.pdf'

                if '*' in name:
                    name= name.replace('*','')
                file_path= saving_path+ name
                time.sleep(random.random()* 2)
                headers['User-Agent']= random.choice(User_Agent)
                r = requests.get(url,headers = headers)
                with open(file_path,'wb') as f:                
                    f.write(r.content)                    
                    print(f"{name}下载成功")
            else:
                continue
    except:        
        pass


if __name__ == '__main__':
    # step1. 构造时间区域
    #date_range = ['2016-01-01~2016-04-30','2017-01-01~2017-04-30','2018-01-01~2018-04-30','2019-01-01~2019-04-30','2021-01-01~2021-04-30']
    date_range = ['2019-01-01~2019-04-30','2021-01-01~2021-04-30']
    # step2. 构建日志记录器
    log_name = "log"
    logging.basicConfig(level=logging.INFO,
                        filemode='w',
                        filename="./"+log_name)
    logger = logging.getLogger(__name__)
    for date in date_range:
        for cur_page in range(1010):
            logger.info(f"当前正在进行的时间段是：{date},当前的访问页是：{cur_page}")
            page_data= single_page(cur_page,date) # page_data
            year = str(int(date.split("-")[0]) - 1) # 得到年份            
            download(page_data,year)