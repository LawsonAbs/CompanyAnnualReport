import copy
from pdfminer.high_level import extract_text
import json
import xlwt
import os 
import sys


"""
根据日志文件将结果写到xls文件中
"""
def anslysis_log2_xls(file_path):
    with open(file_path,'r') as f:
        line = f.readline()
        line = f.readline() 
        temp = line.strip("\n").replace("开始分析第","")
        year = temp.replace("年的数据","")

        cur_report_name = ""
        pre_value = 0 
        cur_value = 0 # key_word_map 的值

        pre_value_acc = 0
        cur_value_acc = 0 # # key_word_map_acc 中的值
        res = [] # 存储最后的结果        
        pre_key_word_map_acc = {}  #上一次的累计值
        key_word_difference_value = {}
        # cur_key_map_acc = {'一带一路': 0, '供给侧': 0, '八项规定': 0, '利率市场化': 0, '反腐败': 0, '和谐社会': 0, '国内生产总值': 0, '国际经济': 0,'非公有制经济':0,'国内外经济形势':0,'宏观调控':0,'国民经济':0,'货币政策':0,'通货膨胀':0,'国际经济环境':0,'结构性改革':0,
        # '中国制造“2025”':0,'京津冀协同发展':0,'海上丝绸之路':0,'产能过剩':0,'亚太自贸区':0,'长江经济带':0,'混合所有制经济':0,
        # '利率市场化':0,'高质量发展':0,'美丽中国':0,'经济合作区':0 } 
        cur_key_map_acc = {'数字化': 0, '数字经济': 0, '产业数字化': 0, '数字产业化': 0, '智能制造': 0, '产业升级': 0, '供应链升级': 0, '智慧物流': 0, '智能化': 0, '互联网': 0, '大数据': 0, '云计算': 0, '人工智能': 0, '电子商务': 0, '云平台': 0, '云服务': 0,'中国制造2025': 0, '移动互联网': 0, '数据分析': 0, '数据挖掘': 0, '智能化': 0, '信息化': 0, '网络销售': 0} # 累计，最后必须要返回的项
        while(line):
            line = f.readline()
            line = line.strip("\n")
            #print(line)
            if line.startswith("key_word_map={"):
                line = line.replace("key_word_map=","")
                line = line.replace("\'","\"")
                # print(line)
                cur_key_map = json.loads(line)
                for item in cur_key_map.items():
                    key,value = item
                    cur_value += value
            elif line.startswith("key_word_map_accumulate"):
                line = line.replace("key_word_map_accumulate=","")
                line = line.replace("\'","\"")
                # print(line)
                
                pre_key_word_map_acc = cur_key_map_acc.copy()
                cur_key_map_acc = json.loads(line)
                for item in cur_key_map_acc.items():
                    key,value =item
                    cur_value_acc += value                    
                    key_word_difference_value[key] = cur_key_map_acc[key] - pre_key_word_map_acc[key]                    
                res.append([cur_report_name,cur_value-pre_value,cur_value_acc-pre_value_acc,key_word_difference_value.copy()])
                pre_value = cur_value
                cur_value = 0
                pre_value_acc = cur_value_acc
                cur_value_acc = 0
            elif line.startswith("/home"):                    
                line = line.split("/")[-1]
                cur_report_name = line.split(".")[0] 
                cur_report_name = cur_report_name.replace(f"{year}年年度报告","")
            
            # 写入文件
            elif(line.startswith("将结果写入到文件中")):
                out_path = "/home/lawson/program/CompanyAnnualReport/" +f"{year}.xls"
                xls = xlwt.Workbook()
                try:
                    #生成excel的方法，声明excel                    
                    sheet = xls.add_sheet(sheetname='sheet1',cell_overwrite_ok=True)
                    row = 0   #在excel开始写的位置（y）
                    column = 0
                    
                    # head = ["公司名称",'是否出现','出现次数累计','一带一路', '供给侧', '八项规定', '利率市场化', '反腐败', '和谐社会', '国内生产总值', '国际经济','非公有制经济','国内外经济形势','宏观调控','国民经济','货币政策','通货膨胀','国际经济环境','结构性改革','中国制造“2025”','京津冀协同发展','海上丝绸之路','产能过剩','亚太自贸区','长江经济带','混合所有制经济','利率市场化','高质量发展','美丽中国','经济合作区']

                    head = ['公司名称','是否出现','出现次数累计','数字化' , '数字经济' , '产业数字化' , '数字产业化' , '智能制造' , '产业升级' , '供应链升级' , '智慧物流' , '智能化' , '互联网' , '大数据' , '云计算' , '人工智能' , '电子商务' , '云平台' , '云服务' ,'中国制造2025' , '移动互联网' , '数据分析' , '数据挖掘' , '智能化' , '信息化' , '网络销售' ]
                    for column in range(len(head)):
                        sheet.write(row,column,head[column])

                    row+=1 #另起一行
                    for cont in res :     #循环读取文本里面的内容                        
                        for column in range(0,len(cont)-1):
                            sheet.write(row,column,cont[column])
                        temp_dict = cont[3]
                        for column in range(3,len(head)):
                            if(head[column]=='智能化'):
                                val = temp_dict[head[column]]/2
                                sheet.write(row,column,val)
                            else:
                                sheet.write(row,column,temp_dict[head[column]])
                            column += 1
                        row += 1
                    xls.save(out_path)        #保存为xls文件
                except:
                    print("写xls文件出现异常了。。。。")
                
                finally:
                    pass
            elif line.startswith("开始分析第"):
                temp = line.strip("\n").replace("开始分析第","")
                year = temp.replace("年的数据","")


def get_key_words():
    # 下面这个关键字是王老师+岑岑师姐做的
    # key_words_text="国内生产总值、反腐败、八项规定、高质量发展、供给侧、和谐社会、美丽中国、国际经济、非公有制经济、国内外经济形势、宏观调控、国民经济、货币政策、通货膨胀、国际经济环境、结构性改革、中国制造“2025”、一带一路、京津冀协同发展、海上丝绸之路、产能过剩、亚太自贸区、长江经济带、混合所有制经济、经济合作区、利率市场化、RCEP、碳中和、国企改革三年行动、产业升级、智能制造、数字经济"

    # 下面这个关键字是岑岑师姐自己制定的
    key_words_text = "数字化、数字经济、产业数字化、数字产业化、智能制造、产业升级、供应链升级、智慧物流、智能化、互联网、大数据、云计算、人工智能、电子商务、云平台、云服务、中国制造2025、移动互联网、数据分析、数据挖掘、信息化、网络销售"
    key_words = []
    key_words = key_words_text.split("、")    
    return key_words
 

'''
将统计的字数结果写入文件
'''
def txt2_xls(file_path,xls_name):    
    try:
        for year in range(2010,2021):
            xls = xlwt.Workbook()
            file_name = file_path + "/" + str(year) + "_part.txt"
            xls_name = file_path + "/" + str(year) + "_part.xls"
            f = open(file_name)
            #生成excel的方法，声明excel
            sheet_name = str(year)
            sheet = xls.add_sheet(sheetname=sheet_name,cell_overwrite_ok=True)
            x = 0   #在excel开始写的位置（y）
            line = f.readline() # 过滤第一行
            line = f.readline()
            line = line.strip("\n")
            while(line): #循环读取文本里面的内容                   
                left,right = line.split('\t')
                right = int(right)
                if right > 20000 or right < 500:
                    line = f.readline()     #一行一行的读
                    line = line.strip("\n")
                    continue
                sheet.write(x,0,left)      #x单元格经度，i单元格纬度
                sheet.write(x,1,right)      #x单元格经度，i单元格纬度
                
                # for i in range(len(line.split('\t'))):   #\t即tab健分隔                    
                #     item = line.split('\t')[i]
                #     sheet.write(x,i,item)      #x单元格经度，i单元格纬度
                x += 1  #另起一行
                line = f.readline()     #一行一行的读
                line = line.strip("\n")
            f.close()
            xls.save(xls_name)        #保存为xls文件
    except:
        raise
    
    finally:
        pass


"""
删除无用的文件
"""
def delete_invalid_file(file_path):
    with open(file_path,'r') as f:
        line = f.readline()
        
        while(line):
            line = line.strip()
            if line!="" and os.path.exists(line):
                os.remove(line)
            line = f.readline()    


if __name__ == "__main__":        
    file_path = '/home/lawson/program/CompanyAnnualReport/analysis_2017.log'
    xls_name = '/home/lawson/program/CompanyAnnualReport/result'
    anslysis_log2_xls(file_path)
    # file_path = './a.txt'
    # delete_invalid_file(file_path)
    #txt2_xls(file_path,xls_name)
