import copy
from pdfminer.high_level import extract_text
import json
import xlwt
import os 
import sys

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
        cur_key_map_acc = {'一带一路': 0, '供给侧': 0, '八项规定': 0, '利率市场化': 0, '反腐败': 0, '和谐社会': 0, '国内生产总值': 0, '国际经济': 0,'非公有制经济':0,'国内外经济形势':0,'宏观调控':0,'国民经济':0,'货币政策':0,'通货膨胀':0,'国际经济环境':0,'结构性改革':0,
        '中国制造“2025”':0,'京津冀协同发展':0,'海上丝绸之路':0,'产能过剩':0,'亚太自贸区':0,'长江经济带':0,'混合所有制经济':0,
        '利率市场化':0,'高质量发展':0,'美丽中国':0,'经济合作区':0 } 
        while(line):
            line = f.readline()
            line = line.strip("\n")
            # print(line)
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
                out_path = file_path +f"_{year}.xls"
                xls = xlwt.Workbook()
                try:
                    #生成excel的方法，声明excel                    
                    sheet = xls.add_sheet(sheetname='sheet1',cell_overwrite_ok=True)
                    row = 0   #在excel开始写的位置（y）
                    column = 0
                    
                    head = ["公司名称",'是否出现','出现次数累计','一带一路', '供给侧', '八项规定', '利率市场化', '反腐败', '和谐社会', '国内生产总值', '国际经济','非公有制经济','国内外经济形势','宏观调控','国民经济','货币政策','通货膨胀','国际经济环境','结构性改革','中国制造“2025”','京津冀协同发展','海上丝绸之路','产能过剩','亚太自贸区','长江经济带','混合所有制经济','利率市场化','高质量发展','美丽中国','经济合作区']
                    for column in range(len(head)):
                        sheet.write(row,column,head[column])

                    row+=1 #另起一行
                    for cont in res :     #循环读取文本里面的内容                        
                        for column in range(0,len(cont)-1):
                            sheet.write(row,column,cont[column])
                        temp_dict = cont[3]                        
                        for column in range(3,len(head)):
                            sheet.write(row,column,temp_dict[head[column]])
                            column += 1
                        row += 1
                    xls.save(out_path)        #保存为xls文件
                except:
                    print("出现异常了。。。。")
                
                finally:
                    pass
            elif line.startswith("开始分析第"):
                temp = line.strip("\n").replace("开始分析第","")
                year = temp.replace("年的数据","")


def get_key_words():
    key_words_text="国内生产总值、反腐败、八项规定、高质量发展、供给侧、和谐社会、美丽中国、国际经济、非公有制经济、国内外经济形势、宏观调控、国民经济、货币政策、通货膨胀、国际经济环境、结构性改革、中国制造“2025”、一带一路、京津冀协同发展、海上丝绸之路、产能过剩、亚太自贸区、长江经济带、混合所有制经济、经济合作区、利率市场化"
    key_words = []
    key_words = key_words_text.split("、")    
    return key_words


#file_affilication = open('Affiliations.txt','r')
 
 
def txt2_xls(file_path,xlsname):
    xls = xlwt.Workbook()
    try:
        for year in range(2010,2020):
            file_name = file_path + "/" + str(year) + "_accumulate.txt"
            f = open(file_name)            
            #生成excel的方法，声明excel
            sheet_name = str(year)
            sheet = xls.add_sheet(sheetname=sheet_name,cell_overwrite_ok=True)
            x = 0   #在excel开始写的位置（y）
            line = f.readline()
            while(line):     #循环读取文本里面的内容            
                for i in range(len(line.split('\t'))):   #\t即tab健分隔
                    item = line.split('\t')[i]
                    sheet.write(x,i,item)      #x单元格经度，i单元格纬度
                x += 1  #另起一行
                line = f.readline()     #一行一行的读
        f.close()
        xls.save(xlsname)        #保存为xls文件
    except:
        raise
    
    finally:
        pass



def test_123(pdf_file_path):
    try:
        text = extract_text(pdf_file_path)
    except:
        text = []
        print(f"解析{pdf_file_path}文件出现异常")
        return
    print(text)
    cnt = text.count("产能过剩")
    print(cnt)
    

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
    file_path = './analysis_20_part_1.log'
    anslysis_log2_xls(file_path)
    # file_path = './a.txt'
    # delete_invalid_file(file_path)