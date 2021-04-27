import xlwt
import os 
import sys

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


if __name__ == "__main__":
    # a = get_key_words()
    # print(a)
    file_path = './result_1300'
    xlsname = './res_1300_accumulate.xls'
    txt2_xls(file_path,xlsname)