"""
功能：对pdf的文件进行一个数据分析
"""
import copy
import logging
import os
from utils import get_key_words
from pdfminer.high_level import extract_text
import argparse


parser = argparse.ArgumentParser()
parser.add_argument("--year",type=int)
# pdf_file_path = "./勉刘申.pdf"
# text = extract_text(pdf_file=pdf_file_path)
# print(text)


key_words = get_key_words()
part_key_words = ['董事会报告','董事会报','董事局报告','董事会工作报告','管理层讨论与分析','经营情况讨论与分析'] 


def judge(i):
    if i.startswith("第一节"):
        return True
    if i.startswith("第二节"):
        return True
    if i.startswith("第三节"):
        return True
    if i.startswith("第四节"):
        return True
    
    if i.startswith("第五节"):
        return True
    if i.startswith("第六节"):
        return True
    
    if i.startswith("第七节"):
        return True
    if i.startswith("第八节"):
        return True
    
    if i.startswith("第九节"):
        return True
    if i.startswith("第十节"):
        return True
    
    
    if i.startswith("第一章"):
        return True
    if i.startswith("第二章"):
        return True
    if i.startswith("第三章"):
        return True
    if i.startswith("第四章"):
        return True
    
    if i.startswith("第五章"):
        return True
    if i.startswith("第六章"):
        return True
    
    if i.startswith("第七章"):
        return True
    if i.startswith("第八章"):
        return True
    
    if i.startswith("第九章"):
        return True
    if i.startswith("第十章"):
        return True


    if i.startswith("一、"):
        return True
    if i.startswith("二、"):
        return True
    if i.startswith("三、"):
        return True
    if i.startswith("四、"):
        return True
    
    if i.startswith("五、"):
        return True
    if i.startswith("六、"):
        return True
    
    if i.startswith("七、"):
        return True
    if i.startswith("八、"):
        return True
    
    if i.startswith("九、"):
        return True
    if i.startswith("十、"):
        return True
    if "...." in i or "···" in i or "-------" in i or "……" in i or "… … " in i or "„„„„„„" in i:
        return True
    return False

"""
key_word_map_acumulate 是累积map
key_word_map 是非累积map
"""
def analysis(pdf_file_path,key_word_map, key_word_map_accumulate):
    try:
        text = extract_text(pdf_file_path)
        text = text.replace("\n","") # 将所有的换行转成空白部分
    except:
        text = []
        print(f"解析{pdf_file_path}文件出现异常")
        return
    for word in key_words:
        cnt = text.count(word)
        if cnt >= 1:
            if word not in key_word_map.keys():
                key_word_map[word] = 1
            else :
                key_word_map[word] += 1
        
        if word not in key_word_map_accumulate.keys():
            key_word_map_accumulate[word] = cnt

        else:
            key_word_map_accumulate[word] += cnt




# text 表示的是原pdf解析的结果
# 如果在出现 -1, -2 ,-3 这种情况时，就调用这个函数
def reverse_analysis(text,cur_key_word,next_key_word):
    text = text[::-1] # 翻转
    cur_key_word = cur_key_word[::-1]
    next_key_word = next_key_word[::-1]

    index_1 = text.index(cur_key_word)
    index_2 = text.index(next_key_word)
    while index_1 >= index_2 and cur_key_word in part_key_words and next_key_word in part_key_words:
        index_2 = text.index(next_key_word,index_2+1)
    
    aim_part = text[index_1:index_2] # 找出目标内容
    aim_part = aim_part.replace("\n","")
    aim_part = aim_part.replace(".","") 
    aim_part = aim_part.replace(",","") 
    for i in range(10):
        num = str(i)
        aim_part = aim_part.replace(num,"") # 将所有的数字替换掉
    aim_part = aim_part.replace(" ","") 
    aim_part = aim_part.replace("\xa0","") 
    aim_part = aim_part.replace("\x0c","") 
    
    punctuation = [ "。","？", "！", "，","、", "；", "：","‘",
            "’", "“", "”", "（","）", "〔", "〕", "【", "】", "—", "…","–", "―", '《', '》', '．','%','','．', # chinese
            ',','.',':', '-','(',')',
            '√','□','[',']','+','/',
            'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'
            ]
    for i in punctuation:
        aim_part = aim_part.replace(i,"") # 将所有的符号替换掉
    return len(aim_part)


"""
-1 表示解析pdf时出现异常
-2 表示解析目录时出现异常
-3 表示关键字提取失败
-4 表示其它问题
"""
# 分析部分数据，主要是："董事会报告"，“管理层讨论与分析”，“经营情况讨论与分析” 三个中的一个
def part_analysis(pdf_file_path):
    try:
        text = extract_text(pdf_file_path)
        raw_text = copy.copy(text)
    except:
        text = []
        print(f"解析{pdf_file_path}文件出现异常")
        return -1
    
    try:
        # 提取出目录内容
        #index_content = text.index("目录") # 找到目录的索引 => 因为有时候不一定能找到，但是目录一般都放在开头，所以这里我直接取0做下标
        index_content = 0
        temp = text[index_content:index_content+2500]        
        temp.replace("\xa0","")
        temp.replace(" ","") #替换空格    
        temp = temp.split("\n")    
        catalog = [] #存储下目录的内容
        
        # 用于判断当前这个财报是什么关键字，并找出其接下来的一个关键字
        cur_key_word = "" 
        next_key_word = ""
        cur_key_word_2 = ""  # 用_2存储第二部分
        next_key_word_2 = ""
        for i in temp:
            if judge(i):
                i = i.replace("\xa0","")  
                i= i.strip("第节页章1234567890()（）．.„·‐—… ") # 去掉后面的杂项                     
                if len(i) == 0:
                    continue
                catalog.append(i)
                if len(i) == 0:
                    continue
                a = i.split()
                if len(a) == 1:
                    a = i.split("、") # 再以、号分割
                cur = a[-1]
                #cur = a[0] # 以“X节/章” 作为索引
                
                # TODO:会不会出现part_key_words 中的关键字都出现在目录中了？            
                if cur_key_word!="" and next_key_word =="" and a[-1] != cur_key_word:
                    # 这里是为了避免next_key_word_2 和 cur_key_word 相同的情况
                        if len(a[0]) > 1: # 如果是七节、八节这种，则当索引，否则还是使用“董事会报告”，“监事会报告”做索引
                            next_key_word = a[0]
                            next_key_word = next_key_word.replace("届","节") # 有的公司会把节错写成届
                            if next_key_word.count("节") > 1:# 防止出现“第七节第七节第七节第七节” 这种情况
                                next_key_word = a[-1]
                            next_key_word_2 = a[-1]
                        else:
                            next_key_word = a[-1]
                            #next_key_word = next_key_word.strip("第一二三四五六七八九十节页 章")
                if cur in part_key_words : 
                    if len(a[0]) > 1 :
                        cur_key_word = a[0]
                        cur_key_word = cur_key_word.replace("届","节")
                        if cur_key_word.count("节") > 1: # 防止出现
                            cur_key_word = a[-1]
                        cur_key_word_2 = a[-1]
                    else:                        
                        cur_key_word = a[-1]
                #print(i)
        if len(catalog) < 5: # 如果解析的结果数目不够，那么可能是存在问题的
            index_content = 0
            temp = text[index_content:index_content+4000]            
            temp.replace("\xa0","")
            temp.replace(" ","") #替换空格    
            temp = temp.split("\n")    
            catalog = [] #存储下目录的内容
            
            # 用于判断当前这个财报是什么关键字，并找出其接下来的一个关键字
            cur_key_word = "" 
            next_key_word = ""
            for i in temp:
                if judge(i):
                    i = i.replace("\xa0","")
                    i= i.strip("第节页章1234567890()（）.„·-… ") # 去掉后面的杂项                     
                    if len(i) == 0:
                        continue
                    catalog.append(i)
                    if len(i) == 0:
                        continue
                    a = i.split()
                    if len(a) == 1:
                        a = i.split("、") # 再以、号分割
                    cur = a[-1]
                    #cur = a[0] # 以“X节/章” 作为索引
                    
                    # TODO:会不会出现part_key_words 中的关键字都出现在目录中了？            
                    if cur_key_word!="" and next_key_word =="":
                        if len(a[0]) > 1: # 如果是七节、八节这种，则当索引，否则还是使用“董事会报告”，“监事会报告”做索引
                            next_key_word = a[0]
                            next_key_word = next_key_word.replace("届","节") # 有的公司会把节错写成届
                        else:
                            next_key_word = a[-1]
                            #next_key_word = next_key_word.strip("第一二三四五六七八九十节页 章")
                    if cur in part_key_words : 
                        if len(a[0]) > 1 :
                            cur_key_word = a[0]
                            cur_key_word = cur_key_word.replace("届","节")
                        else:                        
                            cur_key_word = a[-1]
            text = text[4000::] #取剩下的部分

            print(catalog)
            if len(catalog) < 5:  # 就是没有目录的pdf              
                if cur_key_word_2 == "" or next_key_word_2 =="":
                    return -3
                else:
                    return reverse_analysis(raw_text,cur_key_word_2,next_key_word_2)
        else:
            print(catalog)
            text = text[2500::] #取剩下的部分
        text = text.replace("\n","")
        if cur_key_word == '' or next_key_word=='':
            if cur_key_word_2 == "" or next_key_word_2 =="":                
                return -3
            else:
                return reverse_analysis(raw_text,cur_key_word_2,next_key_word_2)
        index_1 = text.index(cur_key_word)
        index_2 = text.index(next_key_word)
        while index_1 + 100 >= index_2 : # 必须超过一定的阈值范围
            index_2 = text.index(next_key_word,index_2+1)
        
        aim_part = text[index_1:index_2] # 找出目标内容
        aim_part = aim_part.replace("\n","")
        aim_part = aim_part.replace(".","") 
        aim_part = aim_part.replace(",","") 
        for i in range(10):
            num = str(i)
            aim_part = aim_part.replace(num,"") # 将所有的数字替换掉
        aim_part = aim_part.replace(" ","") 
        aim_part = aim_part.replace("\xa0","") 
        aim_part = aim_part.replace("\x0c","") 
        
        punctuation = [ "。","？", "！", "，","、", "；", "：","‘",
                "’", "“", "”", "（","）", "〔", "〕", "【", "】", "—", "…","–", "―", '《', '》', '．','%','','．', # chinese
                ',','.',':', '-','(',')',
                '√','□','[',']','+','/',
                'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'
                ]
        for i in punctuation:
            aim_part = aim_part.replace(i,"") # 将所有的符号替换掉
    except:
        print("提取内容遇到问题...")
        return -4
    return len(aim_part)


# log_filename = "analysis.log" # 得到每个年份的日志
# # step2. 构建日志记录器
# logging.basicConfig(level=logging.INFO,
#                 filemode='w',
#                 filename="./log/"+log_filename,
#                 )
# logg = logging.getLogger("analysis")

args = parser.parse_args()

if __name__ == "__main__":
    # part_1 => 详细分析每个报表中所含关键字
    # file_path = "/home/lawson/program/CompanyAnnualReport/report/"
    # res_path = "/home/lawson/program/CompanyAnnualReport/result/"
    # year = args.year

    # print(f"开始分析第{year}年的数据")
    # key_word_map = {}
    # key_word_map_accumulate = {}
    # cur_path = file_path + str(year)
    # cur_res_file = res_path + str(year) +".txt"
    # cur_res_file_accumulate = res_path + str(year) +"_accumulate.txt"

    # #print(cur_path)
    # lists = os.listdir(cur_path)
    # for name in lists:
    #     cur_file_name = cur_path +"/"+ name
    #     print(cur_file_name)
    #     part_analysis(cur_file_name,key_word_map,key_word_map_accumulate)
    #     print(f"key_word_map={key_word_map}")
    #     print(f"key_word_map_accumulate={key_word_map_accumulate}")
    # print(f"第{year}年的数据分析已结束")

    # print(f"将结果写入到文件中...")
    # # 写入非累积结果
    # with open(cur_res_file,'a') as f:
    #     f.write(f"下面是第{year}年的结果")
    #     for item in key_word_map.items():
    #         key,value = item
    #         f.write(key+"\t"+str(value)+"\n")
    #     f.write("\n")

    # # 写入累积结果
    # with open(cur_res_file_accumulate,'a') as f:
    #     f.write(f"下面是第{year}年的结果")
    #     for item in key_word_map_accumulate.items():
    #         key,value = item
    #         f.write(key+"\t"+str(value)+"\n")
    #     f.write("\n")


    # part_2
    file_path = "/home/lawson/program/CompanyAnnualReport/report/"
    res_path = "/home/lawson/program/CompanyAnnualReport/result/"
    year = args.year

    print(f"开始分析第{year}年的数据")
    word_map = {} # name => num
    cur_path = file_path + str(year)
    cur_res_file = res_path + str(year) +"_part.txt"
    

    #print(cur_path)
    lists = os.listdir(cur_path)
    for name in lists:
        cur_file_name = cur_path +"/"+ name
        print(cur_file_name)
        word_num = part_analysis(cur_file_name)
        word_map[name] = word_num

    print(f"第{year}年的数据分析已结束")

    print(f"将结果写入到文件中...")
    # 写入非累积结果
    with open(cur_res_file,'w') as f:
        f.write(f"下面是第{year}年的结果")
        for item in word_map.items():
            key,value = item
            f.write(key+"\t"+str(value)+"\n")
        f.write("\n")