"""
功能：对pdf的文件进行一个数据分析
"""
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
    

# log_filename = "analysis.log" # 得到每个年份的日志
# # step2. 构建日志记录器
# logging.basicConfig(level=logging.INFO,
#                 filemode='w',
#                 filename="./log/"+log_filename,
#                 )
# logg = logging.getLogger("analysis")

args = parser.parse_args()


if __name__ == "__main__":
    file_path = "/home/lawson/program/CompanyAnnualReport/report/"
    res_path = "/home/lawson/program/CompanyAnnualReport/result/"
    year = args.year

    print(f"开始分析第{year}年的数据")
    key_word_map = {}
    key_word_map_accumulate = {}
    cur_path = file_path + str(year)
    cur_res_file = res_path + str(year) +".txt"
    cur_res_file_accumulate = res_path + str(year) +"_accumulate.txt"

    #print(cur_path)
    lists = os.listdir(cur_path)
    for name in lists:
        cur_file_name = cur_path +"/"+ name
        print(cur_file_name)
        analysis(cur_file_name,key_word_map,key_word_map_accumulate)
        print(f"key_word_map={key_word_map}")
        print(f"key_word_map_accumulate={key_word_map_accumulate}")
    print(f"第{year}年的数据分析已结束")

    print(f"将结果写入到文件中...")
    # 写入非累积结果
    with open(cur_res_file,'a') as f:
        f.write(f"下面是第{year}年的结果")
        for item in key_word_map.items():
            key,value = item
            f.write(key+"\t"+str(value)+"\n")
        f.write("\n")

    # 写入累积结果
    with open(cur_res_file_accumulate,'a') as f:
        f.write(f"下面是第{year}年的结果")
        for item in key_word_map_accumulate.items():
            key,value = item
            f.write(key+"\t"+str(value)+"\n")
        f.write("\n")