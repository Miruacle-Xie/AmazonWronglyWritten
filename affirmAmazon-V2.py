# -*- coding: utf-8 -*-
"""
Created on Thu Sep 15 17:23:38 2022

@author: Administrator
"""
import pandas as pd
import requests
import re
import os
import time
import sys


def getSubjectName(fileName):
    df = pd.read_excel(fileName, sheet_name=0)
    print(df)
    return df


def getAmazonResult(subjectName):
    try:
        url = "https://www.amazon.com/s?k=" + subjectName.replace(" ", "+") + "+shirt"
        # print(url)
        # '''
        web_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.71 Safari/537.36 Core/1.94.172.400 QQBrowser/11.1.5140.400',
            'authority': 'dr3fr5q4g2ul9.cloudfront.net'
            }
        # '''
        # web_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36'}
        r = requests.get(url, headers=web_header)
# =============================================================================
#         fd = open("1.txt", mode='w', encoding="utf8")
#         fd.write(r.text)
#         fd.close()
# =============================================================================
        # print(r.status_code)
        if r.status_code != 200:
            raise requests.ConnectionError("Expected status code 200, but got {}".format(r.status_code))

        # if r.status_code != 200:
        #     raise requests.ConnectionError("Expected status code 200, but got {}".format(r.status_code))
        r.encoding = r.apparent_encoding
        result = re.search(r'<span class="a-size-medium a-text-italic">(.*?) shirt</span>', r.text)
        # result1 = re.search(r"<span>.*(.*) results for</span>", r.text)
        result1 = re.search(r"<span>(.*) results for</span>|<span>.* of (.*) results for</span>", r.text)
        # print(result1)
        if result is not None:
            # print("已进行错别字修改")
            # print(result)
# =============================================================================
#             返回值1: 访问网站成功
#             返回值2: 匹配到亚马逊进行错别字修改
#             返回值3: 亚马逊修改后的内容
# =============================================================================
            return True, True, result[1], result1[1]
        else:
            # print(result)
            return True, False, None, result1[1]
        # print(r.text[:1000])
    except:
        print(r.status_code)
        # print(r.text[:1000])
        print("访问亚马逊异常")
        return False, None, None, None
    
    
def contrastSubjectName(crawlerFlag, crawlerResult, showingResult, searchResult, subjectName, tmplist1, tmplist2, tmplist3, tmplist4):
    tmp = []
    if crawlerFlag:
        tmplist3.append(searchResult)
        if crawlerResult:
            tmplist1.append(showingResult.capitalize())
            if showingResult.lower() == subjectName.lower():
                tmplist2.append("True-mode1")
                tmplist4.append("无")
                return
            else:
                tmplist2.append("False")
                tmp1 = showingResult.lower().split()
                tmp2 = subjectName.lower().split()
                for j, k in zip(tmp1, tmp2):
                    if j != k:
                        tmp.append(j)
                tmplist4.append(" | ".join(tmp))
                return
        else:
            tmplist1.append(subjectName.capitalize())
            tmplist2.append("True-mode2")
            tmplist4.append("无")
    else:
        tmplist1.append("异常")
        tmplist2.append("异常")
        tmplist3.append("异常")
        tmplist4.append("异常")
    return
    

def main():
    filepath = input("\n文件路径：\n")
    filepath = filepath.replace("\"", "").replace("\'", "")
    df = getSubjectName(filepath)
    # print(df)
    tmplist1 = []
    tmplist2 = []
    tmplist3 = []
    tmplist4 = []
    time_start = time.time()
    for i in range(len(df)):
    # for i in range(1):
        print("第{}次验证".format((i+1)))
        crawlerFlag, crawlerResult, showingResult, searchResult = getAmazonResult(df[df.columns.values[1]][i])
        contrastSubjectName(crawlerFlag, crawlerResult, showingResult, searchResult, df[df.columns.values[0]][i], tmplist1, tmplist2, tmplist3, tmplist4)
    series_tmplist1 = pd.Series(tmplist1)
    series_tmplist2 = pd.Series(tmplist2)
    series_tmplist3 = pd.Series(tmplist3)
    series_tmplist4 = pd.Series(tmplist4)
    df.insert(loc=len(df.columns), column="亚马逊匹配结果", value=series_tmplist1)
    df.insert(loc=len(df.columns), column="是否匹配", value=series_tmplist2)
    df.insert(loc=len(df.columns), column="搜索结果数量", value=series_tmplist3)
    df.insert(loc=len(df.columns), column="不匹配单词", value=series_tmplist4)
    reportpath = os.path.splitext(filepath)[0] + "-亚马逊检测报告" + ".xlsx"
    writer = pd.ExcelWriter(reportpath)
    # header = None：数据不含列名，index=False：不显示行索引（名字）
    df.to_excel(writer, header=None, index=False)
    writer.save()
    time_end = time.time()
    input("已生成报告, 耗时时间:{}, 平均耗时:{}, 按回车键结束".format(time_end - time_start, (time_end - time_start)/len(df)))
    
    
if __name__ == "__main__":
    # df = getSubjectName("0905-快主题-拆分结果.xlsx")
    try:
        os.environ['REQUESTS_CA_BUNDLE'] =  os.path.join(os.path.dirname(sys.argv[0]), 'cacert.pem')
        main()
    except Exception as e:
        repr(e)
    # df["显示结果"] = ""
    # df["匹配结果"] = ""
    # print(df)
    # for i in len(df)
    # getAmazonResult("calmm minnd stronng heaart violent hannds")
# =============================================================================
#     fd = open("1.txt", mode='r', encoding="utf8")
#     text = fd.read()
#     fd.close()
#     result = re.search(r'<span class="a-size-medium a-text-italic">(.*?)</span>', text)
#     print(result)
#     print(result[0])
#     print(result[1])
# =============================================================================
