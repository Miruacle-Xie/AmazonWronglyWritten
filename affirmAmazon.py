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
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36',
            'Accept': '*/*',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Cookie': 'session-id=144-1148363-6624935; i18n-prefs=USD; ubid-main=130-0277329-2980128;\
            lc-main=en_US; AMCV_4A8581745834114C0A495E2B%40AdobeOrg=-1124106680%7CMCIDTS%7C18998%\
            7CMCMID%7C58168389969635584703677386317991191684%7CMCAAMLH-1641958883%7C11%7CMCAAMB-1\
            641958883%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1641361284s\
            %7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.2.0; mbox=session#508dd06d82d74223b2a50bc2cc492d42\
            #1641355945|PC#508dd06d82d74223b2a50bc2cc492d42.32_0#1704598885; _mkto_trk=id:365-EFI-026\
            &token:_mch-amazon.com-1641354086219-49525; s_lv=1641354098854; aws-target-data=%7B%22support\
            %22%3A%221%22%7D; aws-target-visitor-id=1641354102852-582186.32_0; aws-ubid-main=464-3147758-7407865; \
            AMCV_7742037254C95E840A4C98A6%40AdobeOrg=1585540135%7CMCIDTS%7C19052%7CMCMID%7C577434032950566551737165\
            08246953791791%7CMCAAMLH-1646651490%7C11%7CMCAAMB-1646651490%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3\
            xzPWQmdj0y%7CMCOPTOUT-1646053890s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C4.4.0; regStatus=pre-register; \
            session-id-time=2082787201l; s_vnum=2058851218014%26vn%3D2; s_nr=1651113556237-New; s_dslv=1651113556239; \
            skin=noskin; session-token=kZDxTwLctVuWUE1FkwVs/Rmw6a/iEuvgITRqqAaaTr5r02UcbrU5KHG+E+RZemoUdpkJestzezMjUPc\
            OVgC2gCCkvQik7ZihRuDjc4B2zarCc5Hjgp57LWXA/vMinR6GqhbqcwKlC5tKAwuNvITvE+5QCSIbfteRi57FDpnNbBkFRWbR+OpgVcmj2y\
            lRKnPkWgKzaF2lpT6TxW4YzUhoMlEZIxYvFCJs',
            'TE': 'Trailers'}
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
        
        result = re.search(
            r'<span class="a-size-medium a-text-italic">(.*?) shirt</span>', r.text)
        # print(result)
        if result is not None:
            # print("已进行错别字修改")
            # print(result)
# =============================================================================
#             返回值1: 访问网站成功
#             返回值2: 匹配到亚马逊进行错别字修改
#             返回值3: 亚马逊修改后的内容
# =============================================================================
            return True, True, result[1]
        else:
            print(result)
            return True, False, None
        # print(r.text[:1000])
    except:
        print(r.status_code)
        # print(r.text[:1000])
        print("访问亚马逊异常")
        return False, None, None
    
    
def contrastSubjectName(crawlerFlag, crawlerResult, showingResult, subjectName, tmplist1, tmplist2):
    if crawlerFlag:
        if crawlerResult:
            tmplist1.append(showingResult.capitalize())
            if showingResult.lower() == subjectName.lower():
                tmplist2.append("True")
                return
            else:
                tmplist2.append("False")
                return
        else:
            tmplist1.append(subjectName.capitalize())
            tmplist2.append("True")
    else:
        tmplist1.append("异常")
        tmplist2.append("异常")
    return
    

def main():
    filepath = input("\n文件路径：\n")
    filepath = filepath.replace("\"", "").replace("\'", "")
    df = getSubjectName(filepath)
    # print(df)
    tmplist1 = []
    tmplist2 = []
    time_start = time.time()
    for i in range(len(df)):
    # for i in range(1):
        print("第{}次验证".format((i+1)))
        crawlerFlag, crawlerResult, showingResult = getAmazonResult(df[df.columns.values[1]][i])
        contrastSubjectName(crawlerFlag, crawlerResult, showingResult, df[df.columns.values[0]][i], tmplist1, tmplist2)
    # print("tmplist1:{}".format(tmplist1))
    # print("tmplist2:{}".format(tmplist2))
    series_tmplist1 = pd.Series(tmplist1)
    series_tmplist2 = pd.Series(tmplist2)
    df.insert(loc=len(df.columns), column="显示结果", value=series_tmplist1)
    df.insert(loc=len(df.columns), column="匹配结果", value=series_tmplist2)
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
