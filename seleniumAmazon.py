# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 14:23:56 2022

@author: Administrator
"""
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
import os
import time
import random


def amazonDeliverInit(driver):
    url = "https://www.amazon.com"
    try:
        driver.get(url)
        content = driver.page_source
        if "New York 10041" not in content:
            driver.find_element_by_xpath('//*[@id="nav-packard-glow-loc-icon"]').click()
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="GLUXZipUpdateInput"]')))  # 元素是否可见
            driver.find_element_by_xpath('//*[@id="GLUXZipUpdateInput"]').send_keys("10041")
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="GLUXZipUpdate"]/span/input').click()
            time.sleep(3)
            driver.refresh()
            content1 = driver.page_source
            fd = open("1.txt", mode='w', encoding="utf8")
            fd.write(content1)
            fd.close()
            if "New York 10041" not in content1:
                print("设置地址失败")
                return False
            else:
                print("地址设为10041成功")
                return True
    except:
        print("设置地址异常")
        return False


def getSubjectName(fileName):
    df = pd.read_excel(fileName, sheet_name=0)
    # print(df)
    return df


def getAmazonResult(driver, subjectName):
    try:
        # url = "https://www.amazon.com/s?k=" + subjectName.replace(" ", "+") + "+shirt"
        # print(url)
        driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').click()
        eleValue = driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]')
        if eleValue.get_attribute('value') is not None:
            driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').clear()
        driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').send_keys(subjectName + " shirt")
        driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').send_keys(Keys.ENTER)
        pageSource = driver.page_source
# =============================================================================
        fd = open("2.txt", mode='w', encoding="utf8")
        fd.write(pageSource)
        fd.close()
# =============================================================================
        return True, pageSource
    except Exception as e:
        repr(e)
        return False, None
    
    
def contrastSubjectName(correctWritten, subjectName, tmpFlag, matchReslut, uncorrectWord):
    tmp = []
    if correctWritten.lower() == subjectName.lower():
        if tmpFlag == 1:
            matchReslut.append("True-mode1")
        else:
            matchReslut.append("True-mode2")
        uncorrectWord.append("无")
        return
    else:
        matchReslut.append("False")
        tmp1 = correctWritten.lower().split()
        tmp2 = subjectName.lower().split()
        for j, k in zip(tmp1, tmp2):
            if j != k:
                tmp.append(j)
        uncorrectWord.append(" | ".join(tmp))
        return
    

def checkmodify(driver, df, filepath):
    amaShowReslut = []
    matchReslut = []
    productNumReslut = []
    uncorrectWord = []
    time_start = time.time()
    for i in range(len(df)):
    # for i in range(3):
        print("第{}次验证".format((i + 1)))
        accessFlag, pageSource = getAmazonResult(driver, df[df.columns.values[1]][i])
        if accessFlag:
            correctWritten = re.search(r'<span class="a-size-medium a-text-italic">(.*?) shirt</span>', pageSource)
            searchResult = re.search(r"<span>(.*) results for</span>|<span>.* of (.*) results for</span>", pageSource)
            productNumReslut.append(searchResult[1])
            if correctWritten is not None:
                amaShowReslut.append(correctWritten[1])
                tmpFlag = 1
                contrastSubjectName(correctWritten[1], df[df.columns.values[0]][i], tmpFlag, matchReslut, uncorrectWord)
            else:
                print(correctWritten)
                amaShowReslut.append(df[df.columns.values[1]][i].capitalize())
                tmpFlag = 2
                contrastSubjectName(df[df.columns.values[1]][i], df[df.columns.values[0]][i], tmpFlag, matchReslut, uncorrectWord)
        else:
            amaShowReslut.append("异常")
            matchReslut.append("异常")
            productNumReslut.append("异常")
            uncorrectWord.append("异常")
        # time.sleep(random.uniform(1, 5))
    series_amaShowReslut = pd.Series(amaShowReslut)
    series_matchReslut = pd.Series(matchReslut)
    series_productNumReslut = pd.Series(productNumReslut)
    series_uncorrectWord = pd.Series(uncorrectWord)
    df.insert(loc=len(df.columns), column="亚马逊显示结果", value=series_amaShowReslut)
    df.insert(loc=len(df.columns), column="是否匹配", value=series_matchReslut)
    df.insert(loc=len(df.columns), column="产品数量", value=series_productNumReslut)
    df.insert(loc=len(df.columns), column="不匹配单词", value=series_uncorrectWord)
    reportpath = os.path.splitext(filepath)[0] + "-亚马逊检测报告" + ".xlsx"
    writer = pd.ExcelWriter(reportpath)
    # header = None：数据不含列名，index=False：不显示行索引（名字）
    df.to_excel(writer, header=None, index=False)
    writer.save()
    time_end = time.time()
    driver.quit()
    input("已生成报告, 耗时时间:{}, 平均耗时:{}, 按回车键结束".format(time_end - time_start, (time_end - time_start) / len(df)))


def main():
    filepath = input("\n文件路径：\n")
    filepath = filepath.replace("\"", "").replace("\'", "")
    df = getSubjectName(filepath)
    df = df.iloc[:, [0, 1]]
    print(df)
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(chrome_options=chrome_options)
    # driver = webdriver.Chrome()
    if amazonDeliverInit(driver):
        print("即将开始...")
        checkmodify(driver, df, filepath)
    else:
        input("访问亚马逊异常, 回车结束")

    
    
if __name__ == "__main__":
    # df = getSubjectName("0905-快主题-拆分结果.xlsx")
    try:
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
