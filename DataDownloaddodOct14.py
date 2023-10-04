#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 15 15:56:38 2019

"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import xlsxwriter


linklist_xlsx = "/Users/lakshmimenon/Desktop/Working Python/data/Cartel1.xlsx"
WRITE_XLS =  "/Users/lakshmimenon/Desktop/Working Python/data/DoD14.xlsx"

def getLinkListdf(linklistexcelPath):
    data = pd.read_excel(linklistexcelPath)
    df = pd.DataFrame(data, columns= ['link'])
    return df


def downloadLink(link):
    url = link
    print(url)
    response = requests.get(url)
    print(response.status_code)

    if response.status_code == 200:
        return response
    else:
        print("Errors making the request",response.status_code)
        return None

def getCompDets(awardPara,datePresent,linkDate):
    while(awardPara.startswith('Â') or awardPara.startswith('"') or awardPara.startswith(' ')):
        awardPara = awardPara[2:]
    #print(awardPara)
    
    if(datePresent == 1):
        companyEndIndex = awardPara.find('was awarded on') - 2
        companyName = awardPara[0:companyEndIndex]
        dateStartIndex = companyEndIndex + 16
        dateEndIndex = dateStartIndex + 14
        date = awardPara[dateStartIndex:dateEndIndex]
    else:
        if(awardPara.find('is being awarded')!= -1):
            companyEndIndex = awardPara.find('is being awarded') - 2
        if(awardPara.find('was awarded')!= -1):
            companyEndIndex = awardPara.find('was awarded') - 2
        if(awardPara.find('has been awarded')!= -1):
            companyEndIndex = awardPara.find('has been awarded') - 2
        companyName = awardPara[0:companyEndIndex]
        date = linkDate
    
    amountStartIndex = awardPara.find('$')
    if(amountStartIndex != -1):
        amountStartIndex = amountStartIndex + 1
        amountStartString = awardPara[amountStartIndex:]
        amountEndIndex = amountStartString.index(' ')
        amount = amountStartString[:amountEndIndex]
    else:
        amount = 'NA'
        
    
    
    #print(companyName)
    #print(date)
    #print(amount)
    
    compDet = [companyName.strip(),date.strip(),amount.strip()]
    return compDet

def getMonthIndex(linkDateText):
    if(linkDateText.find('January')!=-1):
        return linkDateText.find('January')
    elif(linkDateText.find('February')!=-1):
        return linkDateText.find('February')
    elif(linkDateText.find('March')!=-1):
        return linkDateText.find('March')
    elif(linkDateText.find('April')!=-1):
        return linkDateText.find('April')
    elif(linkDateText.find('May')!=-1):
        return linkDateText.find('May')
    elif(linkDateText.find('June')!=-1):
        return linkDateText.find('June')
    elif(linkDateText.find('July')!=-1):
        return linkDateText.find('July')
    elif(linkDateText.find('August')!=-1):
        return linkDateText.find('August')
    elif(linkDateText.find('September')!=-1):
        return linkDateText.find('September')
    elif(linkDateText.find('October')!=-1):
        return linkDateText.find('October')
    elif(linkDateText.find('November')!=-1):
        return linkDateText.find('November')
    elif(linkDateText.find('December')!=-1):
        return linkDateText.find('December')
    else:
        return -1
    
    
def getCompDetsList(resp):
    compDetList = []
    soup = BeautifulSoup(resp.text, 'html.parser')
            #print(soup)
    count = 0
    
    for body in soup.find_all('div',class_='PressOpsContentBody'):
        print('body number: ',count)
        print(body.text)
        paraPresent = 0
        if(count == 0):
            linkDateText= body.text
            print(linkDateText)
            MonthIndex = getMonthIndex(linkDateText)
            if(MonthIndex != -1):
                linkDate = linkDateText[MonthIndex:]
        for para in body.find_all('p',style=''):
            if(paraPresent == 0):
                paraPresent = 1
            award = para.text
            datePresent = -1
                
            if(award.find('was awarded')!=-1):
                datePresent = 0
                if(award.find('was awarded on')!=-1):
                    datePresent = 1
                
                        
            if(award.find('is being awarded')!=-1):
                datePresent = 0
                
            if(award.find('has been awarded')!=-1):
                datePresent = 0
                    
            if(datePresent!= -1):
                compDet = getCompDets(award,datePresent,linkDate)
                print(compDet)
                compDetList.append(compDet)
                        
                    
                    #award.lstrip()
                    #firstSentenceEnd = award.find('Â')
                    #print(firstSentenceEnd)
        if(count!=0 and paraPresent==0):
            for div in body.find_all('div',style=''):
                award = div.text
                datePresent = -1
                
                if(award.find('was awarded')!=-1):
                    datePresent = 0
                    if(award.find('was awarded on')!=-1):
                        datePresent = 1
                
                        
                if(award.find('is being awarded')!=-1):
                    datePresent = 0
                
                if(award.find('has been awarded')!=-1):
                    datePresent = 0
                
                if(datePresent!= -1):
                    compDet = getCompDets(award,datePresent,linkDate)
                    print(compDet)
                    compDetList.append(compDet)
                
        count += 1
    return compDetList
    


def writeExcel(totalCompDetList):
    workbook = xlsxwriter.Workbook(WRITE_XLS)
    worksheet = workbook.add_worksheet("Company Details sheet") 
    row = 0
    col = 0
    for compDetList in totalCompDetList:
        for compDet in compDetList:
            print("companyName")
            print(compDet[0])
            print("companyDate")
            print(compDet[1])
            print("companyAmt")
            print(compDet[2])
        
            worksheet.write(row, col, compDet[0]) 
            worksheet.write(row, col + 1, compDet[1]) 
            worksheet.write(row, col + 2, compDet[2])
            row += 1
    workbook.close()
        
def main():
    dflinklist = getLinkListdf(linklist_xlsx)
    totalCompDetList = []
    noOfRows = len(dflinklist)
#    noOfRows = 51
    print(noOfRows)
    for x in range(noOfRows):
        link = dflinklist['link'][x]
        print (link)
        resp = downloadLink(link)
        if(resp):
            linkCompDetList = getCompDetsList(resp)
            #print(linkCompDetList)
            totalCompDetList.append(linkCompDetList)
    
    print("checkpoint")
    
    writeExcel(totalCompDetList)
    

main()
                