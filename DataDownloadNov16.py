#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 15 13:10:07 2019

"""
import pandas as pd
import requests


linklist_xlsx = "/Users/lakshmimenont/Desktop/Working Python/data/DataFrancesco/ctorder_complete.xlsx"
dowloadPath = "/Users/lakshmimenont/Desktop/Working Python/data/DataFrancesco/PDFToRead/"


def getLinkListdf(linklistexcelPath):
    data = pd.read_excel(linklistexcelPath)
    print("read" + linklistexcelPath)
    df = pd.DataFrame(data, columns= ['id','FileName'])
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
    
def main():
    dflinklist = getLinkListdf(linklist_xlsx)
    
#    noOfRows = len(dflinklist)
    noOfRows = 10
    print(noOfRows)
    for x in range(noOfRows):
        link = dflinklist['FileName'][x]
        linkId = dflinklist['id'][x]
        print (link)
        print (linkId)
        
        resp = downloadLink(link)
        if(resp):
            
            filePath = dowloadPath + str(linkId) + ".pdf"
            open(filePath, 'wb').write(resp.content)
        
main()

