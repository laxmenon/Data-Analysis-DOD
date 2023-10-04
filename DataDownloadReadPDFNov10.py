#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 28 13:50:11 2019

@author: lakshmimenon
"""

import PyPDF2
import xlsxwriter
import traceback

uploadFilesPath = "/Users/lakshmimenont/Desktop/Working Python/data/DataFrancesco/PDFToRead/"
dowloadPath = "/Users/lakshmimenont/Desktop/Working Python/data/DataFrancesco/DownloadedFiles/pdfDets.xlsx"


# pdf file object
def readPDF(numberOfPDFToRead):
    filePath = uploadFilesPath + str(numberOfPDFToRead) +".pdf"
    
    try:
        # pdf file object
        pdfFileObj = open(filePath, 'rb')
        # pdf reader object
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        # a page object
        pageObj = pdfReader.getPage(0)
        
        # extracting text from page.
        # this will print the text you can also save that into String
        #print(pageObj.extractText())
        pdfText = pageObj.extractText()
    except Exception:
        traceback.print_exc()
        print("Error reading pdf number",numberOfPDFToRead)
        pdfText = "-1"
    return pdfText
    
def getCompDate(pdfText):
    mainDate = ''
    if(pdfText.find('COMMISSION')!=-1):
        try:
            mainDateSearchIndex = pdfText.find('COMMISSION') + 10
            #print(mainDateSearchIndex)
            while(pdfText[mainDateSearchIndex]==' '):
                mainDateSearchIndex = mainDateSearchIndex + 1
            mainDateStartIndex = mainDateSearchIndex
            #print(mainDateStartIndex)
            
            mainDateSearchEndIndex = mainDateStartIndex
            while(pdfText[mainDateSearchEndIndex]!=','):
                mainDateSearchEndIndex = mainDateSearchEndIndex + 1
            mainDateSearchEndIndex = mainDateSearchEndIndex + 1
            #print(pdfText[mainDateSearchEndIndex])
            #print(mainDateSearchEndIndex)
            while(pdfText[mainDateSearchEndIndex] == ' '):
                mainDateSearchEndIndex = mainDateSearchEndIndex + 1
            #print(pdfText[mainDateSearchEndIndex])
            while(pdfText[mainDateSearchEndIndex]!= ' '):
                mainDateSearchEndIndex = mainDateSearchEndIndex + 1
            mainDateEndIndex = mainDateSearchEndIndex
            #print(mainDateEndIndex)
            
            mainDate = pdfText[mainDateStartIndex:mainDateEndIndex]
        except Exception:
            traceback.print_exc()
            mainDate = 'NA' 
    else:
        mainDate = 'NA'
    return mainDate

def getCompName(pdfText):
    compName = ''
    if(pdfText.find('1934')!=-1):
        try:
            compNameSearchIndex = pdfText.find('1934') + 4
            while(pdfText[compNameSearchIndex]==' '):
                compNameSearchIndex = compNameSearchIndex + 1
            compNameStartIndex = compNameSearchIndex
           
            compNameSearchEndIndex = compNameStartIndex
            while(pdfText[compNameSearchEndIndex]!='.'):
                compNameSearchEndIndex = compNameSearchEndIndex + 1
            compNameEndIndex = compNameSearchEndIndex + 1
            
            compName = pdfText[compNameStartIndex:compNameEndIndex]
        except Exception:
            traceback.print_exc()
            compName = 'NA'
    else:
        compName = 'NA'
    return compName

def getCompReq(pdfText):
    compReq = ''
    if(pdfText.find('24b-2')!=-1):
        try:
            compReqSearchIndex = pdfText.find('24b-2') + 5
            while(pdfText[compReqSearchIndex]==' '):
                compReqSearchIndex = compReqSearchIndex + 1
            compReqStartIndex = compReqSearchIndex 
            
            if(pdfText.find('for information') != -1):
                compReqEndIndex = pdfText.find('for information') - 1
                compReq = pdfText[compReqStartIndex:compReqEndIndex]
            else:
                compReq = 'NA'
        except Exception:
            traceback.print_exc()
            compReq = 'NA'
    else:
        print("24b-2 not found")
        compReq = 'NA'
    return compReq

def getCompFormFiled(pdfText):
    compFormFiled =''
    if(pdfText.find('Form')!=-1):
        try:
            compFormFiledSearchIndex = pdfText.find('Form')
            compFormFiledStartIndex = compFormFiledSearchIndex
            compFormFiledSearchEndIndex = compFormFiledStartIndex + 4
            while(pdfText[compFormFiledSearchEndIndex] ==' '):
                compFormFiledSearchEndIndex = compFormFiledSearchEndIndex + 1
            while(pdfText[compFormFiledSearchEndIndex] !=' '):
                compFormFiledSearchEndIndex = compFormFiledSearchEndIndex + 1
            compFormFiledEndIndex = compFormFiledSearchEndIndex
            
            compFormFiled = pdfText[compFormFiledStartIndex:compFormFiledEndIndex]
        except Exception:
            traceback.print_exc()
            compFormFiled = 'NA'
    else:
        compFormFiled = 'NA'
    return compFormFiled

def getCompFormFiledOnDate(pdfText):
    compFormFiledOnDate = ''
    if(pdfText.find('filed  on')!=-1 or pdfText.find('filed on')!= -1):
        try:
            if(pdfText.find('filed  on')!=-1):
                compFormFiledOnDateSearchIndex = pdfText.find('filed  on') + 9
            else:
                compFormFiledOnDateSearchIndex = pdfText.find('filed on') + 8
            
            while(pdfText[compFormFiledOnDateSearchIndex] == ' '):
                compFormFiledOnDateSearchIndex = compFormFiledOnDateSearchIndex + 1
            
            compFormFiledOnDateStartIndex = compFormFiledOnDateSearchIndex
            compFormFiledOnDateSearchEndIndex = compFormFiledOnDateSearchIndex
            while(pdfText[compFormFiledOnDateSearchEndIndex]!='.'):
                compFormFiledOnDateSearchEndIndex = compFormFiledOnDateSearchEndIndex + 1
            compFormFiledOnDateEndIndex = compFormFiledOnDateSearchEndIndex
            
            compFormFiledOnDate = pdfText[compFormFiledOnDateStartIndex:compFormFiledOnDateEndIndex]
        except Exception:
            traceback.print_exc()
            compFormFiledOnDate = 'NA'
    else:
        compFormFiledOnDate = 'NA'
    return compFormFiledOnDate


def getCompExhibitDets(pdfText):
    compExhibitDets = {}
    if(pdfText.find('specified:')!=-1):
        try:
            compExhibitSearchIndex = pdfText.find('specified:') + 10
            if(pdfText.find('Exhibit',compExhibitSearchIndex)!= -1):
                compExhibitNumberSearchIndex = pdfText.find('Exhibit',compExhibitSearchIndex) + 7
            else:
                compExhibitDets = {'number':'NA','date':'NA'}
                return compExhibitDets
            
            while(pdfText[compExhibitNumberSearchIndex]==' '):
                compExhibitNumberSearchIndex = compExhibitNumberSearchIndex + 1
            compExhibitNumberStartIndex = compExhibitNumberSearchIndex
            compExhibitNumberSearchEndIndex = compExhibitNumberStartIndex
            while(pdfText[compExhibitNumberSearchEndIndex]!= ' '):
                compExhibitNumberSearchEndIndex = compExhibitNumberSearchEndIndex + 1
            
            compExhibitNumberEndIndex = compExhibitNumberSearchEndIndex
            
            compExhibitNumber = pdfText[compExhibitNumberStartIndex:compExhibitNumberEndIndex]
            compExhibitNumberText = "Exhibit " + compExhibitNumber
            
            compExhibitDateSearchIndex = compExhibitNumberEndIndex
            while(pdfText[compExhibitDateSearchIndex]== ' '):
                compExhibitDateSearchIndex = compExhibitDateSearchIndex + 1
            compExhibitDateStartIndex = compExhibitDateSearchIndex
            compExhibitDateSearchEndIndex = compExhibitDateStartIndex
            
            while(pdfText[compExhibitDateSearchEndIndex]!= ','):
                compExhibitDateSearchEndIndex = compExhibitDateSearchEndIndex + 1
            compExhibitDateSearchEndIndex = compExhibitDateSearchEndIndex+ 1
            
            while(pdfText[compExhibitDateSearchEndIndex]== ' '):
                compExhibitDateSearchEndIndex = compExhibitDateSearchEndIndex + 1
            while(pdfText[compExhibitDateSearchEndIndex]!= ' '):
                compExhibitDateSearchEndIndex = compExhibitDateSearchEndIndex + 1
                
            compExhibitDateEndIndex = compExhibitDateSearchEndIndex
            compExhibitDate = pdfText[compExhibitDateStartIndex:compExhibitDateEndIndex]
            compExhibitDets = {'number':compExhibitNumberText,'date':compExhibitDate}
        except Exception:
            traceback.print_exc()
            compExhibitDets = {'number':'NA','date':'NA'}
    else:
        compExhibitDets = {'number':'NA','date':'NA'}
    return compExhibitDets
       
   
def getCompDets(numberOfPDFRead,pdfText):

    print(pdfText)
    pdfText = pdfText.replace('\n', '')
    print("Removing new line")
    print(pdfText)
    
    compMainDate = getCompDate(pdfText)
    print(compMainDate)
    compName = getCompName(pdfText)
    print(compName)
    compReq = getCompReq(pdfText)
    print(compReq)
    compFormFiled = getCompFormFiled(pdfText)
    print(compFormFiled)
    compFormFiledOnDate = getCompFormFiledOnDate(pdfText)
    print(compFormFiledOnDate)
    compExhibitDets = getCompExhibitDets(pdfText)
    compExhibitNumber = compExhibitDets["number"]
    compExhibitDate = compExhibitDets["date"]
    print(compExhibitDets["number"])
    print(compExhibitDets["date"])
    
    compDets = [str(numberOfPDFRead),compMainDate.strip(),compName.strip(),compReq.strip(),compFormFiled.strip(),compFormFiledOnDate.strip(),compExhibitNumber.strip(),compExhibitDate.strip()]
    return compDets

def writeExcel(totalDocDetList):
    workbook = xlsxwriter.Workbook(dowloadPath)
    worksheet = workbook.add_worksheet("Company Details sheet") 
    row = 0
    col = 0
    print("InWriteExcel")
    for docDetList in totalDocDetList:
        colNum = 0
        for compDet in docDetList:
            print(compDet)
            
            worksheet.write(row, col + colNum, compDet) 
            colNum = colNum+1
            
        row += 1
    workbook.close()

def main():
    noOfPdfsToRead = 10
    totalDocDetList = []
    for i in range(noOfPdfsToRead):
        print(i)
        numberOfPDFRead = i
        #numberOfPDFRead = 1154
        pdfText = readPDF(numberOfPDFRead)
        
        if(pdfText!="-1"):
            docDets = getCompDets(numberOfPDFRead,pdfText)
            totalDocDetList.append(docDets)
        
    #print(totalDocDetList)
    writeExcel(totalDocDetList)

main()

    

