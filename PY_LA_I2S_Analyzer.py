# https://stackoverflow.com/questions/8396874/sine-function-on-16-bit-microcontroller
# https://www.cnblogs.com/hankleo/p/10789401.html 
# https://openhome.cc/Gossip/Python/NumericType.html
# https://ithelp.ithome.com.tw/articles/10203129
# https://officeguide.cc/python-openpyxl-excel-charts-tutorial-examples/
# https://openpyxl.readthedocs.io/en/stable/styles.html#applying-styles
# https://www.itread01.com/content/1543829948.html
# https://docs.python.org/2/library/wave.html
# https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/362015/ 
# https://blog.xuite.net/yeehonge/wretch/36157450

import sys
import csv
import os
import openpyxl
from openpyxl.chart import LineChart, Reference
import matplotlib.pyplot as plt
from math import ceil
import wave
import La_saleae

#CONFIGURATIONS
SALEAE = "SALEAE"
device_type = SALEAE

#DEFINITIONS
COL_INDEX_TIME = 1
COL_INDEX_DATA_24bit_HEX_L = 2  
COL_INDEX_DATA_24bit_HEX_R = 3  
COL_INDEX_DATA_24bit_DEC_L = 4  
COL_INDEX_DATA_24bit_DEC_R = 5 
COL_INDEX_DATA_16bit_DEC_L = 6  
COL_INDEX_DATA_16bit_DEC_R = 7

COL_INDEX_SAMPLING_RATE = 9
COL_INDEX_WAVEFORM = COL_INDEX_SAMPLING_RATE

MAX_PERIOD_IN_ms = 100

#===========================#    
def Init_Excel_Table(row, sheet):
    if(device_type == SALEAE):
        colInfo = La_saleae.Init_Excel_Table(row, sheet):
    return colInfo;

def convert_hexstr2hexval(valStr):
    CONST_SIGN_BIT      = 0x800000
    CONST_SIGN_VAL_MAX  = 0x7FFFFF
    
    val_24 = int(valStr, 16)
    if(val_24 & CONST_SIGN_BIT):
        val_24 = (val_24 & CONST_SIGN_VAL_MAX) - CONST_SIGN_VAL_MAX - 1 
    return val_24
    
def is_right_channel(string):
    if( string.find("1") != -1 ):
        return 1;
    else:
        return 0;

def raw_convert_bytearr(rawList):
    retList = []
    for i in range(0, len(rawList)):
        retList.append( rawList[i] & 0xFF)
        retList.append( (rawList[i] >> 8) & 0xFF ) 
    return bytearray(retList)
    
def dual_raw_convert_bytearr(rawList1, rawList2):
    retList = []
    for i in range(0, len(rawList1)):
        retList.append( rawList1[i] & 0xFF)
        retList.append( (rawList1[i] >> 8) & 0xFF ) 
        retList.append( rawList2[i] & 0xFF)
        retList.append( (rawList2[i] >> 8) & 0xFF )        
    #print(retList)
    return bytearray(retList)


def data_plot(ws, dataLen):
    chart = LineChart()
    chart.title = "Data - left, 24bit"
    chart.style = 11
    chart.height = 8
    chart.width = 32    
    chart.y_axis.title = 'value'
    chart.x_axis.title = 'time'
    data = Reference(ws, min_col=COL_INDEX_DATA_24bit_DEC_L, min_row=1, max_col=COL_INDEX_DATA_24bit_DEC_L, max_row = dataLen)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "I4")
    s = chart.series[0]
    s.smooth = True
    s.graphicalProperties.line.width = 10

    chart = LineChart()
    chart.title = "Data - right, 24bit"
    chart.style = 11
    chart.height = 8
    chart.width = 32
    chart.y_axis.title = 'value'
    chart.x_axis.title = 'time'
    data = Reference(ws, min_col=COL_INDEX_DATA_24bit_DEC_R, min_row=1, max_col=COL_INDEX_DATA_24bit_DEC_R, max_row = dataLen)
    chart.add_data(data, titles_from_data=True)    
    ws.add_chart(chart, "I20")
    s = chart.series[0]
    s.smooth = True
    s.graphicalProperties.line.width = 10

    chart = LineChart()
    chart.title = "Data - left, 16bit"
    chart.style = 11
    chart.height = 8
    chart.width = 32    
    chart.y_axis.title = 'value'
    chart.x_axis.title = 'time'
    data = Reference(ws, min_col=COL_INDEX_DATA_16bit_DEC_L, min_row=1, max_col=COL_INDEX_DATA_16bit_DEC_L, max_row = dataLen)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "I36")
    s = chart.series[0]
    s.smooth = True
    s.graphicalProperties.line.width = 10

    chart = LineChart()
    chart.title = "Data - right, 16bit"
    chart.style = 11
    chart.height = 8
    chart.width = 32    
    chart.y_axis.title = 'value'
    chart.x_axis.title = 'time'
    data = Reference(ws, min_col=COL_INDEX_DATA_16bit_DEC_R, min_row=1, max_col=COL_INDEX_DATA_16bit_DEC_R, max_row = dataLen)
    chart.add_data(data, titles_from_data=True)    
    ws.add_chart(chart, "I52")   
    s = chart.series[0]
    s.smooth = True
    s.graphicalProperties.line.width = 10    

def dictionaryInit():
    dataDict = {}    
    dataDict["Left_24bit"] = []
    dataDict["Right_24bit"] = []
    dataDict["Left_16bit"] = []
    dataDict["Right_16bit"] = []
    
    timeDict = {}
    timeDict["prev"] = 0
    timeDict["present"] = 0
    timeDict["diffAcc"] = 0
    timeDict["SamplingRate"] = 0
    return dataDict, timeDict 

def get_first_row(rows):
    if(device_type == SALEAE):
        headers = La_saleae.get_first_row(rows)
    return headers
    
def get_value(hexStr):
    if(device_type == SALEAE):
        val = La_saleae.get_value(hexStr)
    return val;

def AppStart():
    inputFileName = (sys.argv[1]);
    sr_khz = int(sys.argv[2]);
    channelNum = int(sys.argv[3]);
    newFileName = "output.xls"
    wavFileName = "output.wav"
    
    try:
        # open input file with CSV format
        inFile = open(inputFileName, 'r')
        # read CSV data in rows
        rows = csv.reader(inFile, delimiter=',')
        
        # Creae a excel object 
        wb = openpyxl.Workbook()
        # Creat a sheet
        sheet = wb.create_sheet(index=0)
        
        wavFile = wave.open( wavFileName , 'wb')
        wavFile.setparams((channelNum, 2, sr_khz*1000, 0, 'NONE', 'NONE'))
        #channelNum / bytes per sample / samping rate / compression type
            
    except:
        exit(1)
    
    headers = get_first_row(rows)
    dataDict,timeDict =  dictionaryInit();
    colInfo = Init_Excel_Table(headers,sheet)

    # Calculate the ending index
    readIdx = 1
    endIdx = readIdx + (MAX_PERIOD_IN_ms/1000) * (sr_khz * 1000)
    endIdx = int(endIdx)

    for row in rows:

        val = get_value(row[colInfo["data"]])
        
        time_str = row[colInfo["start_time"]]
        timeDict["prev"] = timeDict["present"]
        timeDict["present"] = float(time_str)

        if( readIdx > 1):
            timeDict["diffAcc"] = timeDict["diffAcc"] + ( timeDict["present"] - timeDict["prev"])
        
        if(is_right_channel( row[colInfo["channel"]] )):
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_24bit_HEX_R, val["str"], sheet) 
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_24bit_DEC_R, val["24bit"], sheet)
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_16bit_DEC_R, val["16bit"], sheet)
            dataDict["Right_24bit"].append(val["24bit"])
            dataDict["Right_16bit"].append(val["16bit"])
        else:
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_24bit_HEX_L, val["str"], sheet) 
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_24bit_DEC_L, val["24bit"], sheet)
            fill_data_into_excel(int((readIdx + 3)>>1), COL_INDEX_DATA_16bit_DEC_L, val["16bit"], sheet)
            dataDict["Left_24bit"].append(val["24bit"])
            dataDict["Left_16bit"].append(val["16bit"])         

        readIdx += 1     
        
        if(readIdx > endIdx):
            break;
    
    timeDict["SamplingRate"]  = (readIdx - 1) / timeDict["diffAcc"] / 2
    fill_data_into_excel(2, COL_INDEX_SAMPLING_RATE, timeDict["SamplingRate"], sheet)

    data_plot(sheet, readIdx / channelNum)

    if(channelNum == 1):
        pcmdata = raw_convert_bytearr(dataDict["Left_16bit"])
    else:
        pcmdata = dual_raw_convert_bytearr(dataDict["Left_16bit"], dataDict["Right_16bit"])

    wavFile.writeframesraw(pcmdata)
    inFile.close()
    wb.save(newFileName)
    wavFile.close()

    print("Success")

        
if __name__ == '__main__':
    AppStart()