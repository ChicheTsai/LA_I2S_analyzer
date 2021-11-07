def get_first_row(rows):
    headers = next(rows) 
    return headers

def get_value(hexStr):
    val = {"str":hexStr, "32bit": 0, "24bit": 0, "16bit":0}
    CONST_SIGN_BIT      = 0x800000
    CONST_SIGN_VAL_MAX  = 0x7FFFFF
    
    val["24bit"] = int(val["str"], 16)
    if(val["24bit"] & CONST_SIGN_BIT):
        val["24bit"] = (val["24bit"] & CONST_SIGN_VAL_MAX) - CONST_SIGN_VAL_MAX - 1 
    val["16bit"] = val["24bit"] >> 8
    
    return val

def fill_data_into_excel(r, c, data,sheet):
    sheet.cell(r, c).value = data    
    
def Init_Excel_Table(row, sheet):    
    colInfo = {}
    fill_data_into_excel(1, COL_INDEX_TIME, 'start_time', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_24bit_HEX_L, 'HEX_24_L', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_24bit_HEX_R, 'HEX_24_R', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_24bit_DEC_L, 'DEC_24_L', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_24bit_DEC_R, 'DEC_24_R', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_16bit_DEC_L, 'DEC_16_L', sheet)
    fill_data_into_excel(1, COL_INDEX_DATA_16bit_DEC_R, 'DEC_16_R', sheet)     
    fill_data_into_excel(1, COL_INDEX_SAMPLING_RATE, 'SamplingRate(Hz)', sheet)
    
    for i in range(0,len(row)):
        if(row[i] == "start_time"):
            colInfo["start_time"] = (i)
        elif(row[i] == "channel"):
            colInfo["channel"] = (i)
        elif(row[i] == "data"):
            colInfo["data"] = (i) 
    return colInfo