# This class loads the mapping file, which has customer side
# ERP Product name mapped to the equavalent name in Govt Portal
from numpy import maximum
import pandas as pd
from pandas.core.indexes.base import Index


class MappingFileLoader:
    data = None
    data1 = None

    def __init__(self):
        global data,data1
        xls = pd.ExcelFile('config/mapping_file.xlsx')
        data = xls.parse(xls.sheet_names[0])
        data1 = xls.parse(xls.sheet_names[1])
        #print("hello")

    @staticmethod
    def fetch_Product_Mapping():
        global data
        mappedData = []
        df = pd.DataFrame(data)
        for index, row in df.iterrows():
            rowData = {row['IN EXCEL SHEET']: row['IN UI']}
            mappedData.append(rowData)
        return mappedData

    @staticmethod    
    def fetch_Parse_param_mapping():
        global data1
        mappedData1 = []
        df1 = pd.DataFrame(data1)
        for index, row in df1.iterrows():
            rowData1 = {row['param.name']: row['param.value']}
            mappedData1.append(rowData1)
            #print(rowData1)
        return mappedData1
        
        
# if __name__ == "__main__":
mfl = MappingFileLoader()
mfl.fetch_Parse_param_mapping()