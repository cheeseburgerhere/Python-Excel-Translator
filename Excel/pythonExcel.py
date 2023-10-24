import pandas as pd
import translators as ts

   #to detect language use 'auto' for from_langCode
def excelTranslator(excel_sheet: str, to_langCode : str, from_langCode : str, export=False, sheetName="Sheet1"): 
    df = pd.read_excel(excel_sheet)

    for k in range(0,int(df.size/df.iloc[1].size)):
        for i in range(0,df.iloc[1].size):
            series=df.iloc[k]
            word=series.iloc[i]
            if(type(word)==type('str')):
                series.iloc[i]=ts.translate_text(query_text=word, to_language=to_langCode)
   
    column_name=df.columns
      
    for i in range(0,df.iloc[0].size):
        newName=ts.translate_text(query_text=column_name[i], to_language=to_langCode,from_language=from_langCode)
        column_name_map={column_name[i]:newName}
        df.rename(columns=column_name_map,inplace=True)
    
    
    if(not export):
        return df
    else:
        index=str(excel_sheet).find(".xlsx")
        excelName=excel_sheet[:index]
        df.to_excel(excelName+"_translated.xlsx", sheet_name=sheetName, index=False)


#just translates and returns
print(excelTranslator(excel_sheet="excel sheet.xlsx",to_langCode="tr",from_langCode="auto"))

#translates and exports
#excelTranslator(excel_sheet="excel sheet.xlsx",to_langCode="tr",from_langCode="auto",export=True,sheetName="MySheet")

