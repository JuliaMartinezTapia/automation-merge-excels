import os
from pathlib import Path
from functools import reduce
import pandas as pd

#List to store the dataframes extacted from the excel files located in current directory (same directory than this .py file)

list_df = []

#Loop to go through the elements (directories and files) within the current directory,find the xlsx files, 
#appply a mask to eliminate spaces and Nans in the first column (on which the merge will be subsequently performed).

for item in os.listdir(Path.cwd()):
    
    if "xlsx" in item:

        df = pd.read_excel(item, engine = "openpyxl")
        name_col = df.columns[0]
        df = df[((df[name_col]) != " ")][~df[name_col].isna()]
        list_df.append(df)

list_df  
print(list_df)

#Merging the dataframes in the list using funcional programming (reduce).

df_merged = reduce(lambda  left,right: pd.merge(left,right,on=[name_col],
                                            how='outer'), list_df)

#create a new directory in current directory to store the output excel

try:
    Path(str(Path.cwd()) + "./output").mkdir()
except: 
    pass
    
#Download the merged dataframe to an excel spreadsheet created in the current directory

df_merged.to_excel("./output/merged_df.xlsx")