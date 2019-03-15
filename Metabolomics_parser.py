#!/usr/bin/env python3
import os
import pandas as pd


#destiny file
destiny1 = '/home/rafael/Bioinfo/metabolomics.xls'

# The top argument for walk, this is the were all .D directories are.
topdir = '/home/rafael/Bioinfo/Raw/XLSX'
 
# The extension to search for
exten = '.xlsx'

#open "machine" that writes to the destiny file 
writer1 = pd.ExcelWriter(destiny1)


for dirpath, dirnames, files in os.walk(topdir):#"walk" in directories and get names files and more information
    dirnames.sort() #ordenates the directories names
    for name in files:
        tabname_l = name.split('.')
        tabname = tabname_l[0]  
        output_csv = tabname + '_clean.csv'      
        if name.lower().endswith(exten):
            print(tabname)
            wb = pd.read_excel(os.path.join(dirpath, name), usecols="T,V,AH,AJ,AZ,BJ,BS,BT,BW,DO") #collumns with important data 
            df = pd.DataFrame(wb)
            new_order=[9,1,3,4,5,7,8,2,6,0]#put sample anme, compound names etc in convenient order
            df2 = df[df.columns[new_order]]
            df2.to_excel(writer1,sheet_name=str(tabname), index=False,header=False, startcol=0) #put all together in a single file
            df2.to_csv(output_csv, index=False, header=False)#create individual .csv files         
 
writer1.save()

