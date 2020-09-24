import pandas as pd
import xlrd
from glob import glob
from dateutil import parser
import datetime





df=pd.DataFrame(data=[[1,2,3],[4,5,6],[7,8,9]],columns=list('abc'))
ss=df
df2=pd.DataFrame(data=[[2,4,8]],columns=list('abc'))
df3=ss.append(df,ignore_index=True)
print(df3)
