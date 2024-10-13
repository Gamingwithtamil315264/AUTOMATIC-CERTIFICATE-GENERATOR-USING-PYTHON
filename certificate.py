from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert
import os
pd.set_option('display.max_rows',None)
pd.set_option('display.max_columns',None)
j=0
ans=pd.read_csv('data.csv')

n = min(len(ans['NAME']), len(ans['Dept']), len(ans['YEAR']))
print(ans.head(10))

def mkw(n):
    tpl=DocxTemplate("temp.docx")
    context={i+1:{"np1":str(ans['NAME'][i]),"dept":str(ans['Dept'][i]),"year":str(ans['YEAR'][i]),'event':str(ans['e'][i])}for i in range (n)}
    tpl.render(context[n])
    a=f"{i}.docx"
    directory="ONSPOT"
    p=os.path.join("certificate",directory)
    if os.path.exists(p)==False:
        os.mkdir(p)
    o=os.path.join(p,a)
    print(f"docx:{a}")
    tpl.save(o)
    k=f"{i}.pdf"
    print(f"pdf:{k}")
    directory_="ONSPOT pdf"
    y=os.path.join("certificate",directory_)
    if os.path.exists(y)==False:
        os.mkdir(y)
    f=os.path.join(y,k)
    convert(o,f)
print(len(ans['NAME']))
print(len(ans['Dept']))
print(len(ans['YEAR']))

for i in range(1,n+1):    
    if pd.isna(ans['NAME'][i-1]):
        print(f"skip{i}")
        continue     
    mkw(i)
