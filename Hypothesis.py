#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlwt
from xlwt import Workbook

import json

with open('afinn.json') as f:
  data = json.load(f)

wb = Workbook(encoding='ascii')
sheet1 = wb.add_sheet('Standards Data')
sheet2 = wb.add_sheet('Sentiment Count')

f_amul=open("HUL 2018-2019_Annual Report.txt", "r")
#f_colgate=open("Colgate 2018-2019_Annual Report.txt", "r")
#f_dabur=open("Dabur 2018-19_Annual Report.txt", "r")
#f_godrej=open("Godrej 2018-2019_Annual Report.txt", "r")
#f_hul=open("HUL 2018-2019_Annual Report.txt", "r")
#f_itc=open("ITC 2018-2019 Annual Report.txt", "r")
#f_marico=open("Marico 2018-2019_Annual Report.txt", "r")
#f_nestle=open("Nestle 2017-2018_Annual Report.txt", "r")
#f_pnj=open("P&G 2018-2019_Annual Report.txt", "r")
f_glossary=open("glossary.txt", "r")


items_amul = f_amul.read().split(".")
#items_colgate = f_colgate.read().split(".")
#items_dabur = f_dabur.read().split(".")
#items_godrej = f_godrej.read().split(".")
#items_hul = f_hul.read().split(".")
#items_itc = f_itc.read().split(".")
#items_marico = f_marico.read().split(".")
#items_nestle = f_nestle.read().split(".")
#items_pnj = f_pnj.read().split(".")

fin = []
soc = []
env = []
item_g = f_glossary.read().split("|")
fin=item_g[0].split(",")
soc=item_g[1].split(",")
env=item_g[2].split(",")
#print(fin,soc,env)


Amul=[str(i) for i in items_amul]
standards=[env,soc,fin]
standards_title=['env','soc','fin']

for i in range(3):
    sheet1.write(0, i, standards_title[i])
    sheet2.write(0, i, standards_title[i])
    count = 1
    for j in range(len(Amul)):
        sentence=Amul[j].split(' ')
        sentiment_count =0
        flag =0
        for k in sentence: #the words
            if k in data:
                sentiment_count = sentiment_count + data[k]
            for l in standards[i]: #glossary words
                if(l==k):
                    flag=1
                    break
        if (flag == 1):
            sheet1.write(count, i, Amul[j]) # sentences into excel
            sheet2.write(count, i, sentiment_count/len(sentence))
            count = count+1

wb.save('Standards Data.xls')
