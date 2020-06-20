#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame

file_path = [0]*8
file_path[0]='companies_glossary/HUL.xlsx'
file_path[1]='companies_glossary/Colgate.xlsx'
file_path[2]='companies_glossary/ITC.xlsx'
file_path[3]='companies_glossary/Dabur.xlsx'
file_path[4]='companies_glossary/Godrej.xlsx'
file_path[5]='companies_glossary/Marico.xlsx'
file_path[6]='companies_glossary/Nestle.xlsx'
file_path[7]='companies_glossary/PnG.xlsx'

PoS=['Nouns','Verbs','Adverbs','Adjectives']

#READ Nouns
Nouns=[]
for company in file_path:
    sheet = []
    sheet = pd.read_excel(company, sheet_name='Nouns', usecols='B:I')
    sheet = sheet.values.tolist()
    Nouns.extend(sheet)
# print(Nouns)

def reduce_glossary(sorted_words):
    glossary=[]
    while(len(sorted_words)>0):
        word_frequency=[]
        while(len(sorted_words)>1 and sorted_words[0][1]==sorted_words[1][1]):
            sorted_words[0][0]=sorted_words[0][0]+sorted_words[1][0]
            sorted_words=np.delete(sorted_words, 1, 0)
        word_frequency.extend(sorted_words[0])
        glossary.append(word_frequency)
        sorted_words=np.delete(sorted_words, 0, 0)
    return glossary

#Write Nouns
unsorted_nouns = np.array(Nouns)
sorted_nouns=unsorted_nouns[unsorted_nouns[:, 1].argsort()]
sorted_nouns=reduce_glossary(sorted_nouns)
df_nouns = pd.DataFrame(sorted_nouns)
df_nouns.columns=['frequency','text','lemma','pos','eng-tag','dependency','afinn sentiment','mcdonals sentiment']
print(df_nouns)

#glossary to excel
with pd.ExcelWriter("companies_glossary/Nouns.xlsx") as writer:
    df_nouns.to_excel(writer, sheet_name='Nouns')
writer.save()
