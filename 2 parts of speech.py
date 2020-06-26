#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame

#GET COMAPNY GLOSSARIES
file_path = [0]*8
file_path[0]='companies_glossary/HUL.xlsx'
file_path[1]='companies_glossary/Colgate.xlsx'
file_path[2]='companies_glossary/ITC.xlsx'
file_path[3]='companies_glossary/Dabur.xlsx'
file_path[4]='companies_glossary/Godrej.xlsx'
file_path[5]='companies_glossary/Marico.xlsx'
file_path[6]='companies_glossary/Nestle.xlsx'
file_path[7]='companies_glossary/PnG.xlsx'

class standards_and_sentiments:
    def __init__(self, POS):
        self.partofspeech = POS

    #READ PoS
    def read_partofspeech(self):
        master_list=[]
        for company in file_path:
            sheet = []
            sheet = pd.read_excel(company, sheet_name=self.partofspeech, usecols='B:F')
            sheet = sheet.values.tolist()
            master_list.extend(sheet)
        return master_list
    # print(Nouns)


    #REMOVE Duplicates
    def reduce_glossary(self,sorted_words):
        glossary=[]
        while(len(sorted_words)>0):
            keyword=[]
            while(len(sorted_words)>1 and sorted_words[0][2]==sorted_words[1][2]):
                sorted_words[0][0]=sorted_words[0][0]+sorted_words[1][0]
                sorted_words=np.delete(sorted_words, 1, 0)
            if int(sorted_words[0][0]) >= 100:
                keyword.extend(sorted_words[0])
                glossary.append(keyword)
            sorted_words=np.delete(sorted_words, 0, 0)
        return glossary


    #WRITE PoS
    def write_partofspeech(self):
        unsorted_words = np.array(self.read_partofspeech())
        sorted_words=unsorted_words[unsorted_words[:, 1].argsort()]
        sorted_words=self.reduce_glossary(sorted_words)
        df_words = pd.DataFrame(sorted_words)
        df_words.columns=['frequency','text','lemma','pos','afinn sentiment']
        # print(df_words)

        #glossary to excel
        with pd.ExcelWriter("companies_glossary/"+self.partofspeech+".xlsx") as writer:
            df_words.to_excel(writer, sheet_name=self.partofspeech)
        writer.save()


Nouns= standards_and_sentiments("Nouns")
Nouns.write_partofspeech()

Verbs= standards_and_sentiments("Verbs")
Verbs.write_partofspeech()

Adverbs= standards_and_sentiments("Adverbs")
Adverbs.write_partofspeech()

Adjectives= standards_and_sentiments("Adjectives")
Adjectives.write_partofspeech()
