#--psudo code--
#
# for glossary:
#     open glossary.xls
#     for each element in verbs, adverbs and adjectives:
#         find all setences with these WORDS
#         pos to these sentences
#         check if the noun is the sentence is not in glossary_nouns
#         append the nouns to the nouns list

#!/usr/bin/env python3

# from nltk.corpus import wordnet
import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame


# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# READING FAULTERED GLOSSARY EXCEL
filepath="companies_glossary/glossary.xlsx"
def read_sheet(pos):
    sheet=[]
    sheet= pd.read_excel(filepath, sheet_name=pos)
    sheet = sheet.values.tolist()
    sheet = [item.lower() for sublist in sheet for item in sublist]
    return sheet

Nouns = read_sheet("Nouns")
Verbs = read_sheet("Verbs")
Adverbs = read_sheet("Adverbs")
Adjectives = read_sheet("Adjectives")


# CLEAN READING
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']


class reports:
    def __init__(self, name, path):
        self.company = name
        self.filepath = path

    # READING COMPANY REPORTS
    def read_file(self):
        report= open(self.filepath, "r", encoding = "ISO-8859-1")
        report_text= report.read()
        return report_text
    #print(read_file())


    # TOKANIZING REPORT
    def tokenify_glossary(self,report):
        buff = ''
        sentences=[]
        for letter in report:
            letter=letter.lower()
            if letter in avoid:
                if buff != '':
                    sentences.append(buff)
                buff = ''
            elif (buff is not None):
                buff += letter
        if buff is not None:
            sentences.append(buff)
            buff=''
        print(sentences)
        return sentences

    #SEARCHING WITHIN THE FAULTERS
    def search_within(self,sentence,pos):
        sentence=sentence.split(' ')
        for word in pos:
            if word in sentence:
                return 1

    #GLORRARY SRANDARDS
    def find_faulters(self,sentences):
        for sentence in sentences:
            if self.search_within(sentence,Verbs) == 1 or self.search_within(sentence,Adverbs) == 1 or self.search_within(sentence,Adjectives) == 1:
                doc = nlp(sentence)
                for token in doc:
                    if (token.pos_ == 'NOUN' or token.pos_ == 'PROPN') and token.text not in Nouns:
                        Nouns.append(token.text)
        return Nouns
    #print(find_faulters(tokenify_glossary(read_file())))

    #CREATE GLORRARY
    def create_glossary(self, words):
        df_words = pd.DataFrame(words)
        df_words.columns=['text']

        #glossary to excel
        with pd.ExcelWriter("companies_glossary/refinedglossary.xlsx") as writer:
            df_words.to_excel(writer, sheet_name='Nouns')
        writer.save()
        return df_words


HUL = reports("HUL", "HUL 2018-2019_Annual Report.txt")
print(HUL.create_glossary(HUL.find_faulters(HUL.tokenify_glossary(HUL.read_file()))))

# Colgate = reports("Colgate", "Colgate 2018-2019_Annual Report.txt")
# print(Colgate.create_glossary(Colgate.find_faulters(Colgate.tokenify_glossary(Colgate.read_file()))))
#
# ITC = reports("ITC", "ITC 2018-2019 Annual Report.txt")
# print(ITC.create_glossary(ITC.find_faulters(ITC.tokenify_glossary(ITC.read_file()))))
#
# Dabur = reports("Dabur", "Dabur 2018-19_Annual Report.txt")
# print(Dabur.create_glossary(Dabur.find_faulters(Dabur.tokenify_glossary(Dabur.read_file()))))
#
# Godrej = reports("Godrej", "Godrej 2018-2019_Annual Report.txt")
# print(Godrej.create_glossary(Godrej.find_faulters(Godrej.tokenify_glossary(Godrej.read_file()))))
#
# Marico = reports("Marico", "Marico 2018-2019_Annual Report.txt")
# print(Marico.create_glossary(Marico.find_faulters(Marico.tokenify_glossary(Marico.read_file()))))
#
# Nestle = reports("Nestle", "Nestle 2017-2018_Annual Report.txt")
# print(Nestle.create_glossary(Nestle.find_faulters(Nestle.tokenify_glossary(Nestle.read_file()))))
#
# PnG = reports("PnG", "P&G 2018-2019_Annual Report.txt")
# print(PnG.create_glossary(PnG.find_faulters(PnG.tokenify_glossary(PnG.read_file()))))
