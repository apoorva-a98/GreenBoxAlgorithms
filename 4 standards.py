#--psudo code--
#
# for nouns:
#     group by glossary
#     give the same standard and substandard as reporting requirements
#     store the remaining words in AbstractNouns = []

#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame
# from google.colab import drive
# drive.mount('/gdrive')

# from google.colab import drive
# drive.mount('/content/drive')

# INITIALIZING SPACY AND ITS 'en' MODEL
# !python -m spacy download en_core_web_lg
# import en_core_web_lg
# nlp = en_core_web_lg.load()
nlp = spacy.load("en_core_web_md")

# path = "/content/drive/My Drive/GreenBoxAlgorithms/"

frequency=1
text=2
lemma=3
pos=3
sentiment=4

#READ Nouns
def read_nouns():
        sheet = []
        sheet = pd.read_excel(path+'/Nouns.xlsx', usecols='B:F')
        sheet = sheet.values.tolist()
        return sheet

# READ Glossary
def read_glossary():
    sheet=[]
    sheet= pd.read_excel(path+'/Glossary.xlsx')
    sheet = sheet.values.tolist()
    return sheet

AbstractNouns=[]

# GROUPING DESCRIPTIVE WORDS
def sort_standards(Nouns, Glossary):
    while(len(Nouns)>0):
        for word in Glossary:
            flag = 0
            token = nlp(Nouns[0][text]+' '+word[2])
            if int(token[0].similarity(token[1])*100) >= 40:
                new_noun=[]
                new_noun.append(word[0])
                new_noun.append(word[1])
                new_noun.append(Nouns[0][text])
                Glossary.append(new_noun)
                break
            else:
                AbstractNouns.append(Noun[0])
                break
        Nouns.pop(0)
    return Glossary


#CREATE GLORRARY
def create_glossary(glossary,nouns):
    df_words = pd.DataFrame(glossary)
    df_words.columns=['standard','sub-standard','text']

    df_nouns = pd.DataFrame(nouns)
    df_nouns.columns=['frequency','text','lemma','pos','afinn sentiment']

    # glossary to excel
    with pd.ExcelWriter(path+"/FinalGlossary.xlsx") as writer:
        df_words.to_excel(writer)
    writer.save()

    with pd.ExcelWriter(path+"/AbstractNouns.xlsx") as writer:
        df_nouns.to_excel(writer)
    writer.save()

create_glossary(sort_standards(read_nouns(),def read_glossary()),AbstractNouns)
