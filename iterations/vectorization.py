#--psudo code--
#
# for nouns:
#     group by reporting requirements
#     give the same standard and substandard as reporting requirements
#     store the remaining words in AbstractNouns = []
#
# for verbs, adverbs, adjectives, abstractnouns:
#     group by words with afinn sentiment
#     give token id
#     truncate the words without sentiment


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

class standards_and_sentiments:

    def __init__(self, POS):
        self.partofspeech = POS

    #READ PoS
    def read_partofspeech(self):
            sheet = []
            sheet = pd.read_excel("companies_glossary/"+self.partofspeech+".xlsx", usecols='B:F')
            #sheet = pd.read_excel(path+'/'+self.partofspeech+'.xlsx', sheet_name="Nouns", usecols='B:I')
            sheet = sheet.values.tolist()
            unsorted_words = np.array(sheet)
            sorted_words=unsorted_words[unsorted_words[:, 4].argsort()]
            sorted_words=sorted_words.tolist()
            return sorted_words
    #print(read_partofspeech("Nouns"))


    # VECTORISING GLOSSARY
    def sorting_similar_words(self,words):
        i=0
        token_id=1
        while(i < len(words) and words[i][4] is not None):
            count=0
            for j in range(len(words)):
                token = nlp(words[i][2]+' '+words[j][2])
                if int(token[0].similarity(token[1])*100) >= 50 and words[j][4 is None]:
                    print(token[0].lemma_, token[1].lemma_, int(token[0].similarity(token[1])*100))
                    words[j].append(token_id)
                    words[j].append(token[0].similarity(token[1])*100)
                    place_holder=words[j]
                    words[j]=words[i+1+count]
                    words[i+1+count]=place_holder
                    count=count+1
            words[i].append(token_id)
            words[i].append(" ")
            token_id=token_id+1
            i=i+count+1
        df_words = pd.DataFrame(words)
        df_words.columns=['frequency','text','lemma','pos','afinn sentiment','token_id','similarity']

        #glossary to excel
        with pd.ExcelWriter(path+'/'+self.partofspeech+"vectorised.xlsx") as writer:
            df_words.to_excel(writer, sheet_name=self.partofspeech)
        writer.save()


# ReportingRequirements=standards_and_sentiments("ReportingRequirements")
# ReportingRequirements.sorting_similar_words(ReportingRequirements.read_partofspeech())

ReportingRequirements=standards_and_sentiments("Reporting Requirements")
print(ReportingRequirements.read_partofspeech())

# Nouns=standards_and_sentiments("Nouns")
# Nouns.sorting_similar_words(Nouns.read_partofspeech())
#
# Verbs=standards_and_sentiments("Verbs")
# Verbs.sorting_similar_words(Verbs.read_partofspeech())
#
# Adverbs=standards_and_sentiments("Adverbs")
# Adverbs.sorting_similar_words(Adverbs.read_partofspeech())
#
# Adjectives=standards_and_sentiments("Adjectives")
# Adjectives.sorting_similar_words(Adjectives.read_partofspeech())
#
# ReportingRequirements=standards_and_sentiments("ReportingRequirements")
# ReportingRequirements.sorting_similar_words(ReportingRequirements.read_partofspeech())
