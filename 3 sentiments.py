#--psudo code--
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

frequency=1
text=2
lemma=3
pos=3
sentiment=4

class sentiments:
    def __init__(self, POS):
        self.partofspeech = POS

    #READ PoS
    def read_partofspeech(self):
            sheet = []
            sheet = pd.read_excel(path+'/'+self.partofspeech+".xlsx", usecols='B:F')
            #sheet = pd.read_excel(path+'/'+self.partofspeech+'.xlsx', sheet_name="Nouns", usecols='B:I')
            sheet = sheet.values.tolist()
            unsorted_words = np.array(sheet)
            sorted_words=unsorted_words[unsorted_words[:, sentiment].argsort()]
            sorted_words=sorted_words.tolist()
            return sorted_words
    #print(read_partofspeech("Nouns"))


    # GROUPING DESCRIPTIVE WORDS
    def sort_sentiments(self, words):
        token_id=1
        grouped_words=[]

        # while(len(words)>0 and words[0][sentiment] is not None):
        while(len(words)>0 and words[0][sentiment] != 'nan'):
            words[0].append(token_id)
            words[0].append('')
            grouped_words.append(words[0])

            for j in range(1,len(words)):
                index=[]
                token = nlp(words[0][lemma]+' '+words[j][lemma])
                if int(token[0].similarity(token[1])*100) >= 40 and words[j][sentiment]=='nan':
                    # print(words[0][lemma], words[j][lemma], words[j][sentiment], token[0].similarity(token[1])*100)
                    words[j][sentiment]=words[0][sentiment]
                    words[j].append(token_id)
                    words[j].append(token[0].similarity(token[1])*100)
                    grouped_words.append(words[j])
                    print(words[j])
                    index.append(j)

            index.reverse()
            for i in index:
                words=np.delete(words, i, 0)
            words=np.delete(words, 0, 0)
            token_id=token_id+1

        return grouped_words


    #WRITE PoS
    def write_partofspeech(self, words):
        words = np.array(words)
        df_words = pd.DataFrame(words)
        df_words.columns=['frequency','text','lemma','pos','afinn sentiment','token_id','similarity']
        # print(df_words)

        #glossary to excel
        with pd.ExcelWriter(path+'/'+self.partofspeech+"sentiments.xlsx") as writer:
            df_words.to_excel(writer, sheet_name=self.partofspeech)
        writer.save()


Verbs=sentiments("Verbs")
Verbs.sorting_similar_words(Verbs.read_partofspeech())

Adverbs=sentiments("Adverbs")
Adverbs.sorting_similar_words(Adverbs.read_partofspeech())

Adjectives=sentiments("Adjectives")
Adjectives.sorting_similar_words(Adjectives.read_partofspeech())
