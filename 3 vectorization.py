import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame
import json

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_md")


class standards_and_sentiments:
    def __init__(self, POS):
        self.partofspeech = POS

    #READ PoS
    def read_partofspeech(self):
            sheet = []
            sheet = pd.read_excel('companies_glossary/'+self.partofspeech+'.xlsx', sheet_name="Nouns", usecols='B:I')
            sheet = sheet.values.tolist()
            return sheet
    # print(read_partofspeech("Nouns"))


    # VECTORISING GLOSSARY
    def sorting_similar_words(self,words):
        i=0
        j=i+1
        token_id=1
        while(i != j):
            count=0
            for j in range(i+1,len(words),1):
                token = nlp(words[i][2]+' '+words[j][2])
                if int(token[0].similarity(token[1])*100) >= 50:
                    print(token[0].lemma_, token[1].lemma_, int(token[0].similarity(token[1])*100))
                    words[j].append(token_id)
                    place_holder=words[j]
                    words[j]=words[i+1+count]
                    words[i+1+count]=place_holder
                    count=count+1
            words[i].append(token_id)
            token_id=token_id+1
            i=i+count+1
        df_words = pd.DataFrame(words)
        df_words.columns=['frequency','text','lemma','pos','eng-tag','dependency','afinn sentiment','mcdonals sentiment','token_id']

        #glossary to excel
        with pd.ExcelWriter("companies_glossary/"+self.partofspeech+"vectorised.xlsx") as writer:
            df_words.to_excel(writer, sheet_name=self.partofspeech)
        writer.save()


Nouns=standards_and_sentiments("Nouns")
Nouns.sorting_similar_words(Nouns.read_partofspeech())

Verbs=standards_and_sentiments("Verbs")
Verbs.sorting_similar_words(Verbs.read_partofspeech())

Adverbs=standards_and_sentiments("Adverbs")
Adverbs.sorting_similar_words(Adverbs.read_partofspeech())

Adjectives=standards_and_sentiments("Adjectives")
Adjectives.sorting_similar_words(Adjectives.read_partofspeech())

ReportingRequirements=standards_and_sentiments("ReportingRequirements")
ReportingRequirements.sorting_similar_words(ReportingRequirements.read_partofspeech())
