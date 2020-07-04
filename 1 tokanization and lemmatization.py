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
import json

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# OPENING JSON SENTIMENT DICTIONARY
with open('afinn-165.json') as f:
  items_afinn = json.load(f)

# OPENING LOUGHRAN MCDONALD SENTIMENT WORD LIST
file_path = 'LoughranMcDonald_SentimentWordLists_2018.xlsx'
items_mcdonals=[0]*8
for i in range(1,8):
    items_mcdonals[i-1]=pd.read_excel(file_path, sheet_name=i)
    items_mcdonals[i-1]=items_mcdonals[i-1].values.tolist()
    items_mcdonals[i-1] = [item.lower() for sublist in items_mcdonals[i-1] for item in sublist]
#print(items_mcdonals[2])

#CLEANING DATA
f_stop_words=open("StopWords_GenericLong.txt", "r")
stop_words=[str(i[0:-1]) for i in f_stop_words]
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']

#ABBREVIATIONS TO RETAIN
abbreviations=['co2','ch4','n2o','hfcs', 'pfcs', 'sf6', 'nf3', 'pop', 'voc', 'hap', 'pm', 'gwp', 'cfc11', 'ods', 'nox', 'so', 'mwh', 'kw', 'ir', 'odr', 'ldr', 'ar', 'ilo', 'oecd', 'who', 'lgbt', 'csr', "iccs", 'gst', 'ebitda','ifrs', 'iasb', 'ipsas', 'ifac', 'evgd', 'pl', 'eca']

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
        return sentences
    #print(tokenify_glossary(read_file()))


    #GLORRARY SRANDARDS
    def divide_glossary(self,sentences):
        glossary_nouns = []
        glossary_verbs = []
        glossary_adverbs = []
        glossary_adjectives = []
        POS=[]


        for sentence in sentences:
            doc = nlp(sentence)
            print(math.trunc(sentences.index(sentence)/len(sentences)*100))
            for token in doc:
                word=[]
                word.append(token.text)
                word.append(token.lemma_)
                word.append(token.pos_)
                # word.append(token.tag_)
                # word.append(token.dep_)

                #afinn sentiments
                if token.text in items_afinn:              #apoorva create a function for this later
                    word.append(items_afinn[token.text])
                else:
                    word.append('')

                #mcdonald sentiments
                # mc_rating=0
                # for i in range(7):                  #apoorva create a function for this later
                #     if token.text in items_mcdonals[i] or token.lemma_ in items_mcdonals[i]:
                #         mc_rating=i
                # if mc_rating != 0:
                #     word.append(mc_rating)
                # else:
                #     word.append('')

                #parts of speech segregation
                if token.pos_ == 'NOUN' or token.pos_ == 'PROPN':
                    glossary_nouns.append(word)
                elif token.pos_ == 'VERB':
                    glossary_verbs.append(word)
                elif token.pos_ == 'ADV' or token.pos_ == 'ADP':
                    glossary_adverbs.append(word)
                elif token.pos_ == 'ADJ':
                    glossary_adjectives.append(word)

            # 3D List
            POS.append(glossary_nouns)
            POS.append(glossary_verbs)
            POS.append(glossary_adverbs)
            POS.append(glossary_adjectives)

        return POS
    #print(divide_glossary(tokenify_glossary(read_file())))


    #GROUPING SYNONYMS
    # def remove_duplication_from_wordnet(self,keyword):
    #     synonym_list=[]
    #     synonym_list.append(keyword)
    #     for syn in wordnet.synsets(keyword):
    #         for synonym in syn.lemmas():
    #             if synonym.name() not in synonym_list:
    #                 synonym_list.append(synonym.name())
    #     return synonym_list
    #
    # def group_synoynms(self,POS):
    #     token_id=1
    #     for i in POS:
    #         synonym_list= self.remove_duplication_from_wordnet(i[1])
    #         for j in POS:
    #             if j[1] in synonym_list and j[8] is not None:
    #                 j[8] = token_id
    #         token_id = token_id +1
    #     return POS

    #RETAINING ABBREVIATIONS
    def is_abbreviation(self,word):
        if word in abbreviations:
            return 1
        else:
            return 0

    #REDUCE DUPLICATE WORDS AND FREQUENCY
    def reduce_glossary(self,sorted_words):
        glossary=[]
        # token_id=''
        while(len(sorted_words)>0):
            abb=0
            count=1
            word_frequency=[]
            while(len(sorted_words)>1 and sorted_words[0][1]==sorted_words[1][1]):
                count=count+1
                sorted_words=np.delete(sorted_words, 1, 0)
            if  self.is_abbreviation(sorted_words[0][0]) == 1:
                word_frequency.append(count)
                word_frequency.extend(sorted_words[0])
                glossary.append(word_frequency)
                abb=1
            if sorted_words[0][0] not in stop_words and len(sorted_words[0][0])>2 and sorted_words[0][0].isalpha() and abb==0:
                word_frequency.append(count)
                word_frequency.extend(sorted_words[0])
                # word_frequency.append(token_id)
                glossary.append(word_frequency)
            sorted_words=np.delete(sorted_words, 0, 0)
        return glossary


    #CREATE GLORRARY
    def sort_glossary(self,POS):
        sorted_POS=[]

        unsorted_nouns = np.array(POS[0])
        sorted_nouns=unsorted_nouns[unsorted_nouns[:, 1].argsort()]
        sorted_nouns=self.reduce_glossary(sorted_nouns)
        df_nouns = pd.DataFrame(sorted_nouns)
        # df_nouns.columns=['frequency','text','lemma','pos','eng-tag','dependency','afinn sentiment','mcdonals sentiment','token id']
        df_nouns.columns=['frequency','text','lemma','pos','afinn sentiment']

        unsorted_verbs = np.array(POS[1])
        sorted_verbs=unsorted_verbs[unsorted_verbs[:, 1].argsort()]
        sorted_verbs=self.reduce_glossary(sorted_verbs)
        df_verbs = pd.DataFrame(sorted_verbs)
        df_verbs.columns=['frequency','text','lemma','pos','afinn sentiment']

        unsorted_adverbs = np.array(POS[2])
        sorted_adverbs=unsorted_adverbs[unsorted_adverbs[:, 1].argsort()]
        sorted_adverbs=self.reduce_glossary(sorted_adverbs)
        df_adverbs = pd.DataFrame(sorted_adverbs)
        df_adverbs.columns=['frequency','text','lemma','pos','afinn sentiment']

        unsorted_adjectives = np.array(POS[3])
        sorted_adjectives=unsorted_adjectives[unsorted_adjectives[:, 1].argsort()]
        # sorted_adjectives=self.group_synoynms(self.reduce_glossary(sorted_adjectives))
        sorted_adjectives=self.reduce_glossary(sorted_adjectives)
        df_adjective = pd.DataFrame(sorted_adjectives)
        df_adjective.columns=['frequency','text','lemma','pos','afinn sentiment']

        sorted_POS.append(sorted_nouns)
        sorted_POS.append(sorted_verbs)
        sorted_POS.append(sorted_adverbs)
        sorted_POS.append(sorted_adjectives)

        #glossary to excel
        with pd.ExcelWriter("companies_glossary/"+self.company+".xlsx") as writer:
            df_nouns.to_excel(writer, sheet_name='Nouns')
            df_verbs.to_excel(writer, sheet_name='Verbs')
            df_adverbs.to_excel(writer, sheet_name='Adverbs')
            df_adjective.to_excel(writer, sheet_name='Adjectives')
        writer.save()

        return sorted_POS
    #print(sort_glossary(divide_glossary(tokenify_glossary(read_file()))))


# HUL = reports("HUL", "HUL 2018-2019_Annual Report.txt")
# print(HUL.sort_glossary(HUL.divide_glossary(HUL.tokenify_glossary(HUL.read_file()))))
#
# Colgate = reports("Colgate", "Colgate 2018-2019_Annual Report.txt")
# print(Colgate.sort_glossary(Colgate.divide_glossary(Colgate.tokenify_glossary(Colgate.read_file()))))
#
# ITC = reports("ITC", "ITC 2018-2019 Annual Report.txt")
# print(ITC.sort_glossary(ITC.divide_glossary(ITC.tokenify_glossary(ITC.read_file()))))
#
# Dabur = reports("Dabur", "Dabur 2018-19_Annual Report.txt")
# print(Dabur.sort_glossary(Dabur.divide_glossary(Dabur.tokenify_glossary(Dabur.read_file()))))
#
# Godrej = reports("Godrej", "Godrej 2018-2019_Annual Report.txt")
# print(Godrej.sort_glossary(Godrej.divide_glossary(Godrej.tokenify_glossary(Godrej.read_file()))))
#
# Marico = reports("Marico", "Marico 2018-2019_Annual Report.txt")
# print(Marico.sort_glossary(Marico.divide_glossary(Marico.tokenify_glossary(Marico.read_file()))))
#
# Nestle = reports("Nestle", "Nestle 2017-2018_Annual Report.txt")
# print(Nestle.sort_glossary(Nestle.divide_glossary(Nestle.tokenify_glossary(Nestle.read_file()))))
#
# PnG = reports("PnG", "P&G 2018-2019_Annual Report.txt")
# print(PnG.sort_glossary(PnG.divide_glossary(PnG.tokenify_glossary(PnG.read_file()))))


class glossary:
    def __init__(self, glossary, path):
        self.glossary = glossary
        self.filepath = path

    # READING COMPANY REPORTS
    def read_sheet(self):
        sheet=[]
        sheet= pd.read_excel(self.filepath)
        sheet = sheet.values.tolist()
        # sheet = [item.lower() for sublist in sheet for item in sublist]
        print(sheet)
        return sheet

    #GLORRARY SRANDARDS
    def divide_glossary(self,sheet):
        glossary_nouns = []
        glossary_verbs = []
        glossary_adverbs = []
        glossary_adjectives = []
        POS=[]

        for item in sheet:
            doc = nlp(item[2])
            for token in doc:
                word=[]
                word.append(item[0])
                word.append(item[1])
                word.append(token.text)
                word.append(token.lemma_)
                word.append(token.pos_)

                #parts of speech segregation
                if token.pos_ == 'NOUN' or token.pos_ == 'PROPN':
                    glossary_nouns.append(word)
                elif token.pos_ == 'VERB':
                    glossary_verbs.append(word)
                elif token.pos_ == 'ADV' or token.pos_ == 'ADP':
                    glossary_adverbs.append(word)
                elif token.pos_ == 'ADJ':
                    glossary_adjectives.append(word)

            # 3D List
            POS.append(glossary_nouns)
            POS.append(glossary_verbs)
            POS.append(glossary_adverbs)
            POS.append(glossary_adjectives)

        return POS

    #CREATE GLORRARY
    def create_glossary(self,POS):
        nouns = POS[0]
        df_nouns = pd.DataFrame(nouns)
        df_nouns.columns=['standard','sub-standard','text','lemma','pos']

        verbs = POS[1]
        df_verbs = pd.DataFrame(verbs)
        df_verbs.columns=['standard','sub-standard','text','lemma','pos']

        adverbs = POS[2]
        df_adverbs = pd.DataFrame(adverbs)
        df_adverbs.columns=['standard','sub-standard','text','lemma','pos']

        adjectives = POS[3]
        df_adjectives = pd.DataFrame(adjectives)
        df_adjectives.columns=['standard','sub-standard','text','lemma','pos']

        #glossary to excel
        with pd.ExcelWriter("companies_glossary/faulteredglossary.xlsx") as writer:
            df_nouns.to_excel(writer, sheet_name='Nouns')
            df_verbs.to_excel(writer, sheet_name='Verbs')
            df_adverbs.to_excel(writer, sheet_name='Adverbs')
            df_adjectives.to_excel(writer, sheet_name='Adjectives')
        writer.save()

Gl = glossary("Glossary", "consolidatedkeywords.xlsx")
print(Gl.create_glossary(Gl.divide_glossary(Gl.read_sheet())))
