#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame
import stanza
stanza.download('en' processors='tokenize,mwt,pos,lemma,depparse')
nlp = stanza.Pipeline('en')
doc = nlp("Barack Obama was born in Hawaii.  He was elected president in 2008.")
doc.sentences[0].print_dependencies()


# psudo code

I need to read reports.
1. fetch every sentence one by one.
2. find heads of each word.
3. if a child has a sentiment add to the head. repeat untill the root gets all the sentiments.
4. check for glossary words in subject object niuns and give sentiment to both subject and object.
5. if the subject/object of the next sentence is a pronoun, replace the pronoun with the last occured noun.
6. group noun phrases and and prepositional phrases for the polarity of the sentiment.

start with hul copy file.


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

    #IDENTIFY STANDARD

    #GIVE SENTIMENT


    #SINGLE SENTENCE READING
    def read_sentence(self,sentence):
        #psudo code
        array for word, index, head and SENTIMENTS
        identify the index of root.
            array len(sentence) -> how to call this?
            w.append(word[0])
            w.append(word.index)
            w.append(word.head)
            w.append(find_sentiment(word[0]))
            x.append(w)
            while root has child:
                if word has no child:
                    if word has a sentiment:
                        word.head.sentiment=word.head.sentiment +word.sentiment
                        x.pop(index(word))

        sentiment=[]*len(sentence)
        doc = nlp(sentence)

        for word in sentence:


    #print(read_sentence(tokenify_glossary(read_file())))

    #STANDARDS AND SENTIMENTS
    def read_document(self,sentences):
        for sentence in sentences:
            #if sentence has a standard:
                self.read_sentence(sentence)


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
