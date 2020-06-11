#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import math
import pandas as pd
from pandas import DataFrame
import json

# CREATING A EXCEL WORKBOOK
#filepath = pd.ExcelWriter('Glossary.csv', engine='writer')
# wb = Workbook(encoding='ascii')
# sheet1 = wb.add_sheet('nouns')
# sheet2 = wb.add_sheet('verbs')
# sheet3 = wb.add_sheet('adverbs')
# sheet4 = wb.add_sheet('adjectives')
# sheet5 = wb.add_sheet('rest')

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# OPENING JSON SENTIMENT DICTIONARY
with open('afinn-165.json') as f:
  data = json.load(f)

# OPENING LOUGHRAN MCDONALD SENTIMENT WORD LIST
file_path = 'LoughranMcDonald_SentimentWordLists_2018.xlsx'
items_mcdonals=[0]*8
for i in range(1,8):
    items_mcdonals[i-1]=pd.read_excel(file_path, sheet_name=i)
    items_mcdonals[i-1]=items_mcdonals[i-1].values.tolist()
    items_mcdonals[i-1] = [item.lower() for sublist in items_mcdonals[i-1] for item in sublist]
#print(items_mcdonals[2])

#READING COMPANY REPORTS
f_HUL=open("HUL 2018-2019_Annual Report.txt", "r")
items_HUL=f_HUL.read()
#read_HUL = f_HUL.read().split(".")
#items_HUL=[str(i.lower()) for i in read_HUL]

#CLEANING DATA
f_stop_words=open("StopWords_GenericLong.txt", "r")
stop_words=[str(i[0:-1]) for i in f_stop_words]
#avoid=[' ','.','?','!','@','#','$','%','^','&','*','(',')','-','_','=','+','[',']','|','\n','\t',';',':','<','>','/',',']
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']


#TOKANIZATION
def tokenify_glossary(report):
    #S=[]
    #for sentence in sentences:
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
        #S.append(sentences)
    return sentences
#print(tokenify_glossary(items_HUL))


#GLOSSARY SRANDARDS
def divide_glossary(sentences):
    glossary_nouns = []
    glossary_verbs = []
    glossary_adverbs = []
    glossary_adjectives = []
    glossary_POS=[]
    POS=[]

    for sentence in sentences:
        doc = nlp(sentence)
        print(math.trunc(sentences.index(sentence)/len(sentences)*100))
        for token in doc:
            word=[]
            word.append(token.text)
            word.append(token.lemma_)
            word.append(token.pos_)
            word.append(token.tag_)
            word.append(token.dep_)

            #afinn sentiments
            if token.text in data:              #apoorva create a function for this later
                word.append(data[token.text])
            else:
                word.append('')

            #mcdonald sentiments
            mc_rating=0
            for i in range(7):                  #apoorva create a function for this later
                if token.text in items_mcdonals[i] or token.lemma_ in items_mcdonals[i]:
                    mc_rating=i
            if mc_rating != 0:
                word.append(mc_rating)
            else:
                word.append('')

            if token.pos_ == 'NOUN':
                glossary_nouns.append(word)
            elif token.pos_ == 'VERB':
                glossary_verbs.append(word)
            elif token.pos_ == 'ADV' or token.pos_ == 'ADP':
                glossary_adverbs.append(word)
            elif token.pos_ == 'ADJ':
                glossary_adjectives.append(word)
            else:
                glossary_POS.append(word)
        POS.append(glossary_nouns)
        POS.append(glossary_verbs)
        POS.append(glossary_adverbs)
        POS.append(glossary_adjectives)
    return POS
#print(divide_glossary(tokenify_glossary(items_HUL)))

#REDUCE DUPLICATE WORDS AND FREQUENCY
def reduce_glossary(sorted_words):
    glossary=[]
    token_id=1
    while(len(sorted_words)>0):
        count=1
        word_frequency=[]
        while(len(sorted_words)>1 and sorted_words[0][1]==sorted_words[1][1]):
            count=count+1
            sorted_words=np.delete(sorted_words, 1, 0)
        if sorted_words[0][0] not in stop_words and len(sorted_words[0][0])>2 and sorted_words[0][0].isalpha():
            word_frequency.append(count)
            word_frequency.extend(sorted_words[0])
            word_frequency.append(token_id)
            token_id=token_id+1
            glossary.append(word_frequency)
        sorted_words=np.delete(sorted_words, 0, 0)
        # #sorted_words.pop(0)
        count=1
    return glossary


#SENTIMENT SCORE
def sort_glossary(POS):
    sorted_POS=[]

    unsorted_nouns = np.array(POS[0])
    sorted_nouns=unsorted_nouns[unsorted_nouns[:, 1].argsort()]
    sorted_nouns=reduce_glossary(sorted_nouns)
    df_nouns = pd.DataFrame(sorted_nouns)

    unsorted_verbs = np.array(POS[1])
    sorted_verbs=unsorted_verbs[unsorted_verbs[:, 1].argsort()]
    sorted_verbs=reduce_glossary(sorted_verbs)
    df_verbs = pd.DataFrame(sorted_verbs)

    unsorted_adverbs = np.array(POS[2])
    sorted_adverbs=unsorted_adverbs[unsorted_adverbs[:, 1].argsort()]
    sorted_adverbs=reduce_glossary(sorted_adverbs)
    df_adverbs = pd.DataFrame(sorted_adverbs)

    unsorted_adjectives = np.array(POS[3])
    sorted_adjectives=unsorted_adjectives[unsorted_adjectives[:, 1].argsort()]
    sorted_adjectives=reduce_glossary(sorted_adjectives)
    df_adjective = pd.DataFrame(sorted_adjectives)

    sorted_POS.append(sorted_nouns)
    sorted_POS.append(sorted_verbs)
    sorted_POS.append(sorted_adverbs)
    sorted_POS.append(sorted_adjectives)

    with pd.ExcelWriter('Glossary.xls') as writer:
        df_nouns.to_excel(writer, sheet_name='Nouns')
        df_verbs.to_excel(writer, sheet_name='Verbs')
        df_adverbs.to_excel(writer, sheet_name='Adverbs')
        df_adjective.to_excel(writer, sheet_name='Adjectives')
    writer.save()

    return sorted_POS
sort_glossary(divide_glossary(tokenify_glossary(items_HUL)))
print(sort_glossary(divide_glossary(tokenify_glossary(items_HUL))))


#wb.save('Glossary.xls')


    #STOPWORDS FUNCTION*****

# def remove_stopwords(word,stopwords):
#     for i in words:
#         if i in stopwords:
#             index = words.index(i)
#             print(i,index)
#             words.pop(index)
#     return words
