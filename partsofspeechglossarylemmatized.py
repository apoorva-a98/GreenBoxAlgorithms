#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import math
import pandas as pd

# CREATING A EXCEL WORKBOOK
wb = Workbook(encoding='ascii')
sheet1 = wb.add_sheet('nouns')
sheet2 = wb.add_sheet('verbs')
sheet3 = wb.add_sheet('adverbs')
sheet4 = wb.add_sheet('adjectives')
sheet5 = wb.add_sheet('rest')

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

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
def tokenify_glossary(sentence):
    #S=[]
    #for sentence in sentences:
    buff = ''
    L=[]
    for letter in sentence:
        letter=letter.lower()
        if letter in avoid:
            if buff != '':
                L.append(buff)
            buff = ''
        elif (buff is not None):
            buff += letter
    if buff is not None:
        L.append(buff)
        buff=''
        #S.append(L)
    return L
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


#SENTIMENT SCORE
def clean_glossary(POS):
    sorted_POS=[]

    unsorted_nouns = np.array(POS[0])
    sorted_nouns=unsorted_nouns[unsorted_nouns[:, 1].argsort()]

    df_nouns = pd.DataFrame(sorted_nouns)
    filepath = 'Glossary.xlsx'
    df_nouns.to_excel(filepath, index=False)
    #sheet2.write(count, i, sentiment_count/len(sentence))

    unsorted_verbs = np.array(POS[1])
    sorted_verbs=unsorted_verbs[unsorted_verbs[:, 1].argsort()]

    unsorted_adverbs = np.array(POS[2])
    sorted_adverbs=unsorted_adverbs[unsorted_adverbs[:, 1].argsort()]

    unsorted_adjectives = np.array(POS[3])
    sorted_adjectives=unsorted_adjectives[unsorted_adjectives[:, 1].argsort()]

    sorted_POS.append(sorted_nouns)
    sorted_POS.append(sorted_verbs)
    sorted_POS.append(sorted_adverbs)
    sorted_POS.append(sorted_adjectives)
#     words.sort()
#     #print (words)
#     glossary = []
#     token_id =1
#
#     while(len(words)>0):
#         count=1
#         word_frequency=[]
#         while(len(words)>1 and words[0]==words[1]):
#             count=count+1
#             words.pop(1)
#         if words[0] not in stop_words and len(words[0])>2 and words[0].isalpha():
#             doc = nlp(words[0])
#             for token in doc:
#                 print(token.text, token.lemma_, token.pos_, token.tag_, token.dep_)
#             #word_frequency.append(words[0])
#             #print(words[0],lemmatizer.lemmatize(words[0]))
#             #word_frequency.append(count)
#             #word_frequency.append(token_id)
#             #token_id=token_id+1
#             glossary.append(word_frequency)
#         # else:
#         #     print(words[0])
#         words.pop(0)
#         count=1
    return sorted_POS
#
clean_glossary(divide_glossary(tokenify_glossary(items_HUL)))

#wb.save('Glossary.xls')


    #STOPWORDS FUNCTION*****

# def remove_stopwords(word,stopwords):
#     for i in words:
#         if i in stopwords:
#             index = words.index(i)
#             print(i,index)
#             words.pop(index)
#     return words
