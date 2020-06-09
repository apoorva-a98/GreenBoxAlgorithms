#!/usr/bin/env python3

import spacy
import operator
import numpy as np

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

#READING COMPANY REPORTS
f_HUL=open("HUL 2018-2019_Annual Report.txt", "r")
#items_HUL=f_HUL.read()
read_HUL = f_HUL.read().split(".")
items_HUL=[str(i.lower()) for i in read_HUL]

#CLEANING DATA
f_stop_words=open("StopWords_GenericLong.txt", "r")
stop_words=[str(i[0:-1]) for i in f_stop_words]
#avoid=[' ','.','?','!','@','#','$','%','^','&','*','(',')','-','_','=','+','[',']','|','\n','\t',';',':','<','>','/',',']
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']


#TOKANIZATION
def tokenify(sentences):
    S=[]
    for sentence in sentences:
        buff = ''
        L=[]
        for letter in sentence:
            if letter in avoid:
                if buff != '':
                    L.append(buff)
                buff = ''
            elif (buff is not None):
                buff += letter
        if buff is not None:
            L.append(buff)
            buff=''
        S.append(L)
    return S
#tokenify(items_HUL)


#GLOSSARY SRANDARDS

def glossary_standards(sentences):
    glossary_nouns = []
    glossary_verbs = []
    glossary_adverbs = []
    glossary_adjectives = []
    glossary_POS=[]
    token_id =1

    for sentence in sentences:
        print(sentence)
        doc = nlp(str(sentence))
        for token in doc:
            print(token.text, token.lemma_, token.pos_, token.tag_, token.dep_)
            word=[]
            word.append(token.text)
            word.append(token.lemma_)
            word.append(token.pos_)
            word.append(token.tag_)
            word.append(token.dep_)
            if token.pos == 'NOUN':
                glossary_nouns.append(word)
            elif token.pos == 'VERB':
                glossary_nouns.append(word)
            elif token.pos == 'ADJ':
                glossary_nouns.append(word)
            elif token.pos == 'ADV':
                glossary_nouns.append(word)
            else:
                glossary_POS.append(word)
    return glossary_nouns, glossary_verbs, glossary_adverbs, glossary_adjectives, glossary_POS
glossary_standards(tokenify(items_HUL))


#SENTIMENT SCORE
# def glossary_standards(words):
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
#     return glossary
#
# glossary_standards(tokenify(items_HUL))


    #STOPWORDS FUNCTION*****

# def remove_stopwords(word,stopwords):
#     for i in words:
#         if i in stopwords:
#             index = words.index(i)
#             print(i,index)
#             words.pop(index)
#     return words
