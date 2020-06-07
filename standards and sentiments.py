#!/usr/bin/env python

import operator
import numpy as np


f_HUL=open("HUL 2018-2019_Annual Report.txt", "r")
f_stop_words=open("StopWords_GenericLong.txt", "r")
#items_HUL = f_HUL.read().split(".,")
#print(items_HUL)
items_HUL=f_HUL.read()
#print(items)

avoid=[' ','.','?','!','@','#','$','%','^','&','*','(',')','-','_','=','+','[',']','|',"\n",';',':','<','>','/',',' ]


def tokenify(sentences):
    sentences=sentences.lower()
    buff = ''
    words=[]

    for i in sentences:
        if i in avoid:
            if buff != '':
                words.append(buff)
            buff = ''
        elif (buff is not None):
            buff += i
    if buff is not None:
        words.append(buff)
        buff=''
    return words

#print (tokenify(items_HUL))

def count_list(terms):
    terms.sort()
    #print (terms)
    glossary = []

    while(len(terms)>0):
        count=1
        word_frequency=[]
        while(len(terms)>1 and terms[0]==terms[1]):
            count=count+1
            terms.pop(1)
        word_frequency.append(terms[0])
        word_frequency.append(count)
        terms.pop(0)
        count=1
        glossary.append(word_frequency)

    return glossary
    # A=[]
    # i=0
    # while(len(K)>0):
    #     count=1
    #     j=1
    #     while(j<len(K) and K[i]==K[j]):
    #         count=count+1
    #         K.pop(j)
    #         #print (A)
    #     #A[K[i]]=count
    #     #A.append(count)
    #     #B.append(K[i])
    #     #K.pop(i)
    #     #print(len(K))
    #     #print(K[i],K[j],count,limit)
    #     count=1

    #A = sorted(A.items(), key=operator.itemgetter(1), reverse=True)
    # return A
print(count_list(tokenify(items_HUL)))

#B=countList(tokenify(items))




# HUL_words = []
# for i in items_HUL:
#     HUL_words.append(i.split(" "))
#
# print(HUL_words)
