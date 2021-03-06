#!/usr/bin/env python3

import spacy
#from nltk.stem import WordNetLemmatizer
import operator
import numpy as np

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")
#lemmatizer = WordNetLemmatizer()

#READING COMPANY REPORTS
f_HUL=open("HUL 2018-2019_Annual Report.txt", "r")
#items_HUL = f_HUL.read().split(".,")
#print(items_HUL)
items_HUL=f_HUL.read()
#print(items)

#CLEANING DATA
f_stop_words=open("StopWords_GenericLong.txt", "r")
stop_words=[str(i[0:-1]) for i in f_stop_words]
avoid=[' ','.','?','!','@','#','$','%','^','&','*','(',')','-','_','=','+','[',']','|',"\n","\t",';',':','<','>','/',',']
#print(stop_words)

#TOKANIZATION
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

SENTIMENT SCORE
def glossary_standards(words):
    words.sort()
    #print (words)
    glossary = []
    token_id =1

    while(len(words)>0):
        count=1
        word_frequency=[]
        while(len(words)>1 and words[0]==words[1]):
            count=count+1
            words.pop(1)
        if words[0] not in stop_words and len(words[0])>2 and words[0].isalpha():
            doc = nlp(words[0])
            for token in doc:
                print(token.text, token.lemma_, token.pos_, token.tag_, token.dep_)
            word_frequency.append(words[0])
            print(words[0],lemmatizer.lemmatize(words[0]))
            word_frequency.append(count)
            word_frequency.append(token_id)
            token_id=token_id+1
            glossary.append(word_frequency)
        # else:
        #     print(words[0])
        words.pop(0)
        count=1
    return glossary

glossary_standards(tokenify(items_HUL))


    #STOPWORDS FUNCTION*****

# def remove_stopwords(word,stopwords):
#     for i in words:
#         if i in stopwords:
#             index = words.index(i)
#             print(i,index)
#             words.pop(index)
#     return words


    # NOUN, VERB, ADJECTIVE, ADVERB*****

# from nltk.corpus import wordnet as wn
# words = ['amazing', 'interesting', 'love', 'great', 'nice', 'better', 'more', 'bad', 'badly', 'beauty', 'beautiful', 'beautifully']


    #NEGATION*****

# #for w in words:
#     #tmp = wn.synsets(w)[0].pos()
#     #print (w, ":", tmp)
#     #print(lemmatizer.lemmatize(w))
# print(lemmatizer.lemmatize(words[7], pos="r"))

# negate = ["aint", "arent", "cannot", "cant", "couldnt", "darent", "didnt", "doesnt", "ain't", "aren't", "can't",
#           "couldn't", "daren't", "didn't", "doesn't", "dont", "hadnt", "hasnt", "havent", "isnt", "mightnt", "mustnt",
#           "neither", "don't", "hadn't", "hasn't", "haven't", "isn't", "mightn't", "mustn't", "neednt", "needn't",
#           "never", "none", "nope", "nor", "not", "nothing", "nowhere", "oughtnt", "shant", "shouldnt", "wasnt",
#           "werent", "oughtn't", "shan't", "shouldn't", "wasn't", "weren't", "without", "wont", "wouldnt", "won't",
#           "wouldn't", "rarely", "seldom", "despite", "no", "nobody"]
#
#
# def negated(word):
#     """
#     Determine if preceding word is a negation word
#     """
#     if word.lower() in negate:
#         return True
#     else:
#         return False
#
#
# def tone_count_with_negation_check(dict, article):
#     """
#     Count positive and negative words with negation check. Account for simple negation only for positive words.
#     Simple negation is taken to be observations of one of negate words occurring within three words
#     preceding a positive words.
#     """
#     pos_count = 0
#     neg_count = 0
#
#     pos_words = []
#     neg_words = []
#
#     input_words = re.findall(r'\b([a-zA-Z]+n\'t|[a-zA-Z]+\'s|[a-zA-Z]+)\b', article.lower())
#
#     word_count = len(input_words)
#
#     for i in range(0, word_count):
#         if input_words[i] in dict['Negative']:
#             neg_count += 1
#             neg_words.append(input_words[i])
#         if input_words[i] in dict['Positive']:
#             if i >= 3:
#                 if negated(input_words[i - 1]) or negated(input_words[i - 2]) or negated(input_words[i - 3]):
#                     neg_count += 1
#                     neg_words.append(input_words[i] + ' (with negation)')
#                 else:
#                     pos_count += 1
#                     pos_words.append(input_words[i])
#             elif i == 2:
#                 if negated(input_words[i - 1]) or negated(input_words[i - 2]):
#                     neg_count += 1
#                     neg_words.append(input_words[i] + ' (with negation)')
#                 else:
#                     pos_count += 1
#                     pos_words.append(input_words[i])
#             elif i == 1:
#                 if negated(input_words[i - 1]):
#                     neg_count += 1
#                     neg_words.append(input_words[i] + ' (with negation)')
#                 else:
#                     pos_count += 1
#                     pos_words.append(input_words[i])
#             elif i == 0:
#                 pos_count += 1
#                 pos_words.append(input_words[i])
#
#     print('The results with negation check:', end='\n\n')
#     print('The # of positive words:', pos_count)
#     print('The # of negative words:', neg_count)
#     print('The list of found positive words:', pos_words)
#     print('The list of found negative words:', neg_words)
#     print('\n', end='')
#
#     results = [word_count, pos_count, neg_count, pos_words, neg_words]
#
#     return results
