#!/usr/bin/env python3

import tweepy
import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import pandas as pd
from pandas import DataFrame
import json
import re

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# OPENING JSON SENTIMENT DICTIONARY
with open('afinn-165.json') as f:
  items_afinn = json.load(f)

#CLEANING DATA
# f_stop_words=open("StopWords_GenericLong.txt", "r")
# stop_words=[str(i[0:-1]) for i in f_stop_words]
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']

# READING GLOSSARY EXCEL
def read_standards():
    sheet= pd.read_excel("Standards.xlsx")
    return sheet

# READING SENTIMENTS EXCEL
def read_sentiments():
    sheet= pd.read_excel("Sentiments.xlsx")
    return sheet

standards_data=read_standards()
sentiments_data=read_sentiments()

glossary=standards_data.values.tolist()
keywords=[]
for item in glossary:
    if item[2] not in keywords:
        keywords.append(item[2])
    if item[1] not in keywords:
        keywords.append(item[1])


# wb = Workbook(encoding='ascii')
# sheet1 = wb.add_sheet('Tweets')
# sheet2 = wb.add_sheet('Sentiment')

# Authenticate to Twitter
auth = tweepy.OAuthHandler("oGMSF0ZE60EPmq63CnJ5OVy56", "hoqrp3VgWWvzE3OjAT6iCD8upY4vc3UwcswH3Fk5Q4x46xllGK")
auth.set_access_token("847417106-WmqvtW8uv0HSHy0MXMpOjLEesVHoq450uXODLG9E", "ek7td0b383tqC75ZZ3bcuOqzl9ooxOrkpUAdjjHs5qp14")

# Create API object
api = tweepy.API(auth)


class reports:
    def __init__(self, name):
        self.company = name

    for j in range(3):
        sheet1.write(0, j, standards_title[j])
        count = 1
        for k in range(10):
            search_terms=[standards[j][k], "HUL"]
    #search_terms=["plastic HUL", "air HUL", "water HUL", "energy HUL", "materials HUL", "waste HUL", "habitat HUL", "wildlife HUL", "sustainable HUL", "emissions HUL", "plastic HUL"]
        #for i in range(len(search_terms)):
            public_tweets = api.search(q=search_terms, lang="en", count=100)      # q= “string that we are looking for”
            # foreach through all tweets pulled
            for tweet in public_tweets:
                #print(standards[j][k],count)
                sheet1.write(count, j, tweet.text)
                sheet1.write(count, j+3, str(tweet.created_at))
                #print (tweet.text)   #printing the text stored inside the tweet object
                print (tweet.created_at)   #printing the time
                sentence=tweet.text.split(' ')
                sentiment_count =0
                for l in sentence:
                    if l in data:
                        sentiment_count = sentiment_count + data[l]
                sheet2.write(count, j, sentiment_count/len(tweet.text))
                count=count+1
        print(count)


    #READ SENTENCES
    def read_sentences(self,sentences):
        master_list=pd.DataFrame(columns=['standards','sub-standard','sentence', 'sentiment'])
        for sentence in sentences:
            # sentence=re.split('but yet so',sentence)
            # sentence.split("but","yet","so")
            # for section in sentence:
            doc=nlp(sentence)
            try:
                root = [token for token in doc if token.head == token][0]
            except:
                continue
            root=[root]
            tree=[]
            stakeholders = 0
            materiality = self.select_standards(doc)
            if materiality:
                tree=self.create_tree(root,doc,tree)
                sentiment = self.calculate_sentiments(tree)
                for sub_standard in materiality:
                    # print(root,doc,sub_standard[0],sub_standard[1],sentence,sentiment)
                    data_point=[]
                    data_point.append(sub_standard[0])
                    data_point.append(sub_standard[1])
                    data_point.append(sentence)
                    data_point.append(sentiment)
                    to_append = data_point
                    df_length = len(master_list)
                    master_list.loc[df_length] = to_append
                    # print(data_point)
                    # df_data_point=pd.DataFrame(data_point)
                    # master_list=master_list.append(df_data_point)
        return master_list


    #LIST-TREE
    def create_tree(self,temp_head,doc,TREE):
        # doc = self.break_sentence(doc)
        # doc = self.merge_compounds(doc)
        # doc = self.combine_commas(doc)
        for sub_head in temp_head:
            self.track_subjects(sub_head)
            if not self.ignore_fillers(sub_head):
                # BRANCH =[]
                # BRANCH.append(sub_head.i)
                # BRANCH.append(sub_head.dep_)
                # BRANCH.append(sub_head.text)
                # BRANCH.append(sub_head.head.i)
                # TREE.append(BRANCH)
                TREE.append(sub_head)
            sub_tree=[child for child in sub_head.children]
            if not sub_tree:
                continue
            else:
                self.create_tree(sub_tree,doc,TREE,)
        return TREE

    #Track subjects and pronouns
    def track_subjects(self,sub_head):
        if sub_head.dep_=="nsubj" or sub_head.dep_=="csubj":
            subject=sub_head.text
        elif sub_head.dep_=="PRON" and (sub_head.text=="It" or sub_head.text=="it"):
            sub_head.text=subject
        elif sub_head.dep_=="PRON" and (sub_head.text!="It" or sub_head.text!="it"):
            stakeholders=stakeholders+1

    #Ignore dep
    def ignore_fillers(self,sub_head):
        if sub_head.dep_=="aux" or sub_head.pos_=="DET" or sub_head.pos_=="PUNCT" or sub_head.dep_=="preconj" or sub_head.dep_=="prep":
            return 1
        else:
            return 0

    # INDENTIFY MATERIALITY
    def select_standards(self,doc):
        materiality=[]
        for token in doc:
            found=df_standards.loc[df_standards['sub-standard'] == token].head(1).values.tolist()
            # try:
            if found:
                # materiality.append([found[0][0], found[0][1]])
                sub_standard=[]
                sub_standard.append(found[0][0])
                sub_standard.append(found[0][1])
                materiality.append(sub_standard)
                break
            else:
                found=df_standards.loc[df_standards['text'] == token.lemma_].head(1).values.tolist()
                if found:
                    # materiality.append([found[0][0], found[0][1]])
                    sub_standard=[]
                    sub_standard.append(found[0][0])
                    sub_standard.append(found[0][1])
                    materiality.append(sub_standard)
                else:
                    continue
            # except:
                # pass
        if not materiality:
            return 0
        return materiality

    #CALCULATE SENTIMENTS
    def calculate_sentiments(self,TREE):
        #add afinn
        count_descriptive_words=0
        sentiment=0
        for item in reversed(TREE):
            item_sentiment=0
            try:
                found=df_sentiments.loc[df_sentiments["keyword"] == item.lemma_].head(1).values.tolist()
                if found:
                    item_sentiment=item_sentiment+found[0][1]
                    count_descriptive_words=count_descriptive_words+1
            except TypeError:
                pass

            if not item_sentiment:
                item_sentiment=1
            if str(item)=="not" or str(item)=="nor":
                print(item)
                sentiment=item_sentiment+(-1)*sentiment
            else:
                sentiment=item_sentiment+sentiment
        # print(sentiment,count_descriptive_words)
        if count_descriptive_words:
            sentiment=sentiment/count_descriptive_words
        return sentiment


# f_glossary=open("glossary.txt", "r")
# fin = []
# soc = []
# env = []
# item_g = f_glossary.read().split("|")
# fin=item_g[0].split(",")
# soc=item_g[1].split(",")
# env=item_g[2].split(",")
# standards=[env,soc,fin]
# standards_title=['env','soc','fin']
#print(standards)
# Create a tweet
#api.update_status("Hello Tweepy")

# Using the API object to get tweets from your timeline, and storing it in a variable called public_tweets
#public_tweets = api.home_timeline()

# Using the API object to get tweets from a user’s timeline, and storing it in a variable called public_tweets
#public_tweets = api.user_timeline(id,count)     # max is 20 coz twitter put a limit to it
# for j in range(3):
#     sheet1.write(0, j, standards_title[j])
#     count = 1
#     for k in range(10):
#         search_terms=[standards[j][k], "HUL"]
# #search_terms=["plastic HUL", "air HUL", "water HUL", "energy HUL", "materials HUL", "waste HUL", "habitat HUL", "wildlife HUL", "sustainable HUL", "emissions HUL", "plastic HUL"]
#     #for i in range(len(search_terms)):
#         public_tweets = api.search(q=search_terms, lang="en", count=100)      # q= “string that we are looking for”
#         # foreach through all tweets pulled
#         for tweet in public_tweets:
#             #print(standards[j][k],count)
#             sheet1.write(count, j, tweet.text)
#             sheet1.write(count, j+3, str(tweet.created_at))
#             #print (tweet.text)   #printing the text stored inside the tweet object
#             print (tweet.created_at)   #printing the time
#             sentence=tweet.text.split(' ')
#             sentiment_count =0
#             for l in sentence:
#                 if l in data:
#                     sentiment_count = sentiment_count + data[l]
#             sheet2.write(count, j, sentiment_count/len(tweet.text))
#             count=count+1
#     print(count)
    #print (tweet.user.screen_name)   #printing the username
    #print (tweet.user.location)   #printing the location

# API reference index – Twitter developer tools
#https://developer.twitter.com/en/docs/api-reference-index

# Rate Limiting - Twitter
#https://developer.twitter.com/en/docs/basics/rate-limiting

# wb.save('Tweets.xls')
